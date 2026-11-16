package server

import (
	"context"
	"fmt"
	"net/http"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"time"
)

// SetBuildInfo attaches build metadata exported on /metrics.
func (s *Server) SetBuildInfo(version, commit string) *Server {
	s.cfgMu.Lock()
	if strings.TrimSpace(version) != "" {
		s.buildVersion = strings.TrimSpace(version)
	}
	if strings.TrimSpace(commit) != "" {
		s.buildCommit = strings.TrimSpace(commit)
	}
	s.cfgMu.Unlock()
	return s
}

func (s *Server) recordHTTPMetric(method string, status int) {
	key := fmt.Sprintf("%s\x00%d", method, status)
	s.metricsMu.Lock()
	if s.httpReqs == nil {
		s.httpReqs = make(map[string]uint64)
	}
	s.httpReqs[key]++
	s.metricsMu.Unlock()
}

func (s *Server) recordBookingMetric(tenantSlug, action string, ok bool) {
	status := "error"
	if ok {
		status = "success"
	}
	key := fmt.Sprintf("%s\x00%s\x00%s", tenantSlug, action, status)
	s.metricsMu.Lock()
	if s.bookings == nil {
		s.bookings = make(map[string]uint64)
	}
	s.bookings[key]++
	s.metricsMu.Unlock()
}

func (s *Server) handleHealthz(w http.ResponseWriter, r *http.Request) {
	writeJSON(w, http.StatusOK, map[string]string{"status": "ok"})
}

func (s *Server) handleReadyz(w http.ResponseWriter, r *http.Request) {
	ctx, cancel := context.WithTimeout(r.Context(), 2*time.Second)
	defer cancel()

	if err := s.store.Ping(ctx); err != nil {
		writeJSON(w, http.StatusServiceUnavailable, map[string]string{
			"status":   "degraded",
			"database": "down",
			"error":    err.Error(),
		})
		return
	}

	tenants, err := s.store.ListTenants(ctx)
	if err != nil {
		writeJSON(w, http.StatusServiceUnavailable, map[string]string{
			"status":   "degraded",
			"database": "down",
			"error":    err.Error(),
		})
		return
	}

	writeJSON(w, http.StatusOK, map[string]any{
		"status":   "ok",
		"database": "up",
		"tenants":  len(tenants),
	})
}

func (s *Server) handleMetrics(w http.ResponseWriter, r *http.Request) {
	ctx, cancel := context.WithTimeout(r.Context(), 2*time.Second)
	defer cancel()

	s.cfgMu.RLock()
	ver := s.buildVersion
	commit := s.buildCommit
	started := s.startedAt
	s.cfgMu.RUnlock()

	dbUp := 0
	if err := s.store.Ping(ctx); err == nil {
		dbUp = 1
	}
	tenants, _ := s.store.ListTenants(ctx)

	var mem runtime.MemStats
	runtime.ReadMemStats(&mem)

	s.metricsMu.Lock()
	httpSnapshot := make(map[string]uint64, len(s.httpReqs))
	for k, v := range s.httpReqs {
		httpSnapshot[k] = v
	}
	bookingSnapshot := make(map[string]uint64, len(s.bookings))
	for k, v := range s.bookings {
		bookingSnapshot[k] = v
	}
	s.metricsMu.Unlock()

	var sb strings.Builder

	sb.WriteString("# HELP meeteasier_build_info MeetEasier build metadata.\n")
	sb.WriteString("# TYPE meeteasier_build_info gauge\n")
	fmt.Fprintf(&sb, "meeteasier_build_info{version=%q,commit=%q} 1\n", ver, commit)

	sb.WriteString("# HELP meeteasier_up Always 1 when the HTTP server is responding.\n")
	sb.WriteString("# TYPE meeteasier_up gauge\n")
	sb.WriteString("meeteasier_up 1\n")

	sb.WriteString("# HELP meeteasier_db_up 1 if SQLite database ping succeeds, 0 otherwise.\n")
	sb.WriteString("# TYPE meeteasier_db_up gauge\n")
	fmt.Fprintf(&sb, "meeteasier_db_up %d\n", dbUp)

	sb.WriteString("# HELP process_uptime_seconds Seconds since the server started.\n")
	sb.WriteString("# TYPE process_uptime_seconds gauge\n")
	fmt.Fprintf(&sb, "process_uptime_seconds %.3f\n", time.Since(started).Seconds())

	sb.WriteString("# HELP go_goroutines Number of goroutines that currently exist.\n")
	sb.WriteString("# TYPE go_goroutines gauge\n")
	fmt.Fprintf(&sb, "go_goroutines %d\n", runtime.NumGoroutine())

	sb.WriteString("# HELP go_memstats_alloc_bytes Number of bytes allocated and still in use.\n")
	sb.WriteString("# TYPE go_memstats_alloc_bytes gauge\n")
	fmt.Fprintf(&sb, "go_memstats_alloc_bytes %d\n", mem.Alloc)

	sb.WriteString("# HELP meeteasier_tenant_rooms_total Total cached rooms per tenant.\n")
	sb.WriteString("# TYPE meeteasier_tenant_rooms_total gauge\n")
	sb.WriteString("# HELP meeteasier_tenant_rooms_busy Number of currently busy rooms per tenant.\n")
	sb.WriteString("# TYPE meeteasier_tenant_rooms_busy gauge\n")
	sb.WriteString("# HELP meeteasier_tenant_rooms_error Number of rooms reporting calendar errors per tenant.\n")
	sb.WriteString("# TYPE meeteasier_tenant_rooms_error gauge\n")
	sb.WriteString("# HELP meeteasier_tenant_sync_stale 1 when tenant is serving stale cache due to sync failure.\n")
	sb.WriteString("# TYPE meeteasier_tenant_sync_stale gauge\n")
	sb.WriteString("# HELP meeteasier_tenant_last_sync_timestamp_seconds Unix timestamp of last successful tenant sync.\n")
	sb.WriteString("# TYPE meeteasier_tenant_last_sync_timestamp_seconds gauge\n")
	sb.WriteString("# HELP meeteasier_tenant_sync_duration_seconds Duration of the latest calendar sync attempt.\n")
	sb.WriteString("# TYPE meeteasier_tenant_sync_duration_seconds gauge\n")
	sb.WriteString("# HELP meeteasier_tenant_sync_total Total calendar sync attempts per tenant and status.\n")
	sb.WriteString("# TYPE meeteasier_tenant_sync_total counter\n")

	for _, t := range tenants {
		st := s.calMgr.GetStatus(t)
		busyCount := 0
		errCount := 0
		for _, rm := range st.Rooms {
			if rm.Busy {
				busyCount++
			}
			if rm.ErrorMessage != "" {
				errCount++
			}
		}
		staleVal := 0
		if st.IsStale {
			staleVal = 1
		}
		var lastSyncUnix int64
		if !st.LastSync.IsZero() {
			lastSyncUnix = st.LastSync.Unix()
		}
		fmt.Fprintf(&sb, "meeteasier_tenant_rooms_total{tenant=%q,provider=%q} %d\n", t.Slug, t.ProviderType, len(st.Rooms))
		fmt.Fprintf(&sb, "meeteasier_tenant_rooms_busy{tenant=%q} %d\n", t.Slug, busyCount)
		fmt.Fprintf(&sb, "meeteasier_tenant_rooms_error{tenant=%q} %d\n", t.Slug, errCount)
		fmt.Fprintf(&sb, "meeteasier_tenant_sync_stale{tenant=%q} %d\n", t.Slug, staleVal)
		fmt.Fprintf(&sb, "meeteasier_tenant_last_sync_timestamp_seconds{tenant=%q} %d\n", t.Slug, lastSyncUnix)
		fmt.Fprintf(&sb, "meeteasier_tenant_sync_duration_seconds{tenant=%q} %.6f\n", t.Slug, st.LastSyncDurationSeconds)
		fmt.Fprintf(&sb, "meeteasier_tenant_sync_total{tenant=%q,status=\"success\"} %d\n", t.Slug, st.SyncSuccessTotal)
		fmt.Fprintf(&sb, "meeteasier_tenant_sync_total{tenant=%q,status=\"error\"} %d\n", t.Slug, st.SyncErrorTotal)
	}

	sb.WriteString("# HELP meeteasier_room_bookings_total Total touch-screen/API room booking actions.\n")
	sb.WriteString("# TYPE meeteasier_room_bookings_total counter\n")
	bookingKeys := make([]string, 0, len(bookingSnapshot))
	for k := range bookingSnapshot {
		bookingKeys = append(bookingKeys, k)
	}
	sort.Strings(bookingKeys)
	for _, k := range bookingKeys {
		parts := strings.SplitN(k, "\x00", 3)
		if len(parts) != 3 {
			continue
		}
		fmt.Fprintf(&sb, "meeteasier_room_bookings_total{tenant=%q,action=%q,status=%q} %d\n", parts[0], parts[1], parts[2], bookingSnapshot[k])
	}

	sb.WriteString("# HELP meeteasier_http_requests_total Total HTTP requests processed by method and status code.\n")
	sb.WriteString("# TYPE meeteasier_http_requests_total counter\n")
	httpKeys := make([]string, 0, len(httpSnapshot))
	for k := range httpSnapshot {
		httpKeys = append(httpKeys, k)
	}
	sort.Strings(httpKeys)
	for _, k := range httpKeys {
		parts := strings.SplitN(k, "\x00", 2)
		if len(parts) != 2 {
			continue
		}
		statusCode, _ := strconv.Atoi(parts[1])
		fmt.Fprintf(&sb, "meeteasier_http_requests_total{method=%q,status=%q} %d\n", parts[0], strconv.Itoa(statusCode), httpSnapshot[k])
	}

	w.Header().Set("Content-Type", "text/plain; version=0.0.4; charset=utf-8")
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write([]byte(sb.String()))
}

