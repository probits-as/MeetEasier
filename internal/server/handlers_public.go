package server

import (
	"crypto/subtle"
	"encoding/json"
	"errors"
	"fmt"
	"log/slog"
	"net/http"
	"strconv"
	"strings"

	"github.com/go-chi/chi/v5"

	"github.com/probits-as/meeteasier/internal/i18n"
	"github.com/probits-as/meeteasier/internal/model"
	"github.com/probits-as/meeteasier/internal/store"
	"github.com/probits-as/meeteasier/views/flightboard"
	"github.com/probits-as/meeteasier/views/layout"
	"github.com/probits-as/meeteasier/views/singleroom"
)

func (s *Server) enforceTenantIPAllowlist(w http.ResponseWriter, r *http.Request, tenant model.Tenant, tr i18n.Translations, isJSON bool) bool {
	if tenant.IsIPAllowed(r.RemoteAddr) {
		return true
	}
	clientIP := model.ExtractClientIP(r.RemoteAddr)
	slog.Warn("Blocked request from unauthorized IP",
		"tenant", tenant.Slug,
		"client_ip", clientIP,
		"path", r.URL.Path,
	)
	if isJSON {
		writeJSON(w, http.StatusForbidden, map[string]string{
			"error":     tr.IPBlockedMessage,
			"client_ip": clientIP,
		})
		return false
	}
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	w.WriteHeader(http.StatusForbidden)
	_ = layout.AccessDeniedPage(tenant, clientIP, tr).Render(r.Context(), w)
	return false
}

func (s *Server) handleFlightboardPage(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		slog.Error("Resolve tenant failed", "error", err)
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	tr := i18n.For(s.tenantLanguage(r, tenant))
	if !s.enforceTenantIPAllowlist(w, r, tenant, tr, false) {
		return
	}

	status := s.calMgr.GetStatus(tenant)
	activeFilter := r.URL.Query().Get("filter")

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := flightboard.FlightboardPage(tenant, status, activeFilter, explicitTenant, tr).Render(r.Context(), w); err != nil {
		slog.Error("Render flightboard page failed", "error", err)
	}
}

func (s *Server) handleFlightboardPartial(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	tr := i18n.For(s.tenantLanguage(r, tenant))
	if !s.enforceTenantIPAllowlist(w, r, tenant, tr, false) {
		return
	}

	status := s.calMgr.GetStatus(tenant)
	activeFilter := r.URL.Query().Get("filter")

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := flightboard.FlightboardRows(tenant, status, activeFilter, explicitTenant, tr).Render(r.Context(), w); err != nil {
		slog.Error("Render flightboard partial failed", "error", err)
	}
}

func (s *Server) handleSingleRoomPage(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	tr := i18n.For(s.tenantLanguage(r, tenant))
	if !s.enforceTenantIPAllowlist(w, r, tenant, tr, false) {
		return
	}

	alias := chi.URLParam(r, "alias")
	status := s.calMgr.GetStatus(tenant)
	room, found := s.calMgr.GetRoomByAlias(tenant, alias)

	if !found && status.Initialized {
		w.WriteHeader(http.StatusNotFound)
	}
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := singleroom.SingleRoomPage(tenant, room, alias, found, status.Initialized, explicitTenant, tr, "", "").Render(r.Context(), w); err != nil {
		slog.Error("Render single room page failed", "error", err)
	}
}

func (s *Server) handleSingleRoomPartial(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	tr := i18n.For(s.tenantLanguage(r, tenant))
	if !s.enforceTenantIPAllowlist(w, r, tenant, tr, false) {
		return
	}

	alias := chi.URLParam(r, "alias")
	status := s.calMgr.GetStatus(tenant)
	room, found := s.calMgr.GetRoomByAlias(tenant, alias)

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := singleroom.SingleRoomContent(tenant, room, found, status.Initialized, explicitTenant, tr, "", "").Render(r.Context(), w); err != nil {
		slog.Error("Render single room partial failed", "error", err)
	}
}

func (s *Server) handleSingleRoomBook(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	alias := chi.URLParam(r, "alias")
	tr := i18n.For(s.tenantLanguage(r, tenant))
	if !s.enforceTenantIPAllowlist(w, r, tenant, tr, false) {
		return
	}
	if err := r.ParseForm(); err != nil {
		room, found := s.calMgr.GetRoomByAlias(tenant, alias)
		s.renderSingleRoomBookResponse(w, r, tenant, room, alias, found, explicitTenant, tr, "", tr.ErrInvalidRequest, http.StatusBadRequest)
		return
	}

	action := strings.TrimSpace(r.FormValue("action"))
	durationMinutes, _ := strconv.Atoi(strings.TrimSpace(r.FormValue("duration_minutes")))
	if choice := strings.TrimSpace(r.FormValue("booking_choice")); choice != "" {
		if parts := strings.SplitN(choice, ":", 2); len(parts) == 2 {
			action = strings.TrimSpace(parts[0])
			durationMinutes, _ = strconv.Atoi(strings.TrimSpace(parts[1]))
		}
	}

	req := model.BookingRequest{
		RoomAlias:       alias,
		Action:          action,
		DurationMinutes: durationMinutes,
		Subject:         strings.TrimSpace(r.FormValue("subject")),
		PIN:             strings.TrimSpace(r.FormValue("pin")),
	}

	room, statusCode, execErr := s.validateAndExecuteBooking(r, tenant, alias, req, tr)
	if execErr != nil {
		_, found := s.calMgr.GetRoomByAlias(tenant, alias)
		s.renderSingleRoomBookResponse(w, r, tenant, room, alias, found || room.RoomAlias != "", explicitTenant, tr, "", execErr.Error(), statusCode)
		return
	}

	flash := bookingSuccessMessage(tr, req)
	s.renderSingleRoomBookResponse(w, r, tenant, room, alias, true, explicitTenant, tr, flash, "", http.StatusOK)
}

func (s *Server) renderSingleRoomBookResponse(w http.ResponseWriter, r *http.Request, tenant model.Tenant, room model.Room, alias string, found, explicitTenant bool, tr i18n.Translations, flashMsg, errMsg string, statusCode int) {
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	isHTMX := strings.EqualFold(r.Header.Get("HX-Request"), "true")
	if isHTMX {
		// Return 200 for HTMX partial swaps so error banners are rendered inline on the touch panel.
		w.WriteHeader(http.StatusOK)
		if err := singleroom.SingleRoomContent(tenant, room, found, true, explicitTenant, tr, flashMsg, errMsg).Render(r.Context(), w); err != nil {
			slog.Error("Render single room booking response failed", "error", err)
		}
		return
	}

	if statusCode >= 400 {
		w.WriteHeader(statusCode)
	}
	if err := singleroom.SingleRoomPage(tenant, room, alias, found, true, explicitTenant, tr, flashMsg, errMsg).Render(r.Context(), w); err != nil {
		slog.Error("Render single room booking page failed", "error", err)
	}
}

func (s *Server) handleAPIRoomBooking(w http.ResponseWriter, r *http.Request) {
	tenant, _, err := s.resolveTenant(r)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": "Failed to resolve tenant"})
		return
	}

	r.Body = http.MaxBytesReader(w, r.Body, 1<<20)
	var req model.BookingRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": "Invalid JSON: " + err.Error()})
		return
	}
	if strings.TrimSpace(req.RoomAlias) == "" {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": "room_alias is required"})
		return
	}

	tr := i18n.For(s.tenantLanguage(r, tenant))
	if !s.enforceTenantIPAllowlist(w, r, tenant, tr, true) {
		return
	}
	room, statusCode, execErr := s.validateAndExecuteBooking(r, tenant, req.RoomAlias, req, tr)
	if execErr != nil {
		writeJSON(w, statusCode, map[string]string{"error": execErr.Error()})
		return
	}

	writeJSON(w, http.StatusOK, map[string]any{
		"status":  "ok",
		"message": bookingSuccessMessage(tr, req),
		"room":    room,
	})
}

func (s *Server) validateAndExecuteBooking(r *http.Request, tenant model.Tenant, alias string, req model.BookingRequest, tr i18n.Translations) (model.Room, int, error) {
	if !tenant.BookingEnabled {
		room, _ := s.calMgr.GetRoomByAlias(tenant, alias)
		return room, http.StatusForbidden, errors.New(tr.BookingDisabledErr)
	}

	if tenant.HasBookingPIN {
		submitted := strings.TrimSpace(req.PIN)
		if submitted == "" || subtle.ConstantTimeCompare([]byte(submitted), []byte(tenant.BookingPIN)) != 1 {
			room, _ := s.calMgr.GetRoomByAlias(tenant, alias)
			return room, http.StatusUnauthorized, errors.New(tr.BookingInvalidPINErr)
		}
	}

	if strings.TrimSpace(req.Subject) == "" && strings.TrimSpace(tenant.BookingDefaultSubject) == "" {
		req.Subject = tr.DefaultBookedSubject
	}

	room, err := s.calMgr.BookRoom(r.Context(), tenant, alias, req)
	s.recordBookingMetric(tenant.Slug, req.Action, err == nil)
	if err != nil {
		slog.Warn("Room booking failed", "tenant", tenant.Slug, "room", alias, "action", req.Action, "error", err)
		return room, http.StatusBadRequest, err
	}

	slog.Info("Room booking executed", "tenant", tenant.Slug, "room", alias, "action", req.Action, "duration_minutes", req.DurationMinutes)
	return room, http.StatusOK, nil
}

func bookingSuccessMessage(tr i18n.Translations, req model.BookingRequest) string {
	switch req.Action {
	case model.BookingActionBookNow:
		return fmt.Sprintf(tr.FlashBookedNowFmt, req.DurationMinutes)
	case model.BookingActionBookAfter:
		return fmt.Sprintf(tr.FlashBookedAfterFmt, req.DurationMinutes)
	case model.BookingActionExtend:
		return fmt.Sprintf(tr.FlashExtendedFmt, req.DurationMinutes)
	case model.BookingActionEndNow:
		return tr.FlashEndedNow
	default:
		return tr.FlashBookingUpdated
	}
}

func (s *Server) handleTenantAsset(w http.ResponseWriter, r *http.Request) {
	slug := chi.URLParam(r, "slug")
	key := chi.URLParam(r, "key")

	asset, err := s.store.GetAsset(r.Context(), slug, key)
	if errors.Is(err, store.ErrNotFound) {
		http.NotFound(w, r)
		return
	}
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Type", asset.ContentType)
	w.Header().Set("Cache-Control", "public, max-age=3600")
	_, _ = w.Write(asset.Data)
}

func (s *Server) handleAPIRooms(w http.ResponseWriter, r *http.Request) {
	tenant, _, err := s.resolveTenant(r)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": "Failed to resolve tenant"})
		return
	}
	if !s.enforceTenantIPAllowlist(w, r, tenant, i18n.For(s.tenantLanguage(r, tenant)), true) {
		return
	}
	status := s.calMgr.GetStatus(tenant)
	if status.LastError != "" && len(status.Rooms) == 0 {
		writeJSON(w, http.StatusBadGateway, map[string]string{"error": status.LastError})
		return
	}
	writeJSON(w, http.StatusOK, status.Rooms)
}

func (s *Server) handleAPIRoomlists(w http.ResponseWriter, r *http.Request) {
	tenant, _, err := s.resolveTenant(r)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": "Failed to resolve tenant"})
		return
	}
	if !s.enforceTenantIPAllowlist(w, r, tenant, i18n.For(s.tenantLanguage(r, tenant)), true) {
		return
	}
	status := s.calMgr.GetStatus(tenant)
	if status.LastError != "" && len(status.Roomlists) == 0 {
		writeJSON(w, http.StatusBadGateway, map[string]string{"error": status.LastError})
		return
	}
	writeJSON(w, http.StatusOK, status.Roomlists)
}

func (s *Server) handleAPIHeartbeat(w http.ResponseWriter, r *http.Request) {
	writeJSON(w, http.StatusOK, map[string]string{"status": "OK"})
}

func writeJSON(w http.ResponseWriter, statusCode int, payload any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(statusCode)
	_ = json.NewEncoder(w).Encode(payload)
}

