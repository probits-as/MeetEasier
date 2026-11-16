package calendar

import (
	"context"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"sync/atomic"
	"testing"
	"time"

	"github.com/probits-as/meeteasier/internal/model"
)

func TestMSGraphProviderFetchAndPrivacyAndBlacklist(t *testing.T) {
	now := time.Now().UTC()
	start1 := now.Add(-15 * time.Minute).Format("2006-01-02T15:04:05.0000000")
	end1 := now.Add(45 * time.Minute).Format("2006-01-02T15:04:05.0000000")
	start2 := now.Add(90 * time.Minute).Format("2006-01-02T15:04:05.0000000")
	end2 := now.Add(150 * time.Minute).Format("2006-01-02T15:04:05.0000000")

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/oauth2/v2.0/token":
			_ = json.NewEncoder(w).Encode(map[string]string{"access_token": "fake-token"})
		case "/v1.0/places/microsoft.graph.roomlist":
			_ = json.NewEncoder(w).Encode(map[string]any{
				"value": []map[string]string{
					{"displayName": "Oslo HQ", "emailAddress": "oslo@example.com"},
				},
			})
		case "/v1.0/places/oslo@example.com/microsoft.graph.roomlist/rooms":
			_ = json.NewEncoder(w).Encode(map[string]any{
				"value": []map[string]string{
					{"displayName": "Fjord", "emailAddress": "fjord@example.com"},
					{"displayName": "Secret Room", "emailAddress": "secret@example.com"},
				},
			})
		case "/v1.0/users/fjord@example.com/calendar/calendarView":
			_ = json.NewEncoder(w).Encode(map[string]any{
				"value": []map[string]any{
					{
						"subject":     "Board Meeting",
						"sensitivity": "normal",
						"organizer":   map[string]any{"emailAddress": map[string]string{"name": "Kari Nordmann"}},
						"start":       map[string]string{"dateTime": start1, "timeZone": "UTC"},
						"end":         map[string]string{"dateTime": end1, "timeZone": "UTC"},
					},
					{
						"subject":     "Confidential Review",
						"sensitivity": "private",
						"organizer":   map[string]any{"emailAddress": map[string]string{"name": "Ola Hansen"}},
						"start":       map[string]string{"dateTime": start2, "timeZone": "UTC"},
						"end":         map[string]string{"dateTime": end2, "timeZone": "UTC"},
					},
				},
			})
		default:
			http.NotFound(w, r)
		}
	}))
	defer srv.Close()

	p := NewMSGraphProviderWithClient(srv.Client(), srv.URL+"/v1.0")
	tenant := model.Tenant{
		Slug:              "test",
		ProviderType:      model.ProviderMSGraph,
		OAuthClientID:     "cid",
		OAuthAuthority:    srv.URL,
		OAuthClientSecret: "sec",
		MaxDays:           10,
		MaxRoomLists:      10,
		MaxRooms:          10,
		MaxItems:          6,
		RoomBlacklist:     "secret@example.com",
		PrivacyMode:       model.PrivacyNormal,
	}

	rooms, err := p.FetchRooms(context.Background(), tenant)
	if err != nil {
		t.Fatalf("FetchRooms failed: %v", err)
	}
	if len(rooms) != 1 {
		t.Fatalf("expected 1 room after blacklist, got %d", len(rooms))
	}
	if !rooms[0].Busy {
		t.Fatalf("expected room to be busy")
	}
	if len(rooms[0].Appointments) != 2 {
		t.Fatalf("expected 2 appointments, got %d", len(rooms[0].Appointments))
	}
	if rooms[0].Appointments[0].Subject != "Board Meeting" {
		t.Fatalf("unexpected first subject: %q", rooms[0].Appointments[0].Subject)
	}
	if rooms[0].Appointments[1].Subject != "Private" || !rooms[0].Appointments[1].Private {
		t.Fatalf("expected second appointment to be masked as Private, got %+v", rooms[0].Appointments[1])
	}

	// Test organizer_as_subject mode
	tenant.PrivacyMode = model.PrivacyOrganizerAsSubject
	rooms, err = p.FetchRooms(context.Background(), tenant)
	if err != nil {
		t.Fatalf("FetchRooms organizer_as_subject mode failed: %v", err)
	}
	if rooms[0].Appointments[0].Subject != "Kari Nordmann" || !rooms[0].Appointments[0].Private {
		t.Fatalf("expected organizer as subject in organizer_as_subject mode, got %+v", rooms[0].Appointments[0])
	}
}

type staticTenantStore struct {
	tenants []model.Tenant
}

func (s staticTenantStore) ListTenants(context.Context) ([]model.Tenant, error) {
	return s.tenants, nil
}

func TestManagerStaleWhileRevalidateOnGraphOutage(t *testing.T) {
	var fail atomic.Bool
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if fail.Load() {
			w.WriteHeader(http.StatusServiceUnavailable)
			_, _ = w.Write([]byte("Microsoft service unavailable"))
			return
		}
		switch r.URL.Path {
		case "/oauth2/v2.0/token":
			_ = json.NewEncoder(w).Encode(map[string]string{"access_token": "fake-token"})
		case "/v1.0/places/microsoft.graph.roomlist":
			_ = json.NewEncoder(w).Encode(map[string]any{
				"value": []map[string]string{{"displayName": "HQ", "emailAddress": "hq@example.com"}},
			})
		case "/v1.0/places/hq@example.com/microsoft.graph.roomlist/rooms":
			_ = json.NewEncoder(w).Encode(map[string]any{
				"value": []map[string]string{{"displayName": "Fjord", "emailAddress": "fjord@example.com"}},
			})
		case "/v1.0/users/fjord@example.com/calendar/calendarView":
			_ = json.NewEncoder(w).Encode(map[string]any{"value": []any{}})
		}
	}))
	defer srv.Close()

	tenant := model.Tenant{
		ID:                  1,
		Slug:                "prod",
		ProviderType:        model.ProviderMSGraph,
		OAuthClientID:       "cid",
		OAuthAuthority:      srv.URL,
		OAuthClientSecret:   "sec",
		MaxDays:             10,
		MaxRoomLists:        10,
		MaxRooms:            10,
		MaxItems:            6,
		PollIntervalSeconds: 60,
	}

	mgr := NewManager(staticTenantStore{tenants: []model.Tenant{tenant}}, NewMSGraphProviderWithClient(srv.Client(), srv.URL+"/v1.0"), NewMockProvider())

	// First sync succeeds
	st := mgr.SyncTenantNow(context.Background(), tenant)
	if st.LastError != "" || st.IsStale || len(st.Rooms) != 1 {
		t.Fatalf("expected clean initial sync with 1 room, got %+v", st)
	}

	// Simulate Microsoft outage
	fail.Store(true)
	st2 := mgr.SyncTenantNow(context.Background(), tenant)
	if !st2.IsStale {
		t.Fatalf("expected IsStale=true during outage")
	}
	if st2.LastError == "" {
		t.Fatalf("expected LastError to be populated during outage")
	}
	if len(st2.Rooms) != 1 || st2.Rooms[0].Name != "Fjord" {
		t.Fatalf("expected cached room to be preserved during outage, got %+v", st2.Rooms)
	}
}

func TestMSGraphProviderExecuteBooking(t *testing.T) {
	var createdPath string
	var createdPayload map[string]any
	var patchedPath string

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		body, _ := io.ReadAll(r.Body)
		switch {
		case r.URL.Path == "/oauth2/v2.0/token":
			_ = json.NewEncoder(w).Encode(map[string]string{"access_token": "fake-token"})
		case r.Method == http.MethodPost && strings.HasSuffix(r.URL.Path, "/calendar/events"):
			createdPath = r.URL.Path
			_ = json.Unmarshal(body, &createdPayload)
			w.WriteHeader(http.StatusCreated)
			_ = json.NewEncoder(w).Encode(map[string]string{"id": "new-evt-1"})
		case r.Method == http.MethodPatch && strings.Contains(r.URL.Path, "/calendar/events/"):
			patchedPath = r.URL.Path
			w.WriteHeader(http.StatusOK)
			_ = json.NewEncoder(w).Encode(map[string]string{"id": "evt-1"})
		default:
			http.NotFound(w, r)
		}
	}))
	defer srv.Close()

	p := NewMSGraphProviderWithClient(srv.Client(), srv.URL+"/v1.0")
	tenant := model.Tenant{
		Slug:                  "prod",
		ProviderType:          model.ProviderMSGraph,
		OAuthClientID:         "cid",
		OAuthAuthority:        srv.URL,
		OAuthClientSecret:     "sec",
		BookingEnabled:        true,
		BookingOrganizerEmail: "kiosk@example.com",
		BookingDefaultSubject: "Hurtigmøte",
	}

	room := model.Room{
		Name:      "Fjord",
		RoomAlias: "fjord",
		Email:     "fjord@example.com",
	}

	// 1. book_now via BookingOrganizerEmail invites room as resource attendee
	err := p.ExecuteBooking(context.Background(), tenant, room, model.BookingRequest{
		RoomAlias:       "fjord",
		Action:          model.BookingActionBookNow,
		DurationMinutes: 30,
	})
	if err != nil {
		t.Fatalf("ExecuteBooking book_now failed: %v", err)
	}
	if createdPath != "/v1.0/users/kiosk@example.com/calendar/events" {
		t.Fatalf("expected POST to organizer calendar, got %s", createdPath)
	}
	if createdPayload["subject"] != "Hurtigmøte" {
		t.Fatalf("expected default subject Hurtigmøte, got %v", createdPayload["subject"])
	}

	// 2. extend active meeting patches room calendar event by ID
	nowMs := time.Now().UnixMilli()
	busyRoom := model.Room{
		Name:      "Fjord",
		RoomAlias: "fjord",
		Email:     "fjord@example.com",
		Busy:      true,
		Appointments: []model.Appointment{
			{
				ID:    "evt-active-1",
				Start: nowMs - 600_000,
				End:   nowMs + 600_000,
			},
			{
				ID:    "evt-next-2",
				Start: nowMs + 1_800_000, // 30m from now, only 20m gap after curr
				End:   nowMs + 3_600_000,
			},
		},
	}

	// Extending by 30m would collide with evt-next-2 -> must fail
	if err := p.ExecuteBooking(context.Background(), tenant, busyRoom, model.BookingRequest{
		Action:          model.BookingActionExtend,
		DurationMinutes: 30,
	}); err == nil {
		t.Fatalf("expected overlap error when extending past next meeting")
	}

	// Extending by 15m fits before evt-next-2 -> succeeds via PATCH
	if err := p.ExecuteBooking(context.Background(), tenant, busyRoom, model.BookingRequest{
		Action:          model.BookingActionExtend,
		DurationMinutes: 15,
	}); err != nil {
		t.Fatalf("ExecuteBooking extend 15m failed: %v", err)
	}
	if patchedPath != "/v1.0/users/fjord@example.com/calendar/events/evt-active-1" {
		t.Fatalf("expected PATCH to active room event, got %s", patchedPath)
	}

	// 3. end_now patches event end time to now
	patchedPath = ""
	if err := p.ExecuteBooking(context.Background(), tenant, busyRoom, model.BookingRequest{
		Action: model.BookingActionEndNow,
	}); err != nil {
		t.Fatalf("ExecuteBooking end_now failed: %v", err)
	}
	if patchedPath != "/v1.0/users/fjord@example.com/calendar/events/evt-active-1" {
		t.Fatalf("expected PATCH on end_now, got %s", patchedPath)
	}
}

