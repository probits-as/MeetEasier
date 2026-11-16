package calendar

import (
	"context"
	"fmt"
	"sort"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/probits-as/meeteasier/internal/model"
)

// MockProvider generates realistic conference room data relative to the current time
// and stores interactive mock bookings in memory so room booking can be tested offline.
type MockProvider struct {
	mu      sync.Mutex
	tenants map[int64]*mockTenantState
	seq     atomic.Int64
}

type mockTenantState struct {
	domain string
	rooms  []mockRuntimeRoom
}

type mockRuntimeRoom struct {
	Roomlist string
	Name     string
	Alias    string
	Email    string
	Events   []mockRuntimeEvent
}

type mockRuntimeEvent struct {
	ID          string
	Subject     string
	Organizer   string
	Sensitivity string
	Start       time.Time
	End         time.Time
}

// NewMockProvider creates a new MockProvider.
func NewMockProvider() *MockProvider {
	return &MockProvider{
		tenants: make(map[int64]*mockTenantState),
	}
}

func (m *MockProvider) FetchRooms(_ context.Context, tenant model.Tenant) ([]model.Room, error) {
	m.mu.Lock()
	defer m.mu.Unlock()

	now := time.Now()
	nowMs := now.UnixMilli()
	state := m.ensureTenantLocked(tenant, now)

	blacklist := tenant.BlacklistSet()
	maxItems := tenant.MaxItems
	if maxItems <= 0 {
		maxItems = 6
	}

	var rooms []model.Room
	for i := range state.rooms {
		r := &state.rooms[i]
		if _, blocked := blacklist[strings.ToLower(r.Email)]; blocked {
			continue
		}

		// Drop events that have already ended and sort chronologically
		active := make([]mockRuntimeEvent, 0, len(r.Events))
		for _, ev := range r.Events {
			if ev.End.UnixMilli() > nowMs {
				active = append(active, ev)
			}
		}
		sort.Slice(active, func(a, b int) bool {
			return active[a].Start.Before(active[b].Start)
		})
		r.Events = active

		appts := make([]model.Appointment, 0, len(active))
		busy := false
		for idx, ev := range active {
			if idx >= maxItems {
				break
			}
			startMs := ev.Start.UnixMilli()
			endMs := ev.End.UnixMilli()
			if idx == 0 && startMs <= nowMs && nowMs < endMs {
				busy = true
			}
			subject, isPrivate := applyPrivacy(tenant.PrivacyMode, ev.Subject, ev.Sensitivity, ev.Organizer)
			appts = append(appts, model.Appointment{
				ID:        ev.ID,
				Subject:   subject,
				Organizer: ev.Organizer,
				Start:     startMs,
				End:       endMs,
				Private:   isPrivate,
			})
		}

		rooms = append(rooms, model.Room{
			Roomlist:     r.Roomlist,
			Name:         r.Name,
			RoomAlias:    r.Alias,
			Email:        r.Email,
			Busy:         busy,
			Appointments: appts,
		})
	}

	sort.Slice(rooms, func(i, j int) bool {
		return strings.ToLower(rooms[i].Name) < strings.ToLower(rooms[j].Name)
	})

	return rooms, nil
}

// ExecuteBooking applies a booking, extension, or early end mutation to the in-memory mock schedule.
func (m *MockProvider) ExecuteBooking(_ context.Context, tenant model.Tenant, room model.Room, req model.BookingRequest) error {
	m.mu.Lock()
	defer m.mu.Unlock()

	now := time.Now().UTC()
	start, end, targetAppt, err := ComputeBookingWindow(room, req, now)
	if err != nil {
		return err
	}

	state := m.ensureTenantLocked(tenant, now)
	targetIdx := -1
	for i := range state.rooms {
		if strings.EqualFold(state.rooms[i].Alias, room.RoomAlias) || strings.EqualFold(state.rooms[i].Email, room.Email) {
			targetIdx = i
			break
		}
	}
	if targetIdx == -1 {
		return fmt.Errorf("fant ikke mock-rom %q", room.Name)
	}

	r := &state.rooms[targetIdx]
	subject := strings.TrimSpace(req.Subject)
	if subject == "" {
		subject = strings.TrimSpace(tenant.BookingDefaultSubject)
	}
	if subject == "" {
		subject = "Booked by MeetEasier"
	}

	switch req.Action {
	case model.BookingActionBookNow, model.BookingActionBookAfter:
		organizer := strings.TrimSpace(tenant.BookingOrganizerEmail)
		if organizer == "" {
			organizer = "Romskjerm"
		}
		id := fmt.Sprintf("mock-%d", m.seq.Add(1))
		r.Events = append(r.Events, mockRuntimeEvent{
			ID:          id,
			Subject:     subject,
			Organizer:   organizer,
			Sensitivity: "normal",
			Start:       start,
			End:         end,
		})

	case model.BookingActionExtend, model.BookingActionEndNow:
		updated := false
		for i := range r.Events {
			if (targetAppt.ID != "" && r.Events[i].ID == targetAppt.ID) ||
				(targetAppt.Start > 0 && r.Events[i].Start.UnixMilli() == targetAppt.Start) {
				r.Events[i].End = end
				updated = true
				break
			}
		}
		if !updated {
			return fmt.Errorf("fant ikke aktivt møte å oppdatere")
		}
	}

	return nil
}

func (m *MockProvider) ensureTenantLocked(tenant model.Tenant, now time.Time) *mockTenantState {
	domain := strings.TrimSpace(tenant.ExchangeDomain)
	if domain == "" {
		domain = "example.com"
	}

	if existing, ok := m.tenants[tenant.ID]; ok {
		if existing.domain != domain {
			existing.domain = domain
			for i := range existing.rooms {
				existing.rooms[i].Email = existing.rooms[i].Alias + "@" + domain
			}
		}
		return existing
	}

	base := now.Truncate(time.Minute)
	type mockSeedEvent struct {
		Subject     string
		Organizer   string
		Sensitivity string
		StartOffset time.Duration
		Duration    time.Duration
	}
	type mockSeedRoom struct {
		Roomlist string
		Name     string
		Events   []mockSeedEvent
	}

	seeds := []mockSeedRoom{
		{
			Roomlist: "1. Etasje",
			Name:     "Fjord",
			Events: []mockSeedEvent{
				{Subject: "Kvartalsgjennomgang Q4", Organizer: "Kari Nordmann", Sensitivity: "normal", StartOffset: -20 * time.Minute, Duration: 65 * time.Minute},
				{Subject: "Arkitektur & Sikkerhet", Organizer: "Ola Hansen", Sensitivity: "normal", StartOffset: 60 * time.Minute, Duration: 60 * time.Minute},
				{Subject: "1:1 Samtale", Organizer: "Ingrid Solberg", Sensitivity: "private", StartOffset: 150 * time.Minute, Duration: 30 * time.Minute},
			},
		},
		{
			Roomlist: "1. Etasje",
			Name:     "Varde",
			Events: []mockSeedEvent{
				{Subject: "Prosjektstatus Kundeportal", Organizer: "Per Berg", Sensitivity: "normal", StartOffset: 45 * time.Minute, Duration: 45 * time.Minute},
				{Subject: "Sprint Planning", Organizer: "Maren Lie", Sensitivity: "normal", StartOffset: 180 * time.Minute, Duration: 90 * time.Minute},
			},
		},
		{
			Roomlist: "1. Etasje",
			Name:     "Birkebeiner",
			Events: []mockSeedEvent{
				{Subject: "Styremøte", Organizer: "Håkon Dahl", Sensitivity: "private", StartOffset: -10 * time.Minute, Duration: 80 * time.Minute},
				{Subject: "Driftsmøte M365", Organizer: "Silje Bakken", Sensitivity: "normal", StartOffset: 90 * time.Minute, Duration: 45 * time.Minute},
			},
		},
		{
			Roomlist: "2. Etasje",
			Name:     "Nordlys",
			Events: []mockSeedEvent{
				{Subject: "Design Review", Organizer: "Jonas Eide", Sensitivity: "normal", StartOffset: 120 * time.Minute, Duration: 60 * time.Minute},
			},
		},
		{
			Roomlist: "2. Etasje",
			Name:     "Midnattssol",
			Events:   nil,
		},
		{
			Roomlist: "2. Etasje",
			Name:     "Lyngen",
			Events: []mockSeedEvent{
				{Subject: "Kundemøte", Organizer: "Astrid Holm", Sensitivity: "normal", StartOffset: -30 * time.Minute, Duration: 90 * time.Minute},
			},
		},
	}

	rooms := make([]mockRuntimeRoom, 0, len(seeds))
	for _, s := range seeds {
		alias := model.Slugify(s.Name)
		events := make([]mockRuntimeEvent, 0, len(s.Events))
		for _, ev := range s.Events {
			id := fmt.Sprintf("mock-seed-%d", m.seq.Add(1))
			events = append(events, mockRuntimeEvent{
				ID:          id,
				Subject:     ev.Subject,
				Organizer:   ev.Organizer,
				Sensitivity: ev.Sensitivity,
				Start:       base.Add(ev.StartOffset),
				End:         base.Add(ev.StartOffset + ev.Duration),
			})
		}
		rooms = append(rooms, mockRuntimeRoom{
			Roomlist: s.Roomlist,
			Name:     s.Name,
			Alias:    alias,
			Email:    alias + "@" + domain,
			Events:   events,
		})
	}

	state := &mockTenantState{
		domain: domain,
		rooms:  rooms,
	}
	m.tenants[tenant.ID] = state
	return state
}
