package calendar

import (
	"context"
	"sort"
	"strings"
	"time"

	"github.com/probits-as/meeteasier/internal/model"
)

// MockProvider generates realistic conference room data relative to the current time.
// Useful for previewing themes and when Microsoft services are unreachable.
type MockProvider struct{}

// NewMockProvider creates a new MockProvider.
func NewMockProvider() *MockProvider {
	return &MockProvider{}
}

type mockSeedRoom struct {
	Roomlist string
	Name     string
	Email    string
	BusyNow  bool
	Events   []mockSeedEvent
}

type mockSeedEvent struct {
	Subject     string
	Organizer   string
	Sensitivity string
	StartOffset time.Duration
	Duration    time.Duration
}

func (m *MockProvider) FetchRooms(_ context.Context, tenant model.Tenant) ([]model.Room, error) {
	now := time.Now().Truncate(time.Minute)
	domain := strings.TrimSpace(tenant.ExchangeDomain)
	if domain == "" {
		domain = "example.com"
	}

	seeds := []mockSeedRoom{
		{
			Roomlist: "1. Etasje",
			Name:     "Fjord",
			Email:    "fjord@" + domain,
			BusyNow:  true,
			Events: []mockSeedEvent{
				{Subject: "Kvartalsgjennomgang Q4", Organizer: "Kari Nordmann", Sensitivity: "normal", StartOffset: -20 * time.Minute, Duration: 65 * time.Minute},
				{Subject: "Arkitektur & Sikkerhet", Organizer: "Ola Hansen", Sensitivity: "normal", StartOffset: 60 * time.Minute, Duration: 60 * time.Minute},
				{Subject: "1:1 Samtale", Organizer: "Ingrid Solberg", Sensitivity: "private", StartOffset: 150 * time.Minute, Duration: 30 * time.Minute},
			},
		},
		{
			Roomlist: "1. Etasje",
			Name:     "Varde",
			Email:    "varde@" + domain,
			BusyNow:  false,
			Events: []mockSeedEvent{
				{Subject: "Prosjektstatus Kundeportal", Organizer: "Per Berg", Sensitivity: "normal", StartOffset: 45 * time.Minute, Duration: 45 * time.Minute},
				{Subject: "Sprint Planning", Organizer: "Maren Lie", Sensitivity: "normal", StartOffset: 180 * time.Minute, Duration: 90 * time.Minute},
			},
		},
		{
			Roomlist: "1. Etasje",
			Name:     "Birkebeiner",
			Email:    "birkebeiner@" + domain,
			BusyNow:  true,
			Events: []mockSeedEvent{
				{Subject: "Styremøte", Organizer: "Håkon Dahl", Sensitivity: "private", StartOffset: -10 * time.Minute, Duration: 80 * time.Minute},
				{Subject: "Driftsmøte M365", Organizer: "Silje Bakken", Sensitivity: "normal", StartOffset: 90 * time.Minute, Duration: 45 * time.Minute},
			},
		},
		{
			Roomlist: "2. Etasje",
			Name:     "Nordlys",
			Email:    "nordlys@" + domain,
			BusyNow:  false,
			Events: []mockSeedEvent{
				{Subject: "Design Review", Organizer: "Jonas Eide", Sensitivity: "normal", StartOffset: 120 * time.Minute, Duration: 60 * time.Minute},
			},
		},
		{
			Roomlist: "2. Etasje",
			Name:     "Midnattssol",
			Email:    "midnattssol@" + domain,
			BusyNow:  false,
			Events:   nil,
		},
		{
			Roomlist: "2. Etasje",
			Name:     "Lyngen",
			Email:    "lyngen@" + domain,
			BusyNow:  true,
			Events: []mockSeedEvent{
				{Subject: "Kundemøte", Organizer: "Astrid Holm", Sensitivity: "normal", StartOffset: -30 * time.Minute, Duration: 90 * time.Minute},
			},
		},
	}

	blacklist := tenant.BlacklistSet()
	maxItems := tenant.MaxItems
	if maxItems <= 0 {
		maxItems = 6
	}

	var rooms []model.Room
	for _, s := range seeds {
		if _, blocked := blacklist[strings.ToLower(s.Email)]; blocked {
			continue
		}

		appts := make([]model.Appointment, 0, len(s.Events))
		for i, ev := range s.Events {
			if i >= maxItems {
				break
			}
			subject, isPrivate := applyPrivacy(tenant.PrivacyMode, ev.Subject, ev.Sensitivity, ev.Organizer)
			start := now.Add(ev.StartOffset).UnixMilli()
			end := now.Add(ev.StartOffset + ev.Duration).UnixMilli()
			appts = append(appts, model.Appointment{
				Subject:   subject,
				Organizer: ev.Organizer,
				Start:     start,
				End:       end,
				Private:   isPrivate,
			})
		}

		rooms = append(rooms, model.Room{
			Roomlist:     s.Roomlist,
			Name:         s.Name,
			RoomAlias:    model.Slugify(s.Name),
			Email:        s.Email,
			Busy:         s.BusyNow,
			Appointments: appts,
		})
	}

	sort.Slice(rooms, func(i, j int) bool {
		return strings.ToLower(rooms[i].Name) < strings.ToLower(rooms[j].Name)
	})

	return rooms, nil
}

