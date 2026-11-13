package calendar

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/probits-as/meeteasier/internal/model"
)

// MSGraphProvider fetches roomlists, rooms, and calendarView events from Microsoft Graph API.
type MSGraphProvider struct {
	httpClient   *http.Client
	graphBaseURL string
}

// NewMSGraphProvider creates a Microsoft Graph provider with a strict HTTP timeout.
func NewMSGraphProvider() *MSGraphProvider {
	return &MSGraphProvider{
		httpClient: &http.Client{
			Timeout: 10 * time.Second,
		},
		graphBaseURL: "https://graph.microsoft.com/v1.0",
	}
}

// NewMSGraphProviderWithClient creates a provider with custom HTTP client and base URL (used in tests).
func NewMSGraphProviderWithClient(client *http.Client, graphBaseURL string) *MSGraphProvider {
	if client == nil {
		client = &http.Client{Timeout: 10 * time.Second}
	}
	return &MSGraphProvider{
		httpClient:   client,
		graphBaseURL: strings.TrimRight(graphBaseURL, "/"),
	}
}

type tokenResponse struct {
	AccessToken      string `json:"access_token"`
	Error            string `json:"error"`
	ErrorDescription string `json:"error_description"`
}

type graphPlace struct {
	DisplayName  string `json:"displayName"`
	EmailAddress string `json:"emailAddress"`
}

type graphPlacesResponse struct {
	Value []graphPlace `json:"value"`
}

type graphEvent struct {
	Subject     string `json:"subject"`
	Sensitivity string `json:"sensitivity"`
	Organizer   struct {
		EmailAddress struct {
			Name    string `json:"name"`
			Address string `json:"address"`
		} `json:"emailAddress"`
	} `json:"organizer"`
	Start struct {
		DateTime string `json:"dateTime"`
		TimeZone string `json:"timeZone"`
	} `json:"start"`
	End struct {
		DateTime string `json:"dateTime"`
		TimeZone string `json:"timeZone"`
	} `json:"end"`
}

type graphCalendarViewResponse struct {
	Value []graphEvent `json:"value"`
}

// FetchRooms retrieves all rooms and their upcoming appointments for a tenant.
func (p *MSGraphProvider) FetchRooms(ctx context.Context, tenant model.Tenant) ([]model.Room, error) {
	if strings.TrimSpace(tenant.OAuthClientID) == "" ||
		strings.TrimSpace(tenant.OAuthAuthority) == "" ||
		strings.TrimSpace(tenant.OAuthClientSecret) == "" {
		return nil, errors.New("Microsoft Graph OAuth credentials are not configured")
	}

	token, err := p.acquireToken(ctx, tenant)
	if err != nil {
		return nil, fmt.Errorf("acquire oauth token: %w", err)
	}

	roomLists, err := p.getRoomLists(ctx, token, tenant.MaxRoomLists)
	if err != nil {
		return nil, fmt.Errorf("get room lists: %w", err)
	}

	blacklist := tenant.BlacklistSet()
	var rooms []model.Room

	for _, rl := range roomLists {
		if strings.TrimSpace(rl.EmailAddress) == "" {
			continue
		}
		listRooms, err := p.getRoomsInList(ctx, token, rl.EmailAddress, tenant.MaxRooms)
		if err != nil {
			return nil, fmt.Errorf("get rooms in list %q: %w", rl.DisplayName, err)
		}
		for _, item := range listRooms {
			email := strings.TrimSpace(item.EmailAddress)
			if _, blocked := blacklist[strings.ToLower(email)]; blocked {
				continue
			}
			rooms = append(rooms, model.Room{
				Roomlist:     rl.DisplayName,
				Name:         item.DisplayName,
				RoomAlias:    model.Slugify(item.DisplayName),
				Email:        email,
				Appointments: []model.Appointment{},
			})
		}
	}

	now := time.Now()
	startISO := now.UTC().Format(time.RFC3339)
	endISO := now.Add(time.Duration(tenant.MaxDays) * 24 * time.Hour).UTC().Format(time.RFC3339)

	// Fetch calendar views concurrently with bounded concurrency
	sem := make(chan struct{}, 6)
	var wg sync.WaitGroup
	for i := range rooms {
		wg.Add(1)
		go func(idx int) {
			defer wg.Done()
			select {
			case sem <- struct{}{}:
				defer func() { <-sem }()
			case <-ctx.Done():
				rooms[idx].ErrorMessage = ctx.Err().Error()
				return
			}

			events, err := p.getCalendarView(ctx, token, rooms[idx].Email, startISO, endISO, tenant.MaxItems)
			if err != nil {
				rooms[idx].ErrorMessage = err.Error()
				return
			}

			appts := make([]model.Appointment, 0, len(events))
			nowMs := time.Now().UnixMilli()
			for eIdx, ev := range events {
				startMs, errStart := parseGraphDateTime(ev.Start.DateTime, ev.Start.TimeZone)
				endMs, errEnd := parseGraphDateTime(ev.End.DateTime, ev.End.TimeZone)
				if errStart != nil || errEnd != nil {
					continue
				}

				if eIdx == 0 && startMs <= nowMs && nowMs < endMs {
					rooms[idx].Busy = true
				}

				organizer := ev.Organizer.EmailAddress.Name
				subject, isPrivate := applyPrivacy(tenant.PrivacyMode, ev.Subject, ev.Sensitivity, organizer)
				appts = append(appts, model.Appointment{
					Subject:   subject,
					Organizer: organizer,
					Start:     startMs,
					End:       endMs,
					Private:   isPrivate,
				})
			}
			rooms[idx].Appointments = appts
		}(i)
	}
	wg.Wait()

	sort.Slice(rooms, func(i, j int) bool {
		return strings.ToLower(rooms[i].Name) < strings.ToLower(rooms[j].Name)
	})

	return rooms, nil
}

func (p *MSGraphProvider) acquireToken(ctx context.Context, tenant model.Tenant) (string, error) {
	authority := strings.TrimRight(strings.TrimSpace(tenant.OAuthAuthority), "/")
	tokenURL := authority + "/oauth2/v2.0/token"

	form := url.Values{}
	form.Set("client_id", tenant.OAuthClientID)
	form.Set("client_secret", tenant.OAuthClientSecret)
	form.Set("scope", "https://graph.microsoft.com/.default")
	form.Set("grant_type", "client_credentials")

	req, err := http.NewRequestWithContext(ctx, http.MethodPost, tokenURL, strings.NewReader(form.Encode()))
	if err != nil {
		return "", err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")

	resp, err := p.httpClient.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(io.LimitReader(resp.Body, 1<<20))
	if err != nil {
		return "", err
	}

	var parsed tokenResponse
	if err := json.Unmarshal(body, &parsed); err != nil {
		if resp.StatusCode >= 400 {
			return "", fmt.Errorf("token endpoint status %d: %s", resp.StatusCode, strings.TrimSpace(string(body)))
		}
		return "", err
	}

	if resp.StatusCode >= 400 || parsed.AccessToken == "" {
		msg := parsed.ErrorDescription
		if msg == "" {
			msg = parsed.Error
		}
		if msg == "" {
			msg = fmt.Sprintf("http status %d", resp.StatusCode)
		}
		return "", errors.New(msg)
	}

	return parsed.AccessToken, nil
}

func (p *MSGraphProvider) getRoomLists(ctx context.Context, token string, top int) ([]graphPlace, error) {
	u := fmt.Sprintf("%s/places/microsoft.graph.roomlist?$select=displayName,emailAddress&$orderby=displayName&$top=%s",
		p.graphBaseURL, strconv.Itoa(top))
	var out graphPlacesResponse
	if err := p.getJSON(ctx, token, u, &out); err != nil {
		return nil, err
	}
	return out.Value, nil
}

func (p *MSGraphProvider) getRoomsInList(ctx context.Context, token, listEmail string, top int) ([]graphPlace, error) {
	u := fmt.Sprintf("%s/places/%s/microsoft.graph.roomlist/rooms?$select=displayName,emailAddress&$orderby=displayName&$top=%s",
		p.graphBaseURL, url.PathEscape(listEmail), strconv.Itoa(top))
	var out graphPlacesResponse
	if err := p.getJSON(ctx, token, u, &out); err != nil {
		return nil, err
	}
	return out.Value, nil
}

func (p *MSGraphProvider) getCalendarView(ctx context.Context, token, roomEmail, startISO, endISO string, top int) ([]graphEvent, error) {
	q := url.Values{}
	q.Set("startDateTime", startISO)
	q.Set("endDateTime", endISO)
	q.Set("$select", "organizer,subject,start,end,sensitivity")
	q.Set("$orderby", "start/dateTime")
	q.Set("$top", strconv.Itoa(top))

	u := fmt.Sprintf("%s/users/%s/calendar/calendarView?%s",
		p.graphBaseURL, url.PathEscape(roomEmail), q.Encode())

	var out graphCalendarViewResponse
	if err := p.getJSON(ctx, token, u, &out); err != nil {
		return nil, err
	}
	return out.Value, nil
}

func (p *MSGraphProvider) getJSON(ctx context.Context, token, targetURL string, dst any) error {
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, targetURL, nil)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Accept", "application/json")

	resp, err := p.httpClient.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(io.LimitReader(resp.Body, 4<<20))
	if err != nil {
		return err
	}

	if resp.StatusCode >= 400 {
		return fmt.Errorf("graph api %d: %s", resp.StatusCode, strings.TrimSpace(string(body)))
	}

	return json.Unmarshal(body, dst)
}

func parseGraphDateTime(raw, tz string) (int64, error) {
	raw = strings.TrimSpace(raw)
	if raw == "" {
		return 0, errors.New("empty datetime")
	}

	if t, err := time.Parse(time.RFC3339Nano, raw); err == nil {
		return t.UnixMilli(), nil
	}

	loc := time.UTC
	if tz != "" && !strings.EqualFold(tz, "UTC") {
		if loaded, err := time.LoadLocation(tz); err == nil {
			loc = loaded
		}
	}

	layouts := []string{
		"2006-01-02T15:04:05.9999999",
		"2006-01-02T15:04:05",
		"2006-01-02 15:04:05",
	}
	for _, layout := range layouts {
		if t, err := time.ParseInLocation(layout, raw, loc); err == nil {
			return t.UnixMilli(), nil
		}
	}

	return 0, fmt.Errorf("unsupported datetime format: %q", raw)
}

