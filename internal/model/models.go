package model

import (
	"encoding/json"
	"strings"
	"time"
)

const (
	ProviderMock    = "mock"
	ProviderMSGraph = "msgraph"

	PrivacyNormal             = "normal"
	PrivacyOrganizerAsSubject = "organizer_as_subject"
	PrivacyHideSubject        = "hide_subject"

	PresetStandard = "standard"
	PresetCompact  = "compact"
	PresetCustom   = "custom"

	RowLayoutStandard = "standard"
	RowLayoutCompact  = "compact"
)

// ThemeLabels holds all configurable UI strings for a tenant.
type ThemeLabels struct {
	NavbarTitle     string `json:"navbar_title"`
	FilterTitle     string `json:"filter_title"`
	FilterAllTitle  string `json:"filter_all_title"`
	StatusAvailable string `json:"status_available"`
	StatusBusy      string `json:"status_busy"`
	StatusError     string `json:"status_error"`
	NextUp          string `json:"next_up"`
	NowPrefix       string `json:"now_prefix"`
	UpcomingTitle   string `json:"upcoming_title"`
}

// ThemeConfig holds styling, layout, branding assets, and labels for a tenant.
type ThemeConfig struct {
	RowLayout                string      `json:"row_layout"`
	LogoURL                  string      `json:"logo_url"`
	FooterImageURL           string      `json:"footer_image_url"`
	SingleRoomLogoURL        string      `json:"single_room_logo_url"`
	SingleRoomStatusFontSize string      `json:"single_room_status_font_size"`
	ShowWeekdayInTime        bool        `json:"show_weekday_in_time"`
	BodyBg                   string      `json:"body_bg"`
	CardBg                   string      `json:"card_bg"`
	CardBorder               string      `json:"card_border"`
	TitleBarBg               string      `json:"title_bar_bg"`
	PrimaryFont              string      `json:"primary_font"`
	OpenColor                string      `json:"open_color"`
	OpenDarkColor            string      `json:"open_dark_color"`
	BusyColor                string      `json:"busy_color"`
	UpcomingColor            string      `json:"upcoming_color"`
	ErrorColor               string      `json:"error_color"`
	Labels                   ThemeLabels `json:"labels"`
}

// DefaultStandardTheme returns the default MeetEasier theme configuration.
func DefaultStandardTheme() ThemeConfig {
	return ThemeConfig{
		RowLayout:                RowLayoutStandard,
		LogoURL:                  "/img/logo.png",
		FooterImageURL:           "",
		SingleRoomLogoURL:        "",
		SingleRoomStatusFontSize: "18vw",
		ShowWeekdayInTime:        false,
		BodyBg:                   "#191919",
		CardBg:                   "#333333",
		CardBorder:               "#474747",
		TitleBarBg:               "#333333",
		PrimaryFont:              "#ffffff",
		OpenColor:                "#3adb76",
		OpenDarkColor:            "#289952",
		BusyColor:                "#f04124",
		UpcomingColor:            "#ffae00",
		ErrorColor:               "#f3840e",
		Labels: ThemeLabels{
			NavbarTitle:     "Conference Room Availability",
			FilterTitle:     "Locations",
			FilterAllTitle:  "All Conference Rooms",
			StatusAvailable: "Open",
			StatusBusy:      "Busy",
			StatusError:     "Error",
			NextUp:          "Next Up",
			NowPrefix:       "",
			UpcomingTitle:   "Upcoming",
		},
	}
}

// DefaultCompactTheme returns a compact, responsive, two-column theme preset (Norwegian labels).
func DefaultCompactTheme() ThemeConfig {
	return ThemeConfig{
		RowLayout:                RowLayoutCompact,
		LogoURL:                  "/img/logo.png",
		FooterImageURL:           "",
		SingleRoomLogoURL:        "",
		SingleRoomStatusFontSize: "10vw",
		ShowWeekdayInTime:        true,
		BodyBg:                   "#191919",
		CardBg:                   "#333333",
		CardBorder:               "#474747",
		TitleBarBg:               "#333333",
		PrimaryFont:              "#ffffff",
		OpenColor:                "#3adb76",
		OpenDarkColor:            "#289952",
		BusyColor:                "#f04124",
		UpcomingColor:            "#ffae00",
		ErrorColor:               "#f3840e",
		Labels: ThemeLabels{
			NavbarTitle:     "Møterom",
			FilterTitle:     "Lokasjoner",
			FilterAllTitle:  "Alle rom",
			StatusAvailable: "Ledig",
			StatusBusy:      "Opptatt",
			StatusError:     "Feil",
			NextUp:          "Neste",
			NowPrefix:       "Nå",
			UpcomingTitle:   "Kommer",
		},
	}
}

// ThemeForPreset returns a copy of the preset theme or falls back to Standard.
func ThemeForPreset(preset string) ThemeConfig {
	switch strings.ToLower(strings.TrimSpace(preset)) {
	case PresetCompact:
		return DefaultCompactTheme()
	default:
		return DefaultStandardTheme()
	}
}

// ThemeAssetData represents a base64-encoded image bundled inside a ThemeBundle JSON.
type ThemeAssetData struct {
	ContentType string `json:"content_type"`
	DataBase64  string `json:"data_base64"`
}

// ThemeBundle is a portable JSON package for exporting and importing tenant themes and assets.
type ThemeBundle struct {
	Version             int                       `json:"version"`
	Slug                string                    `json:"slug,omitempty"`
	Name                string                    `json:"name,omitempty"`
	Domains             string                    `json:"domains,omitempty"`
	IsDefault           *bool                     `json:"is_default,omitempty"`
	ProviderType        string                    `json:"provider_type,omitempty"`
	OAuthClientID       string                    `json:"oauth_client_id,omitempty"`
	OAuthAuthority      string                    `json:"oauth_authority,omitempty"`
	ExchangeDomain      string                    `json:"exchange_domain,omitempty"`
	MaxDays             int                       `json:"max_days,omitempty"`
	MaxRoomLists        int                       `json:"max_room_lists,omitempty"`
	MaxRooms            int                       `json:"max_rooms,omitempty"`
	MaxItems            int                       `json:"max_items,omitempty"`
	RoomBlacklist       string                    `json:"room_blacklist,omitempty"`
	PollIntervalSeconds int                       `json:"poll_interval_seconds,omitempty"`
	PrivacyMode         string                    `json:"privacy_mode,omitempty"`
	ThemePreset         string                    `json:"theme_preset,omitempty"`
	Theme               ThemeConfig               `json:"theme"`
	CustomCSS           string                    `json:"custom_css,omitempty"`
	Assets              map[string]ThemeAssetData `json:"assets,omitempty"`
}

// Tenant represents an isolated tenant with custom domains, backend credentials, and styling.
type Tenant struct {
	ID                  int64       `json:"id"`
	Slug                string      `json:"slug"`
	Name                string      `json:"name"`
	Domains             string      `json:"domains"`
	IsDefault           bool        `json:"is_default"`
	ProviderType        string      `json:"provider_type"`
	OAuthClientID       string      `json:"oauth_client_id"`
	OAuthAuthority      string      `json:"oauth_authority"`
	OAuthClientSecret   string      `json:"-"`
	HasClientSecret     bool        `json:"has_client_secret"`
	ExchangeDomain      string      `json:"exchange_domain"`
	MaxDays             int         `json:"max_days"`
	MaxRoomLists        int         `json:"max_room_lists"`
	MaxRooms            int         `json:"max_rooms"`
	MaxItems            int         `json:"max_items"`
	RoomBlacklist       string      `json:"room_blacklist"`
	PrivacyMode         string      `json:"privacy_mode"`
	ThemePreset         string      `json:"theme_preset"`
	Theme               ThemeConfig `json:"theme"`
	CustomCSS           string      `json:"custom_css"`
	PollIntervalSeconds int         `json:"poll_interval_seconds"`
	CreatedAt           time.Time   `json:"created_at"`
	UpdatedAt           time.Time   `json:"updated_at"`
}

// DomainList returns normalized lowercased hostnames configured for this tenant.
func (t Tenant) DomainList() []string {
	var out []string
	for _, part := range strings.Split(t.Domains, ",") {
		d := strings.ToLower(strings.TrimSpace(part))
		if d != "" {
			out = append(out, d)
		}
	}
	return out
}

// BlacklistSet returns lowercased room emails on the blacklist.
func (t Tenant) BlacklistSet() map[string]struct{} {
	res := make(map[string]struct{})
	normalized := strings.ReplaceAll(t.RoomBlacklist, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, ",", "\n")
	for _, line := range strings.Split(normalized, "\n") {
		email := strings.ToLower(strings.TrimSpace(line))
		if email != "" {
			res[email] = struct{}{}
		}
	}
	return res
}

// MarshalTheme serializes the ThemeConfig into JSON.
func (t Tenant) MarshalTheme() (string, error) {
	b, err := json.Marshal(t.Theme)
	if err != nil {
		return "", err
	}
	return string(b), nil
}

// Appointment mirrors the calendar event contract used by MeetEasier UI and API.
type Appointment struct {
	Subject   string `json:"Subject"`
	Organizer string `json:"Organizer"`
	Start     int64  `json:"Start"`
	End       int64  `json:"End"`
	Private   bool   `json:"Private"`
}

// Room mirrors the room object returned by /api/rooms and rendered by templates.
type Room struct {
	Roomlist     string        `json:"Roomlist"`
	Name         string        `json:"Name"`
	RoomAlias    string        `json:"RoomAlias"`
	Email        string        `json:"Email"`
	Busy         bool          `json:"Busy"`
	ErrorMessage string        `json:"ErrorMessage,omitempty"`
	Appointments []Appointment `json:"Appointments"`
}

// RoomlistSlug returns the CSS-safe filter slug for the room's roomlist.
func (r Room) RoomlistSlug() string {
	return Slugify(r.Roomlist)
}

// Slugify converts a display name to a URL/CSS-safe lowercased hyphenated slug.
func Slugify(input string) string {
	fields := strings.Fields(strings.ToLower(strings.TrimSpace(input)))
	return strings.Join(fields, "-")
}

// TenantStatus summarizes cached calendar sync status for a tenant.
type TenantStatus struct {
	TenantID      int64     `json:"tenant_id"`
	TenantSlug    string    `json:"tenant_slug"`
	Rooms         []Room    `json:"rooms"`
	Roomlists     []string  `json:"roomlists"`
	LastSync      time.Time `json:"last_sync"`
	LastError     string    `json:"last_error"`
	IsStale       bool      `json:"is_stale"`
	Syncing       bool      `json:"syncing"`
	Initialized   bool      `json:"initialized"`
}

// AdminUser represents an administrator account stored in SQLite.
type AdminUser struct {
	ID           int64
	Username     string
	PasswordHash string
	CreatedAt    time.Time
}

// TenantAsset stores uploaded tenant images (logo, footer, etc.) in SQLite.
type TenantAsset struct {
	TenantSlug  string
	AssetKey    string
	ContentType string
	Data        []byte
	UpdatedAt   time.Time
}

