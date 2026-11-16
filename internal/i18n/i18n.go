package i18n

import (
	"fmt"
	"strings"
	"time"

	"github.com/probits-as/meeteasier/internal/model"
)

const (
	LangEnglish   = "en"
	LangNorwegian = "nb"
)

// LanguageOption describes a selectable UI language.
type LanguageOption struct {
	Code string
	Name string
}

// SupportedLanguages returns all supported interface languages in display order.
func SupportedLanguages() []LanguageOption {
	return []LanguageOption{
		{Code: LangEnglish, Name: "English"},
		{Code: LangNorwegian, Name: "Norsk"},
	}
}

// Normalize maps aliases ("no", "nb-NO", "en-US", etc.) to canonical language codes ("en" or "nb").
func Normalize(code string) string {
	c := strings.ToLower(strings.TrimSpace(code))
	switch {
	case c == "nb" || c == "no" || c == "nn" || strings.HasPrefix(c, "nb-") || strings.HasPrefix(c, "no-") || strings.HasPrefix(c, "nn-"):
		return LangNorwegian
	default:
		return LangEnglish
	}
}

// NormalizeTenantLanguage normalizes an optional per-tenant language setting ("" means inherit system default).
func NormalizeTenantLanguage(code string) string {
	if strings.TrimSpace(code) == "" {
		return ""
	}
	return Normalize(code)
}

// Translations holds localized UI strings for Admin, Flightboard, and Single-Room screens.
type Translations struct {
	LangCode      string
	Locale        string
	Default24Hour bool

	// Admin chrome & auth
	AdminLoginPageTitle     string
	AdminDashboardPageTitle string
	AdminHeader             string
	LoggedInAs              string
	OpenDefaultView         string
	Logout                  string
	LanguageLabel           string
	LoginSubtitle           string
	UsernameLabel           string
	PasswordLabel           string
	LoginButton             string
	ErrInvalidRequest       string
	ErrInvalidCredentials   string

	// Admin sidebar
	TenantsHeading           string
	DefaultBadge             string
	StaleCacheBadge          string
	ErrorBadge               string
	RoomsCountFmt            string
	CreateTenantHeading      string
	DisplayNameLabel         string
	DisplayNamePlaceholder   string
	SlugLabel                string
	SlugPlaceholder          string
	CustomDomainsLabel       string
	CustomDomainsPlaceholder string
	PresetLabel              string
	PresetStandardOption     string
	PresetCompactOption      string
	CreateTenantButton       string
	ImportTenantHeading      string
	ImportTenantHint         string
	ChooseJSONFile           string
	ImportTenantButton       string
	ChangePasswordHeading    string
	NewPasswordLabel         string
	UpdatePasswordButton     string

	// Admin editor
	SelectTenantPrompt         string
	PreviewNoDNS               string
	TestAndSyncNow             string
	DeleteTenant               string
	DeleteTenantConfirm        string
	SyncStatusLabel            string
	SyncStatusCountsFmt        string
	LastUpdatedFmt             string
	NotSyncedYet               string
	LastErrorLabel             string
	StaleCacheHint             string
	QuickPresetTitle           string
	LoadStandardTheme          string
	LoadCompactTheme           string
	ExportThemeJSON            string
	PresetOverwriteHint        string
	ImportThemeSummary         string
	ThemeFileLabel             string
	IncludeMetadataLabel       string
	PasteJSONLabel             string
	ImportThemeButton          string
	Sec1Title                  string
	SlugHintLabel              string
	CustomDomainsEditLabel     string
	CustomDomainsPlaceholder2  string
	DomainsHint                string
	AllowedIPsLabel            string
	AllowedIPsPlaceholder      string
	AllowedIPsHint             string
	ErrInvalidAllowedIPsFmt    string
	IPBlockedPageTitle         string
	IPBlockedHeading           string
	IPBlockedMessage           string
	IPBlockedClientIPFmt       string
	PollIntervalLabel          string
	DefaultTenantCheckbox      string
	TenantLanguageLabel        string
	LangSystemDefaultOption    string
	Sec2Title                  string
	CalendarProviderLabel      string
	ProviderMockOption         string
	ProviderMSGraphOption      string
	PrivacyModeLabel           string
	PrivacyNormalOption        string
	PrivacyOrgAsSubjOption     string
	PrivacyHideOption          string
	OAuthClientIDLabel         string
	OAuthClientIDPlaceholder   string
	OAuthAuthorityLabel        string
	OAuthSecretLabel           string
	SavedBadge                 string
	OAuthSecretPlaceholder     string
	ExchangeDomainLabel        string
	MaxDaysLabel               string
	MaxRoomListsLabel          string
	MaxRoomsLabel              string
	MaxItemsLabel              string
	RoomBlacklistLabel         string
	Sec2bTitle                 string
	BookingEnabledLabel        string
	BookingPINLabel            string
	BookingPINActiveBadge      string
	BookingPINPlaceholder      string
	ClearBookingPINLabel       string
	BookingDefaultSubjectLabel string
	BookingOrganizerLabel      string
	BookingOrganizerHint       string
	Sec3Title                  string
	RowLayoutLabel             string
	RowLayoutStandardOption    string
	RowLayoutCompactOption     string
	SingleRoomStatusSizeLabel  string
	ShowWeekdayInTimeLabel     string
	TimeFormatLabel            string
	TimeFormatAutoOption       string
	TimeFormat24HOption        string
	TimeFormat12HOption        string
	NavbarLogoLabel            string
	FooterBannerLabel          string
	SingleRoomLogoLabel        string
	Sec4Title                  string
	ColorBodyBg                string
	ColorCardBg                string
	ColorCardBorder            string
	ColorTitleBarBg            string
	ColorPrimaryFont           string
	ColorOpen                  string
	ColorOpenDark              string
	ColorBusy                  string
	ColorUpcoming              string
	ColorError                 string
	Sec5Title                  string
	Sec5Hint                   string
	LabelNavbarTitle           string
	LabelFilterTitle           string
	LabelFilterAllTitle        string
	LabelStatusAvailable       string
	LabelStatusBusy            string
	LabelStatusError           string
	LabelNowPrefix             string
	LabelNextUp                string
	LabelUpcomingTitle         string
	DefaultNavbarTitle         string
	DefaultFilterTitle         string
	DefaultFilterAllTitle      string
	DefaultStatusAvailable     string
	DefaultStatusBusy          string
	DefaultStatusError         string
	DefaultNowPrefix           string
	DefaultNextUp              string
	DefaultUpcomingTitle       string
	Sec6Title                  string
	CustomCSSLabel             string
	SaveChangesButton          string

	// Admin flashes & errors
	ErrInvalidForm          string
	FlashPasswordUpdated    string
	FlashTenantCreated      string
	ErrInvalidTenantID      string
	ErrTenantNotFound       string
	ErrMultipartTooLarge    string
	FlashChangesSaved       string
	FlashPresetLoadedFmt    string
	ErrSyncFailedFmt        string
	FlashSyncOKFmt          string
	FlashTenantDeleted      string
	ErrJSONFormRead         string
	ErrJSONFileRead         string
	ErrJSONEmpty            string
	ErrJSONInvalidFmt       string
	FlashTenantUpdatedJSON  string
	FlashTenantImportedJSON string
	FlashThemeImportedJSON  string
	FlashLanguageUpdated    string

	// Public flightboard & single-room & touch booking
	StaleBannerFmt              string
	LastSyncedFmt               string
	FetchRoomsErrorFmt          string
	LoadingAlt                  string
	SingleRoomErrorTitle        string
	SingleRoomErrorBody         string
	BookingDisabledErr          string
	BookingInvalidPINErr        string
	SubjectPlaceholder          string
	BookAfterSubjectPlaceholder string
	PINPlaceholder              string
	BookNowLabel                string
	RoomSoonBusyLabel           string
	CurrentMeetingLabel         string
	EndNowButton                string
	ExtendFmt                   string
	BookAfterLabel              string
	MinSuffix                   string
	FlashBookedNowFmt           string
	FlashBookedAfterFmt         string
	FlashExtendedFmt            string
	FlashEndedNow               string
	FlashBookingUpdated         string
	DefaultBookedSubject        string
}

var catalog = map[string]Translations{
	LangEnglish: {
		LangCode:      "en",
		Locale:        "en-US",
		Default24Hour: false,

		AdminLoginPageTitle:     "Sign in - MeetEasier Admin",
		AdminDashboardPageTitle: "Tenants - MeetEasier Admin",
		AdminHeader:             "MeetEasier Admin",
		LoggedInAs:              "Signed in as",
		OpenDefaultView:         "Open default display",
		Logout:                  "Sign out",
		LanguageLabel:           "Language",
		LoginSubtitle:           "Sign in to configure tenants, custom domains, Microsoft Graph, booking, and styling.",
		UsernameLabel:           "Username",
		PasswordLabel:           "Password",
		LoginButton:             "Sign in",
		ErrInvalidRequest:       "Invalid request",
		ErrInvalidCredentials:   "Invalid username or password",

		TenantsHeading:           "Tenants",
		DefaultBadge:             "Default",
		StaleCacheBadge:          "Stale cache",
		ErrorBadge:               "Error",
		RoomsCountFmt:            "%d rooms",
		CreateTenantHeading:      "Create new tenant",
		DisplayNameLabel:         "Display name",
		DisplayNamePlaceholder:   "e.g. Demo Company Ltd",
		SlugLabel:                "Slug",
		SlugPlaceholder:          "demo-company",
		CustomDomainsLabel:       "Custom domains (comma-separated)",
		CustomDomainsPlaceholder: "rooms.demo.example",
		PresetLabel:              "Starter theme (Preset)",
		PresetStandardOption:     "Standard (MeetEasier)",
		PresetCompactOption:      "Compact (2-column, clickable row)",
		CreateTenantButton:       "Create tenant",
		ImportTenantHeading:      "Import tenant from JSON",
		ImportTenantHint:         "Creates or updates a complete tenant with colors, domains, and logos from a theme JSON bundle.",
		ChooseJSONFile:           "Select .json file",
		ImportTenantButton:       "Import tenant",
		ChangePasswordHeading:    "Change admin password",
		NewPasswordLabel:         "New password (min 6 chars)",
		UpdatePasswordButton:     "Update password",

		SelectTenantPrompt:         "Select a tenant from the list on the left to edit its configuration.",
		PreviewNoDNS:               "Preview without DNS:",
		TestAndSyncNow:             "Test & Sync now",
		DeleteTenant:               "Delete tenant",
		DeleteTenantConfirm:        "Are you sure you want to delete this tenant?",
		SyncStatusLabel:            "Sync status:",
		SyncStatusCountsFmt:        "%d rooms, %d locations.",
		LastUpdatedFmt:             " Last updated: %s.",
		NotSyncedYet:               "Not synced yet.",
		LastErrorLabel:             "Last error:",
		StaleCacheHint:             " (Still displaying last known room data from cache.)",
		QuickPresetTitle:           "Quick theme presets",
		LoadStandardTheme:          "Load Standard theme",
		LoadCompactTheme:           "Load Compact theme (2-column)",
		ExportThemeJSON:            "Export theme (JSON)",
		PresetOverwriteHint:        "Applying a preset replaces colors, layout, and labels with that template.",
		ImportThemeSummary:         "Import theme from JSON (incl. images/logos)",
		ThemeFileLabel:             "Choose .json theme file",
		IncludeMetadataLabel:       "Also update display name, domains, and default status from JSON",
		PasteJSONLabel:             "Or paste JSON directly",
		ImportThemeButton:          "Import theme",
		Sec1Title:                  "1. General & Custom Domains",
		SlugHintLabel:              "Slug (used in ?tenant= and asset URLs)",
		CustomDomainsEditLabel:     "Custom domains (comma-separated, without protocol)",
		CustomDomainsPlaceholder2:  "rooms.company.com, displays.company.com",
		DomainsHint:                "Matches HTTP Host header automatically (also supports *.domain.com).",
		AllowedIPsLabel:            "Allowed IP addresses / CIDR (optional)",
		AllowedIPsPlaceholder:      "e.g. 195.88.54.10, 10.0.0.0/8, 192.168.1.0/24",
		AllowedIPsHint:             "Leave empty to allow all IPs. Supports single IPs and CIDR ranges (comma- or line-separated).",
		ErrInvalidAllowedIPsFmt:    "Invalid IP or CIDR in allowed list: %v",
		IPBlockedPageTitle:         "Access Restricted",
		IPBlockedHeading:           "Access restricted by IP address",
		IPBlockedMessage:           "This conference room display is only accessible from authorized network locations.",
		IPBlockedClientIPFmt:       "Your IP address: %s",
		PollIntervalLabel:          "Polling interval (seconds)",
		DefaultTenantCheckbox:      "Default tenant (fallback when no custom domains match)",
		TenantLanguageLabel:        "Display / booking language",
		LangSystemDefaultOption:    "System default",
		Sec2Title:                  "2. Calendar Source (Microsoft Graph / Mock)",
		CalendarProviderLabel:      "Calendar provider",
		ProviderMockOption:         "Mock / Demo data (offline-friendly)",
		ProviderMSGraphOption:      "Microsoft Graph API (M365)",
		PrivacyModeLabel:           "Privacy / Meeting title",
		PrivacyNormalOption:        "Normal (show meeting subject, hide Private only)",
		PrivacyOrgAsSubjOption:     "Organizer only (show organizer as title)",
		PrivacyHideOption:          `Hide all subjects (show "Private")`,
		OAuthClientIDLabel:         "OAuth Client ID",
		OAuthClientIDPlaceholder:   "Azure App Registration Client ID",
		OAuthAuthorityLabel:        "OAuth Authority",
		OAuthSecretLabel:           "OAuth Client Secret",
		SavedBadge:                 "Saved",
		OAuthSecretPlaceholder:     "Leave empty to keep existing secret",
		ExchangeDomainLabel:        "Domain (used for Mock / fallback)",
		MaxDaysLabel:               "Max days ahead",
		MaxRoomListsLabel:          "Max room lists",
		MaxRoomsLabel:              "Max rooms per list",
		MaxItemsLabel:              "Max events per room",
		RoomBlacklistLabel:         "Blacklisted rooms (email per line or comma)",
		Sec2bTitle:                 "2b. Single-Room Touch Booking",
		BookingEnabledLabel:        "Enable quick book, extend, and \"End now\" on /single-room/* (requires Calendars.ReadWrite in M365)",
		BookingPINLabel:            "Booking PIN / password (optional)",
		BookingPINActiveBadge:      "PIN active",
		BookingPINPlaceholder:      "Leave empty to keep existing (empty without active = open touchscreen)",
		ClearBookingPINLabel:       "Remove existing PIN",
		BookingDefaultSubjectLabel: "Default meeting subject",
		BookingOrganizerLabel:      "Kiosk / Organizer email (optional)",
		BookingOrganizerHint:       "Empty = write directly to the room calendar. Set = create event from this account inviting the room as a resource so Exchange approval rules apply.",
		Sec3Title:                  "3. Layout, Logos & Images",
		RowLayoutLabel:             "Flightboard row layout",
		RowLayoutStandardOption:    "Standard (With display icon on the right)",
		RowLayoutCompactOption:     "Compact (Clickable row, Now + Next in 2 columns, portrait support)",
		SingleRoomStatusSizeLabel:  "Single-room status font size",
		ShowWeekdayInTimeLabel:     "Show weekday and date on Flightboard event times",
		TimeFormatLabel:            "Clock format",
		TimeFormatAutoOption:       "Follow language (12-hour AM/PM for English, 24-hour for Norwegian)",
		TimeFormat24HOption:        "24-hour (14:30)",
		TimeFormat12HOption:        "12-hour (2:30 PM)",
		NavbarLogoLabel:            "Navbar Logo URL",
		FooterBannerLabel:          "Footer Banner / Slogan URL",
		SingleRoomLogoLabel:        "Single-Room Logo URL",
		Sec4Title:                  "4. Colors",
		ColorBodyBg:                "Background",
		ColorCardBg:                "Card background",
		ColorCardBorder:            "Card border",
		ColorTitleBarBg:            "Top bar background",
		ColorPrimaryFont:           "Text color",
		ColorOpen:                  "Available (accent)",
		ColorOpenDark:              "Available (single-room)",
		ColorBusy:                  "Busy",
		ColorUpcoming:              "Next / Upcoming",
		ColorError:                 "Error",
		Sec5Title:                  "5. Labels & Text",
		Sec5Hint:                   "Leave fields empty to automatically follow the selected language. Fill in any field below to override it.",
		LabelNavbarTitle:           "Navbar title",
		LabelFilterTitle:           "Filter menu title",
		LabelFilterAllTitle:        "All rooms option",
		LabelStatusAvailable:       "Status: Available",
		LabelStatusBusy:            "Status: Busy",
		LabelStatusError:           "Status: Error",
		LabelNowPrefix:             `Current meeting prefix (e.g. "Now")`,
		LabelNextUp:                `Next meeting prefix (e.g. "Next Up")`,
		LabelUpcomingTitle:         `Single-room heading (e.g. "Upcoming")`,
		DefaultNavbarTitle:         "Conference Room Availability",
		DefaultFilterTitle:         "Locations",
		DefaultFilterAllTitle:      "All Conference Rooms",
		DefaultStatusAvailable:     "Open",
		DefaultStatusBusy:          "Busy",
		DefaultStatusError:         "Error",
		DefaultNowPrefix:           "Now",
		DefaultNextUp:              "Next Up",
		DefaultUpcomingTitle:       "Upcoming",
		Sec6Title:                  "6. Custom CSS (Optional)",
		CustomCSSLabel:             "Custom CSS for this tenant",
		SaveChangesButton:          "Save changes",

		ErrInvalidForm:          "Invalid form submission",
		FlashPasswordUpdated:    "Password updated",
		FlashTenantCreated:      "Tenant created",
		ErrInvalidTenantID:      "Invalid tenant ID",
		ErrTenantNotFound:       "Tenant not found",
		ErrMultipartTooLarge:    "Could not read form / upload (max 8 MB)",
		FlashChangesSaved:       "Changes saved",
		FlashPresetLoadedFmt:    "Theme preset '%s' loaded",
		ErrSyncFailedFmt:        "Sync failed: %s",
		FlashSyncOKFmt:          "Sync OK (%d rooms loaded)",
		FlashTenantDeleted:      "Tenant deleted",
		ErrJSONFormRead:         "Could not read theme import form",
		ErrJSONFileRead:         "Could not read uploaded JSON file",
		ErrJSONEmpty:            "Select a JSON file or paste JSON first",
		ErrJSONInvalidFmt:       "Invalid JSON: %v",
		FlashTenantUpdatedJSON:  "Tenant updated from JSON",
		FlashTenantImportedJSON: "Tenant imported from JSON",
		FlashThemeImportedJSON:  "Theme imported from JSON",
		FlashLanguageUpdated:    "Language updated",

		StaleBannerFmt:              "Showing cached calendar data (Microsoft Graph unavailable: %s)",
		LastSyncedFmt:               "Last synced: %s",
		FetchRoomsErrorFmt:          "Could not fetch conference rooms: %s",
		LoadingAlt:                  "Loading...",
		SingleRoomErrorTitle:        "Sorry :(",
		SingleRoomErrorBody:         "Either there was an error in processing or this room does not exist.",
		BookingDisabledErr:          "Room booking is disabled for this tenant",
		BookingInvalidPINErr:        "Invalid PIN / password",
		SubjectPlaceholder:          "Optional meeting subject...",
		BookAfterSubjectPlaceholder: "Optional subject (book after)...",
		PINPlaceholder:              "PIN / password",
		BookNowLabel:                "Book this room now:",
		RoomSoonBusyLabel:           "Room is busy soon",
		CurrentMeetingLabel:         "Current meeting:",
		EndNowButton:                "End now",
		ExtendFmt:                   "Extend +%d min",
		BookAfterLabel:              "Book after meeting:",
		MinSuffix:                   "min",
		FlashBookedNowFmt:           "Room booked for %d minutes",
		FlashBookedAfterFmt:         "Room booked for %d min after current meeting",
		FlashExtendedFmt:            "Meeting extended by %d minutes",
		FlashEndedNow:               "Meeting ended and room is now available",
		FlashBookingUpdated:         "Booking updated",
		DefaultBookedSubject:        "Booked by MeetEasier",
	},
	LangNorwegian: {
		LangCode:      "nb",
		Locale:        "nb-NO",
		Default24Hour: true,

		AdminLoginPageTitle:     "Logg inn - MeetEasier Admin",
		AdminDashboardPageTitle: "Tenants - MeetEasier Admin",
		AdminHeader:             "MeetEasier Admin",
		LoggedInAs:              "Innlogget som",
		OpenDefaultView:         "Åpne standardvisning",
		Logout:                  "Logg ut",
		LanguageLabel:           "Språk",
		LoginSubtitle:           "Logg inn for å konfigurere tenants, custom domener, Microsoft Graph, rombooking og utseende.",
		UsernameLabel:           "Brukernavn",
		PasswordLabel:           "Passord",
		LoginButton:             "Logg inn",
		ErrInvalidRequest:       "Ugyldig forespørsel",
		ErrInvalidCredentials:   "Feil brukernavn eller passord",

		TenantsHeading:           "Tenants",
		DefaultBadge:             "Default",
		StaleCacheBadge:          "Stale cache",
		ErrorBadge:               "Feil",
		RoomsCountFmt:            "%d rom",
		CreateTenantHeading:      "Opprett ny tenant",
		DisplayNameLabel:         "Visningsnavn",
		DisplayNamePlaceholder:   "F.eks. Demo Bedrift AS",
		SlugLabel:                "Slug",
		SlugPlaceholder:          "demo-bedrift",
		CustomDomainsLabel:       "Custom domener (kommaseparert)",
		CustomDomainsPlaceholder: "moterom.demo.no",
		PresetLabel:              "Start-tema (Preset)",
		PresetStandardOption:     "Standard (MeetEasier)",
		PresetCompactOption:      "Kompakt (2-kolonne, klikkbar rad)",
		CreateTenantButton:       "Opprett tenant",
		ImportTenantHeading:      "Importer tenant fra JSON",
		ImportTenantHint:         "Oppretter eller oppdaterer en tenant komplett med farger, domener og logoer fra en theme JSON-fil.",
		ChooseJSONFile:           "Velg .json-fil",
		ImportTenantButton:       "Importer tenant",
		ChangePasswordHeading:    "Bytt admin-passord",
		NewPasswordLabel:         "Nytt passord (min 6 tegn)",
		UpdatePasswordButton:     "Oppdater passord",

		SelectTenantPrompt:         "Velg en tenant fra listen til venstre for å redigere oppsettet.",
		PreviewNoDNS:               "Forhåndsvisning uten DNS:",
		TestAndSyncNow:             "Test & Synkroniser nå",
		DeleteTenant:               "Slett tenant",
		DeleteTenantConfirm:        "Er du sikker på at du vil slette denne tenanten?",
		SyncStatusLabel:            "Synkroniseringsstatus:",
		SyncStatusCountsFmt:        "%d rom, %d lokasjoner.",
		LastUpdatedFmt:             " Sist oppdatert: %s.",
		NotSyncedYet:               "Ikke synkronisert ennå.",
		LastErrorLabel:             "Siste feil:",
		StaleCacheHint:             " (Viser fortsatt siste kjente romdata fra cache.)",
		QuickPresetTitle:           "Hurtigvalg for tema-preset",
		LoadStandardTheme:          "Last Standard-tema",
		LoadCompactTheme:           "Last Kompakt-tema (norsk 2-kolonne)",
		ExportThemeJSON:            "Eksporter tema (JSON)",
		PresetOverwriteHint:        "Preset overskriver farger, layout og etiketter med valgt mal.",
		ImportThemeSummary:         "Importer tema fra JSON (inkl. bilder/logoer)",
		ThemeFileLabel:             "Velg .json tema-fil",
		IncludeMetadataLabel:       "Oppdater også visningsnavn, domener og standard-status fra JSON",
		PasteJSONLabel:             "Eller lim inn JSON direkte",
		ImportThemeButton:          "Importer tema",
		Sec1Title:                  "1. Generelt & Custom Domains",
		SlugHintLabel:              "Slug (brukes i ?tenant= og asset-URL)",
		CustomDomainsEditLabel:     "Custom Domener (kommaseparert, uten protokoll)",
		CustomDomainsPlaceholder2:  "moterom.kunde.no, skjermer.kunde.no",
		DomainsHint:                "Matcher HTTP Host-header automatisk (støtter også *.domene.no).",
		AllowedIPsLabel:            "Tillatte IP-adresser / CIDR (valgfritt)",
		AllowedIPsPlaceholder:      "f.eks. 195.88.54.10, 10.0.0.0/8, 192.168.1.0/24",
		AllowedIPsHint:             "La stå tom for å tillate alle IP-adresser. Støtter enkelt-IP og CIDR-nettverk (komma- eller linjeseparert).",
		ErrInvalidAllowedIPsFmt:    "Ugyldig IP eller CIDR i tilgangslisten: %v",
		IPBlockedPageTitle:         "Tilgang begrenset",
		IPBlockedHeading:           "Tilgang begrenset til godkjente nettverk",
		IPBlockedMessage:           "Denne møteromsvisningen er kun tilgjengelig fra godkjente IP-adresser / nettverk.",
		IPBlockedClientIPFmt:       "Din IP-adresse: %s",
		PollIntervalLabel:          "Polling-intervall (sekunder)",
		DefaultTenantCheckbox:      "Standard-tenant (fallback hvis ingen domener matcher)",
		TenantLanguageLabel:        "Visnings-/bookingspråk",
		LangSystemDefaultOption:    "Følg systemstandard",
		Sec2Title:                  "2. Kalender-kilde (Microsoft Graph / Mock)",
		CalendarProviderLabel:      "Kalender-provider",
		ProviderMockOption:         "Mock / Demo-data (offline-vennlig)",
		ProviderMSGraphOption:      "Microsoft Graph API (M365)",
		PrivacyModeLabel:           "Personvern / Møtetittel",
		PrivacyNormalOption:        "Normal (vis møteemne, skjul kun Private)",
		PrivacyOrgAsSubjOption:     "Kun arrangør (vis arrangør som tittel)",
		PrivacyHideOption:          `Skjul alle emner (vis "Private")`,
		OAuthClientIDLabel:         "OAuth Client ID",
		OAuthClientIDPlaceholder:   "Azure App Registration Client ID",
		OAuthAuthorityLabel:        "OAuth Authority",
		OAuthSecretLabel:           "OAuth Client Secret",
		SavedBadge:                 "Lagret",
		OAuthSecretPlaceholder:     "La stå tom for å beholde eksisterende",
		ExchangeDomainLabel:        "Domene (brukt for Mock / fallback)",
		MaxDaysLabel:               "Maks dager frem",
		MaxRoomListsLabel:          "Maks romlister",
		MaxRoomsLabel:              "Maks rom per liste",
		MaxItemsLabel:              "Maks avtaler per rom",
		RoomBlacklistLabel:         "Svartelistede rom (e-post per linje eller komma)",
		Sec2bTitle:                 "2b. Rombooking på enkeltrom-skjerm (Touchskjerm)",
		BookingEnabledLabel:        "Aktiver hurtigbooking, forlengelse og «Avslutt nå» på /single-room/* (krever Calendars.ReadWrite i M365)",
		BookingPINLabel:            "Booking PIN / passord (valgfritt)",
		BookingPINActiveBadge:      "PIN aktiv",
		BookingPINPlaceholder:      "La stå tom for å beholde eksisterende (tom uten aktiv = åpen skjerm)",
		ClearBookingPINLabel:       "Fjern eksisterende PIN-kode",
		BookingDefaultSubjectLabel: "Standard møtetittel",
		BookingOrganizerLabel:      "Kiosk-/Arrangør e-post (valgfritt)",
		BookingOrganizerHint:       "Tom = skriver direkte i rommets kalender. Satt = oppretter møtet fra denne kontoen med rommet som ressurs slik at Exchange-godkjenningsregler følges.",
		Sec3Title:                  "3. Layout, Logoer & Bilder",
		RowLayoutLabel:             "Rad-layout på Flightboard",
		RowLayoutStandardOption:    "Standard (Med skjerm-ikon til høyre)",
		RowLayoutCompactOption:     "Kompakt (Klikkbar rad, Nå + Neste i 2 kolonner, stående støtte)",
		SingleRoomStatusSizeLabel:  "Enkeltrom status-skriftstørrelse",
		ShowWeekdayInTimeLabel:     "Vis ukedag og dato på tidspunkter i Flightboard",
		TimeFormatLabel:            "Klokkeformat",
		TimeFormatAutoOption:       "Følg valgt språk (24-timers for norsk, 12-timers AM/PM for engelsk)",
		TimeFormat24HOption:        "24-timers (14:30)",
		TimeFormat12HOption:        "12-timers (2:30 PM)",
		NavbarLogoLabel:            "Navbar Logo URL",
		FooterBannerLabel:          "Footer Banner / Slagord URL",
		SingleRoomLogoLabel:        "Enkeltrom Logo URL",
		Sec4Title:                  "4. Farger",
		ColorBodyBg:                "Bakgrunn",
		ColorCardBg:                "Kort-bakgrunn",
		ColorCardBorder:            "Kort-kantlinje",
		ColorTitleBarBg:            "Topplinje-bakgrunn",
		ColorPrimaryFont:           "Tekstfarge",
		ColorOpen:                  "Ledig (aksent)",
		ColorOpenDark:              "Ledig (enkeltrom)",
		ColorBusy:                  "Opptatt",
		ColorUpcoming:              "Neste / Kommende",
		ColorError:                 "Feil",
		Sec5Title:                  "5. Tekster & Språk",
		Sec5Hint:                   "La feltene stå tomme for å følge valgt språk automatisk. Fyll inn et felt nedenfor kun dersom du vil overstyre det.",
		LabelNavbarTitle:           "Tittel i toppmeny",
		LabelFilterTitle:           "Filter-meny tittel",
		LabelFilterAllTitle:        "Alle rom-valg",
		LabelStatusAvailable:       "Status: Ledig",
		LabelStatusBusy:            "Status: Opptatt",
		LabelStatusError:           "Status: Feil",
		LabelNowPrefix:             `Prefiks pågående møte (f.eks. "Nå")`,
		LabelNextUp:                `Prefiks neste møte (f.eks. "Neste")`,
		LabelUpcomingTitle:         `Enkeltrom overskrift (f.eks. "Kommer")`,
		DefaultNavbarTitle:         "Møterom",
		DefaultFilterTitle:         "Lokasjoner",
		DefaultFilterAllTitle:      "Alle rom",
		DefaultStatusAvailable:     "Ledig",
		DefaultStatusBusy:          "Opptatt",
		DefaultStatusError:         "Feil",
		DefaultNowPrefix:           "Nå",
		DefaultNextUp:              "Neste",
		DefaultUpcomingTitle:       "Kommer",
		Sec6Title:                  "6. Custom CSS (Valgfritt)",
		CustomCSSLabel:             "Egendefinert CSS for denne tenanten",
		SaveChangesButton:          "Lagre endringer",

		ErrInvalidForm:          "Ugyldig skjema",
		FlashPasswordUpdated:    "Passord oppdatert",
		FlashTenantCreated:      "Tenant opprettet",
		ErrInvalidTenantID:      "Ugyldig tenant ID",
		ErrTenantNotFound:       "Fant ikke tenant",
		ErrMultipartTooLarge:    "Kunne ikke lese skjema / opplasting (maks 8 MB)",
		FlashChangesSaved:       "Endringer lagret",
		FlashPresetLoadedFmt:    "Tema-preset '%s' lastet inn",
		ErrSyncFailedFmt:        "Synkronisering feilet: %s",
		FlashSyncOKFmt:          "Synkronisering OK (%d rom hentet)",
		FlashTenantDeleted:      "Tenant slettet",
		ErrJSONFormRead:         "Kunne ikke lese tema-skjema",
		ErrJSONFileRead:         "Kunne ikke lese opplastet JSON-fil",
		ErrJSONEmpty:            "Velg en JSON-fil eller lim inn JSON først",
		ErrJSONInvalidFmt:       "Ugyldig JSON: %v",
		FlashTenantUpdatedJSON:  "Tenant oppdatert fra JSON",
		FlashTenantImportedJSON: "Tenant importert fra JSON",
		FlashThemeImportedJSON:  "Tema importert fra JSON",
		FlashLanguageUpdated:    "Språk oppdatert",

		StaleBannerFmt:              "Viser bufret kalenderdata (Microsoft Graph utilgjengelig: %s)",
		LastSyncedFmt:               "Sist synkronisert: %s",
		FetchRoomsErrorFmt:          "Kunne ikke hente møterom: %s",
		LoadingAlt:                  "Laster...",
		SingleRoomErrorTitle:        "Beklager :(",
		SingleRoomErrorBody:         "Det oppstod en feil, eller så finnes ikke dette rommet.",
		BookingDisabledErr:          "Rombooking er deaktivert for denne tenanten",
		BookingInvalidPINErr:        "Feil PIN-kode / passord",
		SubjectPlaceholder:          "Valgfri møtetittel...",
		BookAfterSubjectPlaceholder: "Valgfri tittel (book etter)...",
		PINPlaceholder:              "PIN / passord",
		BookNowLabel:                "Book rommet nå:",
		RoomSoonBusyLabel:           "Rommet er straks opptatt",
		CurrentMeetingLabel:         "Pågående møte:",
		EndNowButton:                "Avslutt nå",
		ExtendFmt:                   "Forleng +%d min",
		BookAfterLabel:              "Book etter møtet:",
		MinSuffix:                   "min",
		FlashBookedNowFmt:           "Rommet er booket i %d minutter",
		FlashBookedAfterFmt:         "Rommet er booket i %d min etter pågående møte",
		FlashExtendedFmt:            "Møtet er forlenget med %d minutter",
		FlashEndedNow:               "Møtet er avsluttet og rommet er ledig",
		FlashBookingUpdated:         "Booking oppdatert",
		DefaultBookedSubject:        "Booket via MeetEasier",
	},
}

// For returns the translation bundle for the requested language code, defaulting to English.
func For(lang string) Translations {
	if tr, ok := catalog[Normalize(lang)]; ok {
		return tr
	}
	return catalog[LangEnglish]
}

// PlaceholderLabels returns the default Section 5 template labels for this translation language.
func (tr Translations) PlaceholderLabels() model.ThemeLabels {
	return model.ThemeLabels{
		NavbarTitle:     tr.DefaultNavbarTitle,
		FilterTitle:     tr.DefaultFilterTitle,
		FilterAllTitle:  tr.DefaultFilterAllTitle,
		StatusAvailable: tr.DefaultStatusAvailable,
		StatusBusy:      tr.DefaultStatusBusy,
		StatusError:     tr.DefaultStatusError,
		NowPrefix:       tr.DefaultNowPrefix,
		NextUp:          tr.DefaultNextUp,
		UpcomingTitle:   tr.DefaultUpcomingTitle,
	}
}

func pickLabel(override, fallback string) string {
	if strings.TrimSpace(override) != "" {
		return override
	}
	return fallback
}

// EffectiveLabels resolves template-dependent display labels, using custom overrides when non-empty
// and falling back to defaults for the active language otherwise.
func (tr Translations) EffectiveLabels(theme model.ThemeConfig) model.ThemeLabels {
	nowFallback := ""
	if theme.RowLayout == model.RowLayoutCompact {
		nowFallback = tr.DefaultNowPrefix
	}
	return model.ThemeLabels{
		NavbarTitle:     pickLabel(theme.Labels.NavbarTitle, tr.DefaultNavbarTitle),
		FilterTitle:     pickLabel(theme.Labels.FilterTitle, tr.DefaultFilterTitle),
		FilterAllTitle:  pickLabel(theme.Labels.FilterAllTitle, tr.DefaultFilterAllTitle),
		StatusAvailable: pickLabel(theme.Labels.StatusAvailable, tr.DefaultStatusAvailable),
		StatusBusy:      pickLabel(theme.Labels.StatusBusy, tr.DefaultStatusBusy),
		StatusError:     pickLabel(theme.Labels.StatusError, tr.DefaultStatusError),
		NowPrefix:       pickLabel(theme.Labels.NowPrefix, nowFallback),
		NextUp:          pickLabel(theme.Labels.NextUp, tr.DefaultNextUp),
		UpcomingTitle:   pickLabel(theme.Labels.UpcomingTitle, tr.DefaultUpcomingTitle),
	}
}

// Use24Hour returns true if 24-hour time formatting should be used for the given theme and language.
func (tr Translations) Use24Hour(theme model.ThemeConfig) bool {
	switch strings.ToLower(strings.TrimSpace(theme.TimeFormat)) {
	case model.TimeFormat24H:
		return true
	case model.TimeFormat12H:
		return false
	default:
		return tr.Default24Hour
	}
}

// FormatClock formats a time.Time for initial server-rendered clock display.
func (tr Translations) FormatClock(t time.Time, theme model.ThemeConfig, withSeconds bool) string {
	if tr.Use24Hour(theme) {
		if withSeconds {
			return t.Format("15:04:05")
		}
		return t.Format("15:04")
	}
	if withSeconds {
		return t.Format("3:04:05 PM")
	}
	return t.Format("3:04 PM")
}

// FormatTimeRange formats a meeting start/end timestamp pair for initial server-rendered display.
func (tr Translations) FormatTimeRange(startMs, endMs int64, showWeekday bool, theme model.ThemeConfig) string {
	s := time.UnixMilli(startMs).Local()
	e := time.UnixMilli(endMs).Local()
	if tr.Use24Hour(theme) {
		if showWeekday {
			return fmt.Sprintf("%s - %s", s.Format("02 Mon 15:04"), e.Format("Mon 15:04"))
		}
		return fmt.Sprintf("%s - %s", s.Format("15:04"), e.Format("15:04"))
	}
	if showWeekday {
		return fmt.Sprintf("%s - %s", s.Format("02 Mon 3:04 PM"), e.Format("Mon 3:04 PM"))
	}
	return fmt.Sprintf("%s - %s", s.Format("3:04 PM"), e.Format("3:04 PM"))
}

