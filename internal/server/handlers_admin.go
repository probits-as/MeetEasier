package server

import (
	"context"
	"crypto/hmac"
	"crypto/sha256"
	"crypto/subtle"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"log/slog"
	"net/http"
	"net/url"
	"strconv"
	"strings"
	"time"

	"github.com/go-chi/chi/v5"

	"github.com/probits-as/meeteasier/internal/model"
	"github.com/probits-as/meeteasier/internal/store"
	"github.com/probits-as/meeteasier/views/admin"
)

const (
	adminCookieName = "meeteasier_admin"
	maxUploadBytes  = 2 << 20 // 2 MiB per logo asset
)

type ctxKey string

const adminUserCtxKey ctxKey = "adminUser"

func (s *Server) requireAdmin(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		cookie, err := r.Cookie(adminCookieName)
		if err != nil {
			http.Redirect(w, r, "/admin/login", http.StatusSeeOther)
			return
		}
		username, ok := s.verifySessionToken(cookie.Value)
		if !ok {
			http.Redirect(w, r, "/admin/login", http.StatusSeeOther)
			return
		}
		ctx := context.WithValue(r.Context(), adminUserCtxKey, username)
		next.ServeHTTP(w, r.WithContext(ctx))
	})
}

func (s *Server) signSessionToken(username string) string {
	exp := time.Now().Add(24 * time.Hour).Unix()
	payload := fmt.Sprintf("%s|%d", username, exp)
	mac := hmac.New(sha256.New, s.sessionSecret)
	mac.Write([]byte(payload))
	sig := base64.RawURLEncoding.EncodeToString(mac.Sum(nil))
	return base64.RawURLEncoding.EncodeToString([]byte(payload)) + "." + sig
}

func (s *Server) verifySessionToken(token string) (string, bool) {
	parts := strings.Split(token, ".")
	if len(parts) != 2 {
		return "", false
	}
	rawPayload, err := base64.RawURLEncoding.DecodeString(parts[0])
	if err != nil {
		return "", false
	}
	mac := hmac.New(sha256.New, s.sessionSecret)
	mac.Write(rawPayload)
	expectedSig := base64.RawURLEncoding.EncodeToString(mac.Sum(nil))
	if !hmac.Equal([]byte(parts[1]), []byte(expectedSig)) {
		return "", false
	}
	fields := strings.SplitN(string(rawPayload), "|", 2)
	if len(fields) != 2 {
		return "", false
	}
	exp, err := strconv.ParseInt(fields[1], 10, 64)
	if err != nil || time.Now().Unix() > exp {
		return "", false
	}
	return fields[0], true
}

func (s *Server) handleAdminLoginPage(w http.ResponseWriter, r *http.Request) {
	if cookie, err := r.Cookie(adminCookieName); err == nil {
		if _, ok := s.verifySessionToken(cookie.Value); ok {
			http.Redirect(w, r, "/admin", http.StatusSeeOther)
			return
		}
	}
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	_ = admin.AdminLoginPage("").Render(r.Context(), w)
}

func (s *Server) handleAdminLoginSubmit(w http.ResponseWriter, r *http.Request) {
	if err := r.ParseForm(); err != nil {
		w.WriteHeader(http.StatusBadRequest)
		_ = admin.AdminLoginPage("Ugyldig forespørsel").Render(r.Context(), w)
		return
	}
	username := r.FormValue("username")
	password := r.FormValue("password")

	user, err := s.store.VerifyAdmin(r.Context(), username, password)
	if err != nil {
		slog.Warn("Admin login failed", "username", username)
		w.WriteHeader(http.StatusUnauthorized)
		_ = admin.AdminLoginPage("Feil brukernavn eller passord").Render(r.Context(), w)
		return
	}

	http.SetCookie(w, &http.Cookie{
		Name:     adminCookieName,
		Value:    s.signSessionToken(user.Username),
		Path:     "/",
		HttpOnly: true,
		SameSite: http.SameSiteLaxMode,
		Secure:   r.TLS != nil || strings.EqualFold(r.Header.Get("X-Forwarded-Proto"), "https"),
		MaxAge:   86400,
	})
	slog.Info("Admin logged in", "username", user.Username)
	http.Redirect(w, r, "/admin", http.StatusSeeOther)
}

func (s *Server) handleAdminLogout(w http.ResponseWriter, r *http.Request) {
	http.SetCookie(w, &http.Cookie{
		Name:     adminCookieName,
		Value:    "",
		Path:     "/",
		HttpOnly: true,
		MaxAge:   -1,
	})
	http.Redirect(w, r, "/admin/login", http.StatusSeeOther)
}

func (s *Server) handleAdminDashboard(w http.ResponseWriter, r *http.Request) {
	tenants, err := s.store.ListTenants(r.Context())
	if err != nil {
		http.Error(w, "Database error", http.StatusInternalServerError)
		return
	}

	statuses := make(map[int64]model.TenantStatus, len(tenants))
	var selected *model.Tenant

	requestedID, _ := strconv.ParseInt(r.URL.Query().Get("tenant_id"), 10, 64)
	for i := range tenants {
		t := tenants[i]
		statuses[t.ID] = s.calMgr.GetStatus(t)
		if requestedID > 0 && t.ID == requestedID {
			selected = &tenants[i]
		}
	}
	if selected == nil && len(tenants) > 0 {
		selected = &tenants[0]
	}

	adminUser, _ := r.Context().Value(adminUserCtxKey).(string)
	flash := r.URL.Query().Get("flash")
	errMsg := r.URL.Query().Get("error")

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := admin.AdminDashboardPage(tenants, statuses, selected, flash, errMsg, adminUser).Render(r.Context(), w); err != nil {
		slog.Error("Render admin dashboard failed", "error", err)
	}
}

func (s *Server) handleAdminPasswordUpdate(w http.ResponseWriter, r *http.Request) {
	if err := r.ParseForm(); err != nil {
		redirectAdmin(w, r, 0, "", "Ugyldig skjema")
		return
	}
	adminUser, _ := r.Context().Value(adminUserCtxKey).(string)
	newPassword := r.FormValue("new_password")
	if err := s.store.UpdateAdminPassword(r.Context(), adminUser, newPassword); err != nil {
		redirectAdmin(w, r, 0, "", err.Error())
		return
	}
	slog.Info("Admin password updated", "username", adminUser)
	redirectAdmin(w, r, 0, "Passord oppdatert", "")
}

func (s *Server) handleAdminTenantCreate(w http.ResponseWriter, r *http.Request) {
	if err := r.ParseForm(); err != nil {
		redirectAdmin(w, r, 0, "", "Ugyldig skjema")
		return
	}
	name := strings.TrimSpace(r.FormValue("name"))
	slug := model.Slugify(r.FormValue("slug"))
	if slug == "" {
		slug = model.Slugify(name)
	}
	preset := r.FormValue("theme_preset")
	if preset == "" {
		preset = model.PresetStandard
	}
	privacy := model.PrivacyNormal
	if preset == model.PresetCompact {
		privacy = model.PrivacyOrganizerAsSubject
	}

	t := model.Tenant{
		Slug:                slug,
		Name:                name,
		Domains:             strings.TrimSpace(r.FormValue("domains")),
		ProviderType:        model.ProviderMock,
		MaxDays:             10,
		MaxRoomLists:        10,
		MaxRooms:            20,
		MaxItems:            6,
		PrivacyMode:         privacy,
		ThemePreset:         preset,
		Theme:               model.ThemeForPreset(preset),
		PollIntervalSeconds: 60,
	}

	if err := s.store.CreateTenant(r.Context(), &t); err != nil {
		redirectAdmin(w, r, 0, "", err.Error())
		return
	}

	_ = s.calMgr.ReloadTenants(r.Context())
	s.calMgr.SyncTenantNow(r.Context(), t)
	slog.Info("Created tenant", "slug", t.Slug, "id", t.ID)
	redirectAdmin(w, r, t.ID, "Tenant opprettet", "")
}

func (s *Server) handleAdminTenantUpdate(w http.ResponseWriter, r *http.Request) {
	id, err := strconv.ParseInt(chi.URLParam(r, "id"), 10, 64)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Ugyldig tenant ID")
		return
	}

	existing, err := s.store.GetTenantByID(r.Context(), id)
	if errors.Is(err, store.ErrNotFound) {
		redirectAdmin(w, r, 0, "", "Fant ikke tenant")
		return
	}
	if err != nil {
		redirectAdmin(w, r, 0, "", err.Error())
		return
	}

	r.Body = http.MaxBytesReader(w, r.Body, 8<<20)
	if err := r.ParseMultipartForm(8 << 20); err != nil {
		redirectAdmin(w, r, id, "", "Kunne ikke lese skjema / opplasting (maks 8 MB)")
		return
	}

	t := existing
	t.Name = strings.TrimSpace(r.FormValue("name"))
	if slugVal := model.Slugify(r.FormValue("slug")); slugVal != "" {
		t.Slug = slugVal
	}
	t.Domains = strings.TrimSpace(r.FormValue("domains"))
	t.IsDefault = r.FormValue("is_default") == "1"
	t.PollIntervalSeconds = parseIntDefault(r.FormValue("poll_interval_seconds"), 60)

	t.ProviderType = r.FormValue("provider_type")
	if t.ProviderType != model.ProviderMSGraph {
		t.ProviderType = model.ProviderMock
	}
	t.PrivacyMode = r.FormValue("privacy_mode")
	if t.PrivacyMode == "" {
		t.PrivacyMode = model.PrivacyNormal
	}
	t.OAuthClientID = strings.TrimSpace(r.FormValue("oauth_client_id"))
	t.OAuthAuthority = strings.TrimSpace(r.FormValue("oauth_authority"))
	t.OAuthClientSecret = strings.TrimSpace(r.FormValue("oauth_client_secret"))
	t.ExchangeDomain = strings.TrimSpace(r.FormValue("exchange_domain"))
	t.MaxDays = parseIntDefault(r.FormValue("max_days"), 10)
	t.MaxRoomLists = parseIntDefault(r.FormValue("max_room_lists"), 10)
	t.MaxRooms = parseIntDefault(r.FormValue("max_rooms"), 20)
	t.MaxItems = parseIntDefault(r.FormValue("max_items"), 6)
	t.RoomBlacklist = r.FormValue("room_blacklist")

	t.ThemePreset = model.PresetCustom
	t.Theme = model.ThemeConfig{
		RowLayout:                r.FormValue("row_layout"),
		LogoURL:                  strings.TrimSpace(r.FormValue("logo_url")),
		FooterImageURL:           strings.TrimSpace(r.FormValue("footer_image_url")),
		SingleRoomLogoURL:        strings.TrimSpace(r.FormValue("single_room_logo_url")),
		SingleRoomStatusFontSize: strings.TrimSpace(r.FormValue("single_room_status_font_size")),
		ShowWeekdayInTime:        r.FormValue("show_weekday_in_time") == "1",
		BodyBg:                   strings.TrimSpace(r.FormValue("body_bg")),
		CardBg:                   strings.TrimSpace(r.FormValue("card_bg")),
		CardBorder:               strings.TrimSpace(r.FormValue("card_border")),
		TitleBarBg:               strings.TrimSpace(r.FormValue("title_bar_bg")),
		PrimaryFont:              strings.TrimSpace(r.FormValue("primary_font")),
		OpenColor:                strings.TrimSpace(r.FormValue("open_color")),
		OpenDarkColor:            strings.TrimSpace(r.FormValue("open_dark_color")),
		BusyColor:                strings.TrimSpace(r.FormValue("busy_color")),
		UpcomingColor:            strings.TrimSpace(r.FormValue("upcoming_color")),
		ErrorColor:               strings.TrimSpace(r.FormValue("error_color")),
		Labels: model.ThemeLabels{
			NavbarTitle:     r.FormValue("label_navbar_title"),
			FilterTitle:     r.FormValue("label_filter_title"),
			FilterAllTitle:  r.FormValue("label_filter_all_title"),
			StatusAvailable: r.FormValue("label_status_available"),
			StatusBusy:      r.FormValue("label_status_busy"),
			StatusError:     r.FormValue("label_status_error"),
			NextUp:          r.FormValue("label_next_up"),
			NowPrefix:       r.FormValue("label_now_prefix"),
			UpcomingTitle:   r.FormValue("label_upcoming_title"),
		},
	}
	t.CustomCSS = r.FormValue("custom_css")

	// Handle optional image file uploads stored in SQLite
	if urlPath, err := s.saveUploadedAsset(r, t.Slug, "logo_file", "logo"); err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	} else if urlPath != "" {
		t.Theme.LogoURL = urlPath
	}

	if urlPath, err := s.saveUploadedAsset(r, t.Slug, "footer_file", "footer"); err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	} else if urlPath != "" {
		t.Theme.FooterImageURL = urlPath
	}

	if urlPath, err := s.saveUploadedAsset(r, t.Slug, "singleroom_logo_file", "singleroom_logo"); err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	} else if urlPath != "" {
		t.Theme.SingleRoomLogoURL = urlPath
	}

	if err := s.store.UpdateTenant(r.Context(), &t, true); err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	}

	_ = s.calMgr.ReloadTenants(r.Context())
	s.calMgr.SyncTenantNow(r.Context(), t)
	slog.Info("Updated tenant", "slug", t.Slug, "id", t.ID)
	redirectAdmin(w, r, t.ID, "Endringer lagret", "")
}

func (s *Server) handleAdminTenantPreset(w http.ResponseWriter, r *http.Request) {
	id, err := strconv.ParseInt(chi.URLParam(r, "id"), 10, 64)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Ugyldig tenant ID")
		return
	}
	t, err := s.store.GetTenantByID(r.Context(), id)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Fant ikke tenant")
		return
	}

	if err := r.ParseForm(); err != nil {
		redirectAdmin(w, r, id, "", "Ugyldig skjema")
		return
	}
	preset := r.FormValue("preset")
	t.ThemePreset = preset
	t.Theme = model.ThemeForPreset(preset)
	if preset == model.PresetCompact {
		t.PrivacyMode = model.PrivacyOrganizerAsSubject
	} else {
		t.PrivacyMode = model.PrivacyNormal
	}

	if err := s.store.UpdateTenant(r.Context(), &t, true); err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	}
	_ = s.calMgr.ReloadTenants(r.Context())
	s.calMgr.SyncTenantNow(r.Context(), t)
	redirectAdmin(w, r, id, fmt.Sprintf("Tema-preset '%s' lastet inn", preset), "")
}

func (s *Server) handleAdminTenantSync(w http.ResponseWriter, r *http.Request) {
	id, err := strconv.ParseInt(chi.URLParam(r, "id"), 10, 64)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Ugyldig tenant ID")
		return
	}
	t, err := s.store.GetTenantByID(r.Context(), id)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Fant ikke tenant")
		return
	}

	st := s.calMgr.SyncTenantNow(r.Context(), t)
	if st.LastError != "" {
		redirectAdmin(w, r, id, "", "Synkronisering feilet: "+st.LastError)
		return
	}
	redirectAdmin(w, r, id, fmt.Sprintf("Synkronisering OK (%d rom hentet)", len(st.Rooms)), "")
}

func (s *Server) handleAdminTenantDelete(w http.ResponseWriter, r *http.Request) {
	id, err := strconv.ParseInt(chi.URLParam(r, "id"), 10, 64)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Ugyldig tenant ID")
		return
	}
	if err := s.store.DeleteTenant(r.Context(), id); err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	}
	_ = s.calMgr.ReloadTenants(r.Context())
	slog.Info("Deleted tenant", "id", id)
	redirectAdmin(w, r, 0, "Tenant slettet", "")
}

func (s *Server) saveUploadedAsset(r *http.Request, slug, fieldName, assetKey string) (string, error) {
	file, _, err := r.FormFile(fieldName)
	if errors.Is(err, http.ErrMissingFile) {
		return "", nil
	}
	if err != nil {
		return "", fmt.Errorf("les opplastet fil %s: %w", fieldName, err)
	}
	defer file.Close()

	data, err := io.ReadAll(io.LimitReader(file, maxUploadBytes+1))
	if err != nil {
		return "", fmt.Errorf("les filinnhold: %w", err)
	}
	if len(data) == 0 {
		return "", nil
	}
	if len(data) > maxUploadBytes {
		return "", fmt.Errorf("filen %s er for stor (maks 2 MB)", fieldName)
	}

	contentType := http.DetectContentType(data)
	if !strings.HasPrefix(contentType, "image/") {
		// Allow SVG uploaded as text/xml or text/plain if it looks like SVG
		trimmed := strings.TrimSpace(string(data))
		if strings.Contains(trimmed, "<svg") {
			contentType = "image/svg+xml"
		} else {
			return "", fmt.Errorf("ugyldig filtype for %s (%s) — kun bilder tillatt", fieldName, contentType)
		}
	}

	if err := s.store.ValidateAndSaveAssetBytes(r.Context(), slug, assetKey, data); err != nil {
		return "", err
	}
	return fmt.Sprintf("/tenant-assets/%s/%s?v=%d", url.PathEscape(slug), url.PathEscape(assetKey), time.Now().Unix()), nil
}

func readThemeBundleFromMultipart(w http.ResponseWriter, r *http.Request) (model.ThemeBundle, error) {
	r.Body = http.MaxBytesReader(w, r.Body, 10<<20)
	if err := r.ParseMultipartForm(10 << 20); err != nil {
		return model.ThemeBundle{}, errors.New("Kunne ikke lese tema-skjema")
	}

	var rawJSON []byte
	if file, _, fileErr := r.FormFile("theme_file"); fileErr == nil {
		defer file.Close()
		var err error
		rawJSON, err = io.ReadAll(io.LimitReader(file, 8<<20))
		if err != nil {
			return model.ThemeBundle{}, errors.New("Kunne ikke lese opplastet JSON-fil")
		}
	}
	if len(strings.TrimSpace(string(rawJSON))) == 0 {
		rawJSON = []byte(r.FormValue("theme_json"))
	}
	if len(strings.TrimSpace(string(rawJSON))) == 0 {
		return model.ThemeBundle{}, errors.New("Velg en JSON-fil eller lim inn JSON først")
	}

	var bundle model.ThemeBundle
	if err := json.Unmarshal(rawJSON, &bundle); err != nil {
		return model.ThemeBundle{}, fmt.Errorf("Ugyldig JSON: %w", err)
	}
	return bundle, nil
}

func (s *Server) handleAdminTenantImportCreate(w http.ResponseWriter, r *http.Request) {
	bundle, err := readThemeBundleFromMultipart(w, r)
	if err != nil {
		redirectAdmin(w, r, 0, "", err.Error())
		return
	}

	t, created, err := s.store.UpsertTenantFromBundle(r.Context(), s.defaultCfg, bundle)
	if err != nil {
		redirectAdmin(w, r, 0, "", err.Error())
		return
	}

	_ = s.calMgr.ReloadTenants(r.Context())
	s.calMgr.SyncTenantNow(r.Context(), t)
	slog.Info("Imported tenant via Admin UI", "slug", t.Slug, "created", created)
	flash := "Tenant oppdatert fra JSON"
	if created {
		flash = "Tenant importert fra JSON"
	}
	redirectAdmin(w, r, t.ID, flash, "")
}

func (s *Server) handleAdminThemeExport(w http.ResponseWriter, r *http.Request) {
	id, err := strconv.ParseInt(chi.URLParam(r, "id"), 10, 64)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Ugyldig tenant ID")
		return
	}
	t, err := s.store.GetTenantByID(r.Context(), id)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Fant ikke tenant")
		return
	}

	bundle, err := s.store.BuildThemeBundle(r.Context(), t)
	if err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	}

	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.Header().Set("Content-Disposition", fmt.Sprintf(`attachment; filename="%s-theme.json"`, t.Slug))
	enc := json.NewEncoder(w)
	enc.SetIndent("", "  ")
	_ = enc.Encode(bundle)
}

func (s *Server) handleAdminThemeImport(w http.ResponseWriter, r *http.Request) {
	id, err := strconv.ParseInt(chi.URLParam(r, "id"), 10, 64)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Ugyldig tenant ID")
		return
	}
	t, err := s.store.GetTenantByID(r.Context(), id)
	if err != nil {
		redirectAdmin(w, r, 0, "", "Fant ikke tenant")
		return
	}

	bundle, err := readThemeBundleFromMultipart(w, r)
	if err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	}

	includeMetadata := r.FormValue("include_metadata") == "1"
	if err := s.store.ApplyThemeBundle(r.Context(), &t, bundle, includeMetadata); err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	}
	if err := s.store.UpdateTenant(r.Context(), &t, true); err != nil {
		redirectAdmin(w, r, id, "", err.Error())
		return
	}

	_ = s.calMgr.ReloadTenants(r.Context())
	s.calMgr.SyncTenantNow(r.Context(), t)
	slog.Info("Imported theme via Admin UI", "slug", t.Slug, "id", t.ID, "include_metadata", includeMetadata)
	redirectAdmin(w, r, id, "Tema importert fra JSON", "")
}

func (s *Server) requireAdminAPI(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		// 1. Bearer token (ADMIN_API_TOKEN)
		if authHdr := strings.TrimSpace(r.Header.Get("Authorization")); strings.HasPrefix(strings.ToLower(authHdr), "bearer ") {
			token := strings.TrimSpace(authHdr[len("bearer "):])
			if s.adminAPIToken != "" && subtle.ConstantTimeCompare([]byte(token), []byte(s.adminAPIToken)) == 1 {
				ctx := context.WithValue(r.Context(), adminUserCtxKey, "api-token")
				next.ServeHTTP(w, r.WithContext(ctx))
				return
			}
			writeJSON(w, http.StatusUnauthorized, map[string]string{"error": "Invalid bearer token"})
			return
		}

		// 2. HTTP Basic Auth against admin_users
		if user, pass, ok := r.BasicAuth(); ok {
			adminUser, err := s.store.VerifyAdmin(r.Context(), user, pass)
			if err == nil {
				ctx := context.WithValue(r.Context(), adminUserCtxKey, adminUser.Username)
				next.ServeHTTP(w, r.WithContext(ctx))
				return
			}
			writeJSON(w, http.StatusUnauthorized, map[string]string{"error": "Invalid admin credentials"})
			return
		}

		// 3. Valid admin session cookie
		if cookie, err := r.Cookie(adminCookieName); err == nil {
			if username, ok := s.verifySessionToken(cookie.Value); ok {
				ctx := context.WithValue(r.Context(), adminUserCtxKey, username)
				next.ServeHTTP(w, r.WithContext(ctx))
				return
			}
		}

		w.Header().Set("WWW-Authenticate", `Basic realm="MeetEasier Admin API"`)
		writeJSON(w, http.StatusUnauthorized, map[string]string{"error": "Unauthorized"})
	})
}

func (s *Server) handleAPIAdminListTenants(w http.ResponseWriter, r *http.Request) {
	tenants, err := s.store.ListTenants(r.Context())
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}
	writeJSON(w, http.StatusOK, tenants)
}

func (s *Server) handleAPIAdminUpsertTenant(w http.ResponseWriter, r *http.Request) {
	r.Body = http.MaxBytesReader(w, r.Body, 10<<20)
	var bundle model.ThemeBundle
	if err := json.NewDecoder(r.Body).Decode(&bundle); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": "Invalid JSON: " + err.Error()})
		return
	}

	t, isNew, err := s.store.UpsertTenantFromBundle(r.Context(), s.defaultCfg, bundle)
	if err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": err.Error()})
		return
	}

	_ = s.calMgr.ReloadTenants(r.Context())
	s.calMgr.SyncTenantNow(r.Context(), t)
	slog.Info("Upserted tenant via Admin API", "slug", t.Slug, "created", isNew)

	status := http.StatusOK
	if isNew {
		status = http.StatusCreated
	}
	writeJSON(w, status, t)
}

func (s *Server) handleAPIAdminExportTheme(w http.ResponseWriter, r *http.Request) {
	slug := chi.URLParam(r, "slug")
	t, err := s.store.GetTenantBySlug(r.Context(), slug)
	if errors.Is(err, store.ErrNotFound) {
		writeJSON(w, http.StatusNotFound, map[string]string{"error": "Tenant not found"})
		return
	}
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}

	bundle, err := s.store.BuildThemeBundle(r.Context(), t)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}
	writeJSON(w, http.StatusOK, bundle)
}

func (s *Server) handleAPIAdminImportTheme(w http.ResponseWriter, r *http.Request) {
	slug := chi.URLParam(r, "slug")
	t, err := s.store.GetTenantBySlug(r.Context(), slug)
	if errors.Is(err, store.ErrNotFound) {
		writeJSON(w, http.StatusNotFound, map[string]string{"error": "Tenant not found"})
		return
	}
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}

	r.Body = http.MaxBytesReader(w, r.Body, 10<<20)
	var bundle model.ThemeBundle
	if err := json.NewDecoder(r.Body).Decode(&bundle); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": "Invalid JSON: " + err.Error()})
		return
	}

	if err := s.store.ApplyThemeBundle(r.Context(), &t, bundle, false); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": err.Error()})
		return
	}
	if err := s.store.UpdateTenant(r.Context(), &t, true); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": err.Error()})
		return
	}

	_ = s.calMgr.ReloadTenants(r.Context())
	s.calMgr.SyncTenantNow(r.Context(), t)
	slog.Info("Imported theme via Admin API", "slug", t.Slug)
	writeJSON(w, http.StatusOK, t)
}

func redirectAdmin(w http.ResponseWriter, r *http.Request, tenantID int64, flash, errMsg string) {
	q := url.Values{}
	if tenantID > 0 {
		q.Set("tenant_id", strconv.FormatInt(tenantID, 10))
	}
	if flash != "" {
		q.Set("flash", flash)
	}
	if errMsg != "" {
		q.Set("error", errMsg)
	}
	target := "/admin"
	if encoded := q.Encode(); encoded != "" {
		target += "?" + encoded
	}
	http.Redirect(w, r, target, http.StatusSeeOther)
}

func parseIntDefault(raw string, def int) int {
	v, err := strconv.Atoi(strings.TrimSpace(raw))
	if err != nil || v <= 0 {
		return def
	}
	return v
}

