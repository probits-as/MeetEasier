package server

import (
	"bytes"
	"context"
	"encoding/base64"
	"encoding/json"
	"image"
	"image/color"
	"image/png"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/probits-as/meeteasier/internal/calendar"
	"github.com/probits-as/meeteasier/internal/config"
	"github.com/probits-as/meeteasier/internal/model"
	"github.com/probits-as/meeteasier/internal/store"
)

func setupTestServer(t *testing.T) (*Server, *store.Store, http.Handler) {
	t.Helper()
	cfg := config.AppConfig{
		DBPath:          ":memory:",
		AdminUsername:   "admin",
		AdminPassword:   "secret123",
		AdminAPIToken:   "test-bearer-token",
		SessionSecret:   "test-session-secret",
		PollIntervalSec: 60,
		Domain:          "example.com",
		MaxDays:         10,
		MaxRoomLists:    10,
		MaxRooms:        20,
		MaxItems:        6,
	}
	st, err := store.Open(context.Background(), cfg)
	if err != nil {
		t.Fatalf("store.Open failed: %v", err)
	}
	t.Cleanup(func() { _ = st.Close() })

	calMgr := calendar.NewManager(st, calendar.NewMSGraphProvider(), calendar.NewMockProvider())
	if err := calMgr.ReloadTenants(context.Background()); err != nil {
		t.Fatalf("ReloadTenants failed: %v", err)
	}

	srv := New(st, calMgr, []byte(cfg.SessionSecret), cfg.AdminAPIToken, "../../static")
	return srv, st, srv.Router()
}

func tinyPNG(t *testing.T) []byte {
	t.Helper()
	var buf bytes.Buffer
	img := image.NewRGBA(image.Rect(0, 0, 2, 2))
	img.Set(0, 0, color.RGBA{R: 255, A: 255})
	if err := png.Encode(&buf, img); err != nil {
		t.Fatalf("png.Encode: %v", err)
	}
	return buf.Bytes()
}

func TestMultiTenantFlightboardAndSingleRoom(t *testing.T) {
	_, st, handler := setupTestServer(t)

	// Seed a second compact-theme tenant for multi-tenant host resolution tests
	compactTenant := model.Tenant{
		Slug:                "nord",
		Name:                "Nordic Møterom",
		Domains:             "nord.localhost,rooms.nord.example",
		ProviderType:        model.ProviderMock,
		PrivacyMode:         model.PrivacyOrganizerAsSubject,
		ThemePreset:         model.PresetCompact,
		Theme:               model.DefaultCompactTheme(),
		PollIntervalSeconds: 60,
	}
	if err := st.CreateTenant(context.Background(), &compactTenant); err != nil {
		t.Fatalf("CreateTenant failed: %v", err)
	}

	// 1. Default tenant (Standard)
	req := httptest.NewRequest(http.MethodGet, "/", nil)
	req.Host = "localhost:8080"
	rec := httptest.NewRecorder()
	handler.ServeHTTP(rec, req)
	if rec.Code != http.StatusOK {
		t.Fatalf("expected 200 for default tenant, got %d", rec.Code)
	}
	if !strings.Contains(rec.Body.String(), "Conference Room Availability") {
		t.Fatalf("expected Standard navbar title in response")
	}

	// 2. Compact tenant resolved via custom Host header
	reqCompact := httptest.NewRequest(http.MethodGet, "/", nil)
	reqCompact.Host = "rooms.nord.example"
	recCompact := httptest.NewRecorder()
	handler.ServeHTTP(recCompact, reqCompact)
	if recCompact.Code != http.StatusOK {
		t.Fatalf("expected 200 for compact host, got %d", recCompact.Code)
	}
	compactBody := recCompact.Body.String()
	if !strings.Contains(compactBody, "Møterom") || !strings.Contains(compactBody, "meeting-room-vertical") {
		t.Fatalf("expected compact layout and Norwegian title")
	}

	// 3. Compact tenant resolved via ?tenant=nord query param
	reqQuery := httptest.NewRequest(http.MethodGet, "/?tenant=nord", nil)
	recQuery := httptest.NewRecorder()
	handler.ServeHTTP(recQuery, reqQuery)
	if recQuery.Code != http.StatusOK || !strings.Contains(recQuery.Body.String(), "Møterom") {
		t.Fatalf("expected ?tenant=nord override to render compact tenant")
	}

	// 4. Single room page
	reqRoom := httptest.NewRequest(http.MethodGet, "/single-room/fjord?tenant=nord", nil)
	recRoom := httptest.NewRecorder()
	handler.ServeHTTP(recRoom, reqRoom)
	if recRoom.Code != http.StatusOK || !strings.Contains(recRoom.Body.String(), "Fjord") || !strings.Contains(recRoom.Body.String(), "Opptatt") {
		t.Fatalf("expected single room Fjord in compact Norwegian layout")
	}

	// 5. Partials and API endpoints
	for _, path := range []string{"/partials/flightboard", "/partials/single-room/fjord", "/api/heartbeat", "/api/rooms", "/api/roomlists"} {
		r := httptest.NewRequest(http.MethodGet, path, nil)
		w := httptest.NewRecorder()
		handler.ServeHTTP(w, r)
		if w.Code != http.StatusOK {
			t.Fatalf("expected 200 on %s, got %d", path, w.Code)
		}
	}

	// Verify /api/rooms JSON structure
	rAPI := httptest.NewRequest(http.MethodGet, "/api/rooms", nil)
	wAPI := httptest.NewRecorder()
	handler.ServeHTTP(wAPI, rAPI)
	var rooms []model.Room
	if err := json.Unmarshal(wAPI.Body.Bytes(), &rooms); err != nil {
		t.Fatalf("unmarshal /api/rooms: %v", err)
	}
	if len(rooms) == 0 {
		t.Fatalf("expected mock rooms in /api/rooms")
	}
}

func TestAdminAuthTenantCRUDAndThemeExportImport(t *testing.T) {
	_, st, handler := setupTestServer(t)

	// 1. Unauthenticated /admin redirects to /admin/login
	req := httptest.NewRequest(http.MethodGet, "/admin", nil)
	rec := httptest.NewRecorder()
	handler.ServeHTTP(rec, req)
	if rec.Code != http.StatusSeeOther || rec.Header().Get("Location") != "/admin/login" {
		t.Fatalf("expected redirect to /admin/login, got %d -> %s", rec.Code, rec.Header().Get("Location"))
	}

	// 2. Bad login returns 401
	badForm := url.Values{"username": {"admin"}, "password": {"wrong"}}
	reqBad := httptest.NewRequest(http.MethodPost, "/admin/login", strings.NewReader(badForm.Encode()))
	reqBad.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	recBad := httptest.NewRecorder()
	handler.ServeHTTP(recBad, reqBad)
	if recBad.Code != http.StatusUnauthorized {
		t.Fatalf("expected 401 on wrong password, got %d", recBad.Code)
	}

	// 3. Valid login returns session cookie
	goodForm := url.Values{"username": {"admin"}, "password": {"secret123"}}
	reqGood := httptest.NewRequest(http.MethodPost, "/admin/login", strings.NewReader(goodForm.Encode()))
	reqGood.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	recGood := httptest.NewRecorder()
	handler.ServeHTTP(recGood, reqGood)
	if recGood.Code != http.StatusSeeOther {
		t.Fatalf("expected 303 on login success, got %d", recGood.Code)
	}
	cookies := recGood.Result().Cookies()
	if len(cookies) == 0 {
		t.Fatalf("expected session cookie on login")
	}
	sessionCookie := cookies[0]

	// 4. Create new tenant with compact preset
	createForm := url.Values{
		"name":         {"Acme AS"},
		"slug":         {"acme"},
		"domains":      {"rooms.acme.no"},
		"theme_preset": {"compact"},
	}
	reqCreate := httptest.NewRequest(http.MethodPost, "/admin/tenants", strings.NewReader(createForm.Encode()))
	reqCreate.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	reqCreate.AddCookie(sessionCookie)
	recCreate := httptest.NewRecorder()
	handler.ServeHTTP(recCreate, reqCreate)
	if recCreate.Code != http.StatusSeeOther {
		t.Fatalf("expected 303 on tenant create, got %d", recCreate.Code)
	}

	acme, err := st.GetTenantBySlug(context.Background(), "acme")
	if err != nil {
		t.Fatalf("expected acme tenant in store: %v", err)
	}

	// 5. Update tenant with custom color + uploaded PNG logo
	var buf bytes.Buffer
	mw := multipart.NewWriter(&buf)
	_ = mw.WriteField("name", "Acme AS")
	_ = mw.WriteField("slug", "acme")
	_ = mw.WriteField("domains", "rooms.acme.no")
	_ = mw.WriteField("provider_type", "mock")
	_ = mw.WriteField("privacy_mode", "normal")
	_ = mw.WriteField("row_layout", "compact")
	_ = mw.WriteField("body_bg", "#102030")
	_ = mw.WriteField("label_navbar_title", "Acme Møterom")
	_ = mw.WriteField("custom_css", ".title-bar { border-bottom-width: 3px; }")

	fw, err := mw.CreateFormFile("logo_file", "logo.png")
	if err != nil {
		t.Fatalf("CreateFormFile: %v", err)
	}
	_, _ = fw.Write(tinyPNG(t))
	_ = mw.Close()

	reqUpdate := httptest.NewRequest(http.MethodPost, "/admin/tenants/"+itoa(acme.ID), &buf)
	reqUpdate.Header.Set("Content-Type", mw.FormDataContentType())
	reqUpdate.AddCookie(sessionCookie)
	recUpdate := httptest.NewRecorder()
	handler.ServeHTTP(recUpdate, reqUpdate)
	if recUpdate.Code != http.StatusSeeOther {
		t.Fatalf("expected 303 on tenant update, got %d", recUpdate.Code)
	}

	// 6. Export theme JSON from Admin UI
	reqExport := httptest.NewRequest(http.MethodGet, "/admin/tenants/"+itoa(acme.ID)+"/theme-export", nil)
	reqExport.AddCookie(sessionCookie)
	recExport := httptest.NewRecorder()
	handler.ServeHTTP(recExport, reqExport)
	if recExport.Code != http.StatusOK {
		t.Fatalf("expected 200 on theme export, got %d", recExport.Code)
	}
	var exported model.ThemeBundle
	if err := json.Unmarshal(recExport.Body.Bytes(), &exported); err != nil {
		t.Fatalf("unmarshal exported theme: %v", err)
	}
	if exported.Theme.BodyBg != "#102030" || exported.Assets["logo"].DataBase64 == "" {
		t.Fatalf("expected exported bundle to contain custom color and base64 logo asset")
	}

	// 7. Verify custom domain renders custom color, title, and uploaded logo URL
	reqCustom := httptest.NewRequest(http.MethodGet, "/", nil)
	reqCustom.Host = "rooms.acme.no"
	recCustom := httptest.NewRecorder()
	handler.ServeHTTP(recCustom, reqCustom)
	customHTML := recCustom.Body.String()
	if !strings.Contains(customHTML, "#102030") ||
		!strings.Contains(customHTML, "Acme Møterom") ||
		!strings.Contains(customHTML, "/tenant-assets/acme/logo") {
		t.Fatalf("expected updated theme and uploaded logo on custom domain")
	}

	// 8. Verify uploaded asset endpoint serves image/png
	reqAsset := httptest.NewRequest(http.MethodGet, "/tenant-assets/acme/logo", nil)
	recAsset := httptest.NewRecorder()
	handler.ServeHTTP(recAsset, reqAsset)
	if recAsset.Code != http.StatusOK || recAsset.Header().Get("Content-Type") != "image/png" {
		t.Fatalf("expected 200 image/png from tenant asset, got %d (%s)", recAsset.Code, recAsset.Header().Get("Content-Type"))
	}
}

func TestAuthorizedAdminAPIUpsertAndThemeImportExport(t *testing.T) {
	_, _, handler := setupTestServer(t)

	// 1. Unauthenticated API request returns 401
	reqUnauth := httptest.NewRequest(http.MethodGet, "/api/admin/tenants", nil)
	recUnauth := httptest.NewRecorder()
	handler.ServeHTTP(recUnauth, reqUnauth)
	if recUnauth.Code != http.StatusUnauthorized {
		t.Fatalf("expected 401 on unauthenticated /api/admin/tenants, got %d", recUnauth.Code)
	}

	// 2. Invalid Bearer token returns 401
	reqBadTok := httptest.NewRequest(http.MethodGet, "/api/admin/tenants", nil)
	reqBadTok.Header.Set("Authorization", "Bearer wrong-token")
	recBadTok := httptest.NewRecorder()
	handler.ServeHTTP(recBadTok, reqBadTok)
	if recBadTok.Code != http.StatusUnauthorized {
		t.Fatalf("expected 401 on wrong bearer token, got %d", recBadTok.Code)
	}

	// 3. Upsert a new tenant + theme + base64 image assets via Bearer token
	pngB64 := base64.StdEncoding.EncodeToString(tinyPNG(t))
	bundle := model.ThemeBundle{
		Version:     1,
		Slug:        "custom-brand",
		Name:        "Custom Brand Display",
		Domains:     "display.custom.example",
		PrivacyMode: model.PrivacyOrganizerAsSubject,
		ThemePreset: model.PresetCompact,
		Theme:       model.DefaultCompactTheme(),
		CustomCSS:   ".title-bar { letter-spacing: 0.02em; }",
		Assets: map[string]model.ThemeAssetData{
			"logo":            {ContentType: "image/png", DataBase64: pngB64},
			"footer":          {ContentType: "image/png", DataBase64: pngB64},
			"singleroom_logo": {ContentType: "image/png", DataBase64: pngB64},
		},
	}
	bundle.Theme.BodyBg = "#221133"
	payload, _ := json.Marshal(bundle)

	reqUpsert := httptest.NewRequest(http.MethodPost, "/api/admin/tenants", bytes.NewReader(payload))
	reqUpsert.Header.Set("Authorization", "Bearer test-bearer-token")
	reqUpsert.Header.Set("Content-Type", "application/json")
	recUpsert := httptest.NewRecorder()
	handler.ServeHTTP(recUpsert, reqUpsert)
	if recUpsert.Code != http.StatusCreated {
		t.Fatalf("expected 201 on API upsert create, got %d: %s", recUpsert.Code, recUpsert.Body.String())
	}

	// 4. Verify custom domain renders imported theme & assets
	reqDisplay := httptest.NewRequest(http.MethodGet, "/", nil)
	reqDisplay.Host = "display.custom.example"
	recDisplay := httptest.NewRecorder()
	handler.ServeHTTP(recDisplay, reqDisplay)
	html := recDisplay.Body.String()
	if !strings.Contains(html, "#221133") ||
		!strings.Contains(html, "/tenant-assets/custom-brand/logo") ||
		!strings.Contains(html, "/tenant-assets/custom-brand/footer") {
		t.Fatalf("expected imported API theme and assets on custom domain")
	}

	// 5. Export via HTTP Basic Auth and re-import via PUT
	reqExport := httptest.NewRequest(http.MethodGet, "/api/admin/tenants/custom-brand/theme", nil)
	reqExport.SetBasicAuth("admin", "secret123")
	recExport := httptest.NewRecorder()
	handler.ServeHTTP(recExport, reqExport)
	if recExport.Code != http.StatusOK {
		t.Fatalf("expected 200 on BasicAuth theme export, got %d", recExport.Code)
	}

	var reimported model.ThemeBundle
	if err := json.Unmarshal(recExport.Body.Bytes(), &reimported); err != nil {
		t.Fatalf("unmarshal theme export: %v", err)
	}
	reimported.Theme.BodyBg = "#445566"
	putPayload, _ := json.Marshal(reimported)

	reqPut := httptest.NewRequest(http.MethodPut, "/api/admin/tenants/custom-brand/theme", bytes.NewReader(putPayload))
	reqPut.Header.Set("Authorization", "Bearer test-bearer-token")
	reqPut.Header.Set("Content-Type", "application/json")
	recPut := httptest.NewRecorder()
	handler.ServeHTTP(recPut, reqPut)
	if recPut.Code != http.StatusOK {
		t.Fatalf("expected 200 on PUT theme import, got %d: %s", recPut.Code, recPut.Body.String())
	}

	reqDisplay2 := httptest.NewRequest(http.MethodGet, "/", nil)
	reqDisplay2.Host = "display.custom.example"
	recDisplay2 := httptest.NewRecorder()
	handler.ServeHTTP(recDisplay2, reqDisplay2)
	if !strings.Contains(recDisplay2.Body.String(), "#445566") {
		t.Fatalf("expected updated BodyBg #445566 after PUT theme import")
	}
}

func TestBootstrapBundlesAndAdminTenantImport(t *testing.T) {
	pngB64 := base64.StdEncoding.EncodeToString(tinyPNG(t))
	isDef := true
	bundle := model.ThemeBundle{
		Version:     1,
		Slug:        "boot-brand",
		Name:        "Boot Brand Rooms",
		Domains:     "rooms.boot.example",
		IsDefault:   &isDef,
		PrivacyMode: model.PrivacyOrganizerAsSubject,
		ThemePreset: model.PresetCompact,
		Theme:       model.DefaultCompactTheme(),
		Assets: map[string]model.ThemeAssetData{
			"logo": {ContentType: "image/png", DataBase64: pngB64},
		},
	}
	bundle.Theme.BodyBg = "#0b132b"
	bundleBytes, _ := json.Marshal(bundle)

	bootDir := t.TempDir()
	if err := os.WriteFile(filepath.Join(bootDir, "01-brand.json"), bundleBytes, 0o600); err != nil {
		t.Fatalf("write bootstrap file: %v", err)
	}

	envBundle := bundle
	envBundle.Slug = "env-brand"
	envBundle.Name = "Env Brand Rooms"
	envBundle.Domains = "rooms.env.example"
	envIsDef := false
	envBundle.IsDefault = &envIsDef
	envBundleBytes, _ := json.Marshal(envBundle)

	cfg := config.AppConfig{
		DBPath:             filepath.Join(t.TempDir(), "test.db"),
		AdminUsername:      "admin",
		AdminPassword:      "secret123",
		SessionSecret:      "test-session-secret",
		PollIntervalSec:    60,
		BootstrapDir:       bootDir,
		BootstrapThemeJSON: base64.StdEncoding.EncodeToString(envBundleBytes),
		OAuthClientID:      "11111111-2222-3333-4444-555555555555",
		OAuthAuthority:     "https://login.microsoftonline.com/tenant-guid",
		OAuthClientSecret:  "super-secret-value",
		Domain:             "boot.example",
		MaxDays:            7,
		MaxRoomLists:       5,
		MaxRooms:           15,
		MaxItems:           4,
	}

	st, err := store.Open(context.Background(), cfg)
	if err != nil {
		t.Fatalf("store.Open with bootstrap failed: %v", err)
	}
	defer st.Close()

	// Verify only bootstrapped tenants exist (no unused default tenant was created)
	tenants, err := st.ListTenants(context.Background())
	if err != nil {
		t.Fatalf("ListTenants failed: %v", err)
	}
	if len(tenants) != 2 {
		t.Fatalf("expected 2 bootstrapped tenants, got %d", len(tenants))
	}

	defTenant, err := st.GetDefaultTenant(context.Background())
	if err != nil {
		t.Fatalf("GetDefaultTenant failed: %v", err)
	}
	if defTenant.Slug != "boot-brand" {
		t.Fatalf("expected boot-brand to be default tenant, got %s", defTenant.Slug)
	}
	if defTenant.ProviderType != model.ProviderMSGraph || !defTenant.HasClientSecret || defTenant.OAuthClientID != cfg.OAuthClientID {
		t.Fatalf("expected bootstrapped default tenant to inherit OAuth env vars and MSGraph provider, got provider=%s hasSecret=%v", defTenant.ProviderType, defTenant.HasClientSecret)
	}

	envTenant, err := st.GetTenantBySlug(context.Background(), "env-brand")
	if err != nil {
		t.Fatalf("expected env-brand tenant from base64 env bundle: %v", err)
	}
	if envTenant.ProviderType != model.ProviderMock || envTenant.IsDefault {
		t.Fatalf("expected non-default env tenant to remain mock, got provider=%s isDefault=%v", envTenant.ProviderType, envTenant.IsDefault)
	}

	// Also verify Admin UI /admin/tenants/import endpoint
	_, st2, handler2 := setupTestServer(t)
	goodForm := url.Values{"username": {"admin"}, "password": {"secret123"}}
	reqLogin := httptest.NewRequest(http.MethodPost, "/admin/login", strings.NewReader(goodForm.Encode()))
	reqLogin.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	recLogin := httptest.NewRecorder()
	handler2.ServeHTTP(recLogin, reqLogin)
	sessionCookie := recLogin.Result().Cookies()[0]

	var mpBuf bytes.Buffer
	mw := multipart.NewWriter(&mpBuf)
	fw, err := mw.CreateFormFile("theme_file", "boot-brand.json")
	if err != nil {
		t.Fatalf("CreateFormFile: %v", err)
	}
	_, _ = fw.Write(bundleBytes)
	_ = mw.Close()

	reqImport := httptest.NewRequest(http.MethodPost, "/admin/tenants/import", &mpBuf)
	reqImport.Header.Set("Content-Type", mw.FormDataContentType())
	reqImport.AddCookie(sessionCookie)
	recImport := httptest.NewRecorder()
	handler2.ServeHTTP(recImport, reqImport)
	if recImport.Code != http.StatusSeeOther {
		t.Fatalf("expected 303 on /admin/tenants/import, got %d", recImport.Code)
	}

	imported, err := st2.GetTenantBySlug(context.Background(), "boot-brand")
	if err != nil {
		t.Fatalf("expected imported boot-brand tenant: %v", err)
	}
	if !imported.IsDefault || imported.Theme.BodyBg != "#0b132b" {
		t.Fatalf("unexpected imported tenant state: isDefault=%v bodyBg=%s", imported.IsDefault, imported.Theme.BodyBg)
	}
}

func itoa(v int64) string {
	return url.QueryEscape(strings.TrimSpace(
		func() string {
			b, _ := json.Marshal(v)
			return string(b)
		}(),
	))
}
