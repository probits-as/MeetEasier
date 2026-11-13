package store

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"log/slog"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/probits-as/meeteasier/internal/config"
	"github.com/probits-as/meeteasier/internal/model"
)

// MaxAssetUploadBytes limits individual logo/image assets to 2 MiB.
const MaxAssetUploadBytes = 2 << 20

// ValidateAndSaveAssetBytes validates an uploaded or bundled image asset and saves it in SQLite.
func (s *Store) ValidateAndSaveAssetBytes(ctx context.Context, slug, assetKey string, data []byte) error {
	if len(data) == 0 {
		return nil
	}
	if len(data) > MaxAssetUploadBytes {
		return fmt.Errorf("asset %s er for stor (maks 2 MB)", assetKey)
	}

	contentType := http.DetectContentType(data)
	if !strings.HasPrefix(contentType, "image/") {
		trimmed := strings.TrimSpace(string(data))
		if strings.Contains(trimmed, "<svg") {
			contentType = "image/svg+xml"
		} else {
			return fmt.Errorf("ugyldig filtype for %s (%s) — kun bilder tillatt", assetKey, contentType)
		}
	}

	return s.UpsertAsset(ctx, slug, assetKey, contentType, data)
}

// BuildThemeBundle exports a tenant's theme, labels, custom CSS, and binary image assets into a portable JSON bundle.
// OAuth secrets are intentionally excluded so exported bundles are safe to store in Git.
func (s *Store) BuildThemeBundle(ctx context.Context, t model.Tenant) (model.ThemeBundle, error) {
	assets, err := s.ListAssetsByTenant(ctx, t.Slug)
	if err != nil {
		return model.ThemeBundle{}, err
	}

	bundle := model.ThemeBundle{
		Version:             1,
		Slug:                t.Slug,
		Name:                t.Name,
		Domains:             t.Domains,
		MaxDays:             t.MaxDays,
		MaxRoomLists:        t.MaxRoomLists,
		MaxRooms:            t.MaxRooms,
		MaxItems:            t.MaxItems,
		RoomBlacklist:       t.RoomBlacklist,
		PollIntervalSeconds: t.PollIntervalSeconds,
		PrivacyMode:         t.PrivacyMode,
		ThemePreset:         t.ThemePreset,
		Theme:               t.Theme,
		CustomCSS:           t.CustomCSS,
	}
	if t.IsDefault {
		isDef := true
		bundle.IsDefault = &isDef
	}
	if len(assets) > 0 {
		bundle.Assets = make(map[string]model.ThemeAssetData, len(assets))
		for _, a := range assets {
			bundle.Assets[a.AssetKey] = model.ThemeAssetData{
				ContentType: a.ContentType,
				DataBase64:  base64.StdEncoding.EncodeToString(a.Data),
			}
		}
	}
	return bundle, nil
}

// ApplyThemeBundle applies theme styling, labels, custom CSS, assets, and optional metadata onto a Tenant.
func (s *Store) ApplyThemeBundle(ctx context.Context, t *model.Tenant, bundle model.ThemeBundle, includeMetadata bool) error {
	if includeMetadata {
		if strings.TrimSpace(bundle.Name) != "" {
			t.Name = strings.TrimSpace(bundle.Name)
		}
		if strings.TrimSpace(bundle.Domains) != "" {
			t.Domains = strings.TrimSpace(bundle.Domains)
		}
		if bundle.IsDefault != nil {
			t.IsDefault = *bundle.IsDefault
		}
	}
	if strings.TrimSpace(bundle.PrivacyMode) != "" {
		t.PrivacyMode = strings.TrimSpace(bundle.PrivacyMode)
	}
	if strings.TrimSpace(bundle.ThemePreset) != "" {
		t.ThemePreset = strings.TrimSpace(bundle.ThemePreset)
	} else {
		t.ThemePreset = model.PresetCustom
	}

	// Merge onto preset defaults so omitted JSON fields keep valid fallbacks
	base := model.ThemeForPreset(t.ThemePreset)
	merged := bundle.Theme
	if merged.RowLayout == "" {
		merged.RowLayout = base.RowLayout
	}
	if merged.SingleRoomStatusFontSize == "" {
		merged.SingleRoomStatusFontSize = base.SingleRoomStatusFontSize
	}
	if merged.BodyBg == "" {
		merged.BodyBg = base.BodyBg
	}
	if merged.CardBg == "" {
		merged.CardBg = base.CardBg
	}
	if merged.CardBorder == "" {
		merged.CardBorder = base.CardBorder
	}
	if merged.TitleBarBg == "" {
		merged.TitleBarBg = base.TitleBarBg
	}
	if merged.PrimaryFont == "" {
		merged.PrimaryFont = base.PrimaryFont
	}
	if merged.OpenColor == "" {
		merged.OpenColor = base.OpenColor
	}
	if merged.OpenDarkColor == "" {
		merged.OpenDarkColor = base.OpenDarkColor
	}
	if merged.BusyColor == "" {
		merged.BusyColor = base.BusyColor
	}
	if merged.UpcomingColor == "" {
		merged.UpcomingColor = base.UpcomingColor
	}
	if merged.ErrorColor == "" {
		merged.ErrorColor = base.ErrorColor
	}
	t.Theme = merged
	t.CustomCSS = bundle.CustomCSS

	versionSuffix := time.Now().Unix()
	for key, asset := range bundle.Assets {
		cleanKey := model.Slugify(key)
		if cleanKey == "" || strings.TrimSpace(asset.DataBase64) == "" {
			continue
		}
		b64 := strings.TrimSpace(asset.DataBase64)
		if idx := strings.Index(b64, ";base64,"); idx != -1 {
			b64 = b64[idx+len(";base64,"):]
		}
		raw, err := base64.StdEncoding.DecodeString(b64)
		if err != nil {
			return fmt.Errorf("ugyldig base64 for asset %s: %w", key, err)
		}
		if err := s.ValidateAndSaveAssetBytes(ctx, t.Slug, cleanKey, raw); err != nil {
			return err
		}
		assetURL := fmt.Sprintf("/tenant-assets/%s/%s?v=%d", url.PathEscape(t.Slug), url.PathEscape(cleanKey), versionSuffix)
		switch cleanKey {
		case "logo":
			t.Theme.LogoURL = assetURL
		case "footer":
			t.Theme.FooterImageURL = assetURL
		case "singleroom-logo", "singleroom_logo":
			t.Theme.SingleRoomLogoURL = assetURL
		}
	}

	return nil
}

// HasValidOAuthEnv returns true if usable Microsoft Graph OAuth credentials are present in AppConfig.
func HasValidOAuthEnv(cfg config.AppConfig) bool {
	return cfg.OAuthClientID != "" &&
		cfg.OAuthClientID != templateClientID &&
		cfg.OAuthClientID != "OAUTH_CLIENT_ID_NOT_SET" &&
		cfg.OAuthClientSecret != "" &&
		cfg.OAuthAuthority != ""
}

// UpsertTenantFromBundle creates or updates a tenant from a ThemeBundle and inherits OAuth/search settings from AppConfig when omitted.
func (s *Store) UpsertTenantFromBundle(ctx context.Context, cfg config.AppConfig, bundle model.ThemeBundle) (model.Tenant, bool, error) {
	slug := model.Slugify(bundle.Slug)
	if slug == "" {
		slug = model.Slugify(bundle.Name)
	}
	if slug == "" {
		return model.Tenant{}, false, errors.New("slug or name is required")
	}

	existing, err := s.GetTenantBySlug(ctx, slug)
	isNew := errors.Is(err, ErrNotFound)
	if err != nil && !isNew {
		return model.Tenant{}, false, err
	}

	var t model.Tenant
	if isNew {
		maxDays := cfg.MaxDays
		if maxDays <= 0 {
			maxDays = 10
		}
		maxRoomLists := cfg.MaxRoomLists
		if maxRoomLists <= 0 {
			maxRoomLists = 10
		}
		maxRooms := cfg.MaxRooms
		if maxRooms <= 0 {
			maxRooms = 20
		}
		maxItems := cfg.MaxItems
		if maxItems <= 0 {
			maxItems = 6
		}
		pollSec := cfg.PollIntervalSec
		if pollSec < 10 {
			pollSec = 60
		}

		t = model.Tenant{
			Slug:                slug,
			Name:                bundle.Name,
			Domains:             bundle.Domains,
			ProviderType:        model.ProviderMock,
			ExchangeDomain:      cfg.Domain,
			MaxDays:             maxDays,
			MaxRoomLists:        maxRoomLists,
			MaxRooms:            maxRooms,
			MaxItems:            maxItems,
			PrivacyMode:         model.PrivacyNormal,
			ThemePreset:         model.PresetStandard,
			Theme:               model.DefaultStandardTheme(),
			PollIntervalSeconds: pollSec,
		}
		if t.Name == "" {
			t.Name = slug
		}
	} else {
		t = existing
	}

	if err := s.ApplyThemeBundle(ctx, &t, bundle, true); err != nil {
		return model.Tenant{}, false, err
	}

	// Optional provider, OAuth metadata, and limits from bundle
	explicitProvider := strings.ToLower(strings.TrimSpace(bundle.ProviderType))
	if explicitProvider == model.ProviderMSGraph || explicitProvider == model.ProviderMock {
		t.ProviderType = explicitProvider
	}
	if v := strings.TrimSpace(bundle.OAuthClientID); v != "" {
		t.OAuthClientID = v
	}
	if v := strings.TrimSpace(bundle.OAuthAuthority); v != "" {
		t.OAuthAuthority = v
	}
	if v := strings.TrimSpace(bundle.ExchangeDomain); v != "" {
		t.ExchangeDomain = v
	}
	if strings.TrimSpace(bundle.RoomBlacklist) != "" {
		t.RoomBlacklist = bundle.RoomBlacklist
	}
	if bundle.MaxDays > 0 {
		t.MaxDays = bundle.MaxDays
	}
	if bundle.MaxRoomLists > 0 {
		t.MaxRoomLists = bundle.MaxRoomLists
	}
	if bundle.MaxRooms > 0 {
		t.MaxRooms = bundle.MaxRooms
	}
	if bundle.MaxItems > 0 {
		t.MaxItems = bundle.MaxItems
	}
	if bundle.PollIntervalSeconds >= 10 {
		t.PollIntervalSeconds = bundle.PollIntervalSeconds
	}

	// Ensure at least one default tenant exists if database starts from a bootstrap bundle without is_default
	if isNew && bundle.IsDefault == nil {
		if _, err := s.GetDefaultTenant(ctx); errors.Is(err, ErrNotFound) {
			t.IsDefault = true
		}
	}

	// Inherit OAuth credentials from environment variables (e.g. Kubernetes ExternalSecret)
	if HasValidOAuthEnv(cfg) && (t.IsDefault || t.Slug == "default" || t.ProviderType == model.ProviderMSGraph) {
		if t.OAuthClientID == "" {
			t.OAuthClientID = cfg.OAuthClientID
		}
		if t.OAuthAuthority == "" {
			t.OAuthAuthority = cfg.OAuthAuthority
		}
		if t.OAuthClientSecret == "" {
			t.OAuthClientSecret = cfg.OAuthClientSecret
		}
		if explicitProvider == "" && t.OAuthClientID != "" && t.OAuthAuthority != "" && t.OAuthClientSecret != "" {
			t.ProviderType = model.ProviderMSGraph
		}
	}
	if t.ExchangeDomain == "" && cfg.Domain != "" {
		t.ExchangeDomain = cfg.Domain
	}

	if isNew {
		if err := s.CreateTenant(ctx, &t); err != nil {
			return model.Tenant{}, false, err
		}
	} else {
		if err := s.UpdateTenant(ctx, &t, true); err != nil {
			return model.Tenant{}, false, err
		}
	}

	return t, isNew, nil
}

// BootstrapBundles imports any configured JSON theme/tenant bundles from BOOTSTRAP_DIR, BOOTSTRAP_THEME_FILE, and BOOTSTRAP_THEME_JSON.
func (s *Store) BootstrapBundles(ctx context.Context, cfg config.AppConfig) error {
	var files []string

	if dir := strings.TrimSpace(cfg.BootstrapDir); dir != "" {
		entries, err := os.ReadDir(dir)
		if err != nil {
			if !errors.Is(err, os.ErrNotExist) {
				return fmt.Errorf("read bootstrap_dir %q: %w", dir, err)
			}
		} else {
			for _, entry := range entries {
				name := entry.Name()
				if strings.HasPrefix(name, ".") || !strings.HasSuffix(strings.ToLower(name), ".json") {
					continue
				}
				fullPath := filepath.Join(dir, name)
				info, statErr := os.Stat(fullPath)
				if statErr != nil || info.IsDir() {
					continue
				}
				files = append(files, fullPath)
			}
		}
	}

	if fileList := strings.TrimSpace(cfg.BootstrapThemeFile); fileList != "" {
		for _, part := range strings.Split(fileList, ",") {
			p := strings.TrimSpace(part)
			if p != "" {
				files = append(files, p)
			}
		}
	}

	for _, filePath := range files {
		raw, err := os.ReadFile(filePath)
		if err != nil {
			return fmt.Errorf("read bootstrap file %q: %w", filePath, err)
		}
		if err := s.importRawBundle(ctx, cfg, raw, filePath); err != nil {
			return fmt.Errorf("bootstrap file %q: %w", filePath, err)
		}
	}

	if rawEnv := strings.TrimSpace(cfg.BootstrapThemeJSON); rawEnv != "" {
		rawBytes, err := decodeEnvBundleJSON(rawEnv)
		if err != nil {
			return fmt.Errorf("decode BOOTSTRAP_THEME_JSON: %w", err)
		}
		if err := s.importRawBundle(ctx, cfg, rawBytes, "BOOTSTRAP_THEME_JSON"); err != nil {
			return fmt.Errorf("BOOTSTRAP_THEME_JSON: %w", err)
		}
	}

	return nil
}

func (s *Store) importRawBundle(ctx context.Context, cfg config.AppConfig, raw []byte, source string) error {
	var bundle model.ThemeBundle
	if err := json.Unmarshal(raw, &bundle); err != nil {
		return fmt.Errorf("invalid JSON: %w", err)
	}
	t, created, err := s.UpsertTenantFromBundle(ctx, cfg, bundle)
	if err != nil {
		return err
	}
	slog.Info("Bootstrapped tenant from bundle",
		"source", source,
		"slug", t.Slug,
		"created", created,
		"is_default", t.IsDefault,
		"provider", t.ProviderType,
	)
	return nil
}

func decodeEnvBundleJSON(raw string) ([]byte, error) {
	trimmed := strings.TrimSpace(raw)
	if strings.HasPrefix(trimmed, "{") {
		return []byte(trimmed), nil
	}
	if decoded, err := base64.StdEncoding.DecodeString(trimmed); err == nil {
		return decoded, nil
	}
	if decoded, err := base64.RawStdEncoding.DecodeString(trimmed); err == nil {
		return decoded, nil
	}
	return []byte(trimmed), nil
}

