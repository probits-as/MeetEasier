package store

import (
	"context"
	"crypto/rand"
	"database/sql"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"log/slog"
	"os"
	"path/filepath"
	"strings"
	"time"

	"golang.org/x/crypto/bcrypt"
	_ "modernc.org/sqlite"

	"github.com/probits-as/meeteasier/internal/config"
	"github.com/probits-as/meeteasier/internal/model"
)

var (
	ErrNotFound         = errors.New("not found")
	ErrSlugExists       = errors.New("tenant slug already exists")
	ErrCannotDeleteLast = errors.New("cannot delete the only tenant")
)

const templateClientID = "0c9c8ca8-5a1a-11ed-9b6a-0242ac120002"

// Store manages persistence in SQLite.
type Store struct {
	db *sql.DB
}

// Open initializes the SQLite database, runs migrations, and seeds initial data if empty.
func Open(ctx context.Context, cfg config.AppConfig) (*Store, error) {
	if cfg.DBPath != ":memory:" && !strings.HasPrefix(cfg.DBPath, "file:") {
		dir := filepath.Dir(cfg.DBPath)
		if dir != "" && dir != "." {
			if err := os.MkdirAll(dir, 0o750); err != nil {
				return nil, fmt.Errorf("create db directory: %w", err)
			}
		}
	}

	db, err := sql.Open("sqlite", cfg.DBPath)
	if err != nil {
		return nil, fmt.Errorf("open sqlite: %w", err)
	}
	db.SetMaxOpenConns(1)

	s := &Store{db: db}
	if err := s.initPragmas(ctx, cfg.DBPath); err != nil {
		_ = db.Close()
		return nil, err
	}
	if err := s.migrate(ctx); err != nil {
		_ = db.Close()
		return nil, err
	}
	if err := s.BootstrapBundles(ctx, cfg); err != nil {
		_ = db.Close()
		return nil, err
	}
	if err := s.seed(ctx, cfg); err != nil {
		_ = db.Close()
		return nil, err
	}

	return s, nil
}

// Close closes the underlying database connection.
func (s *Store) Close() error {
	return s.db.Close()
}

func (s *Store) initPragmas(ctx context.Context, dbPath string) error {
	pragmas := []string{
		"PRAGMA busy_timeout = 5000;",
		"PRAGMA foreign_keys = ON;",
	}
	if dbPath != ":memory:" && !strings.HasPrefix(dbPath, "file::memory:") {
		pragmas = append(pragmas, "PRAGMA journal_mode = WAL;")
	}
	for _, p := range pragmas {
		if _, err := s.db.ExecContext(ctx, p); err != nil {
			return fmt.Errorf("exec pragma %q: %w", p, err)
		}
	}
	return nil
}

func (s *Store) migrate(ctx context.Context) error {
	schema := `
CREATE TABLE IF NOT EXISTS tenants (
	id INTEGER PRIMARY KEY AUTOINCREMENT,
	slug TEXT NOT NULL UNIQUE,
	name TEXT NOT NULL,
	domains TEXT NOT NULL DEFAULT '',
	is_default INTEGER NOT NULL DEFAULT 0,
	provider_type TEXT NOT NULL DEFAULT 'mock',
	oauth_client_id TEXT NOT NULL DEFAULT '',
	oauth_authority TEXT NOT NULL DEFAULT '',
	oauth_client_secret TEXT NOT NULL DEFAULT '',
	exchange_domain TEXT NOT NULL DEFAULT '',
	max_days INTEGER NOT NULL DEFAULT 10,
	max_room_lists INTEGER NOT NULL DEFAULT 10,
	max_rooms INTEGER NOT NULL DEFAULT 20,
	max_items INTEGER NOT NULL DEFAULT 6,
	room_blacklist TEXT NOT NULL DEFAULT '',
	privacy_mode TEXT NOT NULL DEFAULT 'normal',
	theme_preset TEXT NOT NULL DEFAULT 'standard',
	theme_json TEXT NOT NULL DEFAULT '{}',
	custom_css TEXT NOT NULL DEFAULT '',
	poll_interval_seconds INTEGER NOT NULL DEFAULT 60,
	created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
	updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS admin_users (
	id INTEGER PRIMARY KEY AUTOINCREMENT,
	username TEXT NOT NULL UNIQUE,
	password_hash TEXT NOT NULL,
	created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS tenant_assets (
	tenant_slug TEXT NOT NULL,
	asset_key TEXT NOT NULL,
	content_type TEXT NOT NULL,
	data BLOB NOT NULL,
	updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
	PRIMARY KEY (tenant_slug, asset_key)
);

CREATE TABLE IF NOT EXISTS app_settings (
	key TEXT PRIMARY KEY,
	value TEXT NOT NULL
);
`
	if _, err := s.db.ExecContext(ctx, schema); err != nil {
		return fmt.Errorf("migrate schema: %w", err)
	}
	return nil
}

func (s *Store) seed(ctx context.Context, cfg config.AppConfig) error {
	var tenantCount int
	if err := s.db.QueryRowContext(ctx, `SELECT COUNT(*) FROM tenants`).Scan(&tenantCount); err != nil {
		return fmt.Errorf("count tenants: %w", err)
	}

	if tenantCount == 0 {
		defaultProvider := model.ProviderMock
		clientID := ""
		authority := ""
		clientSecret := ""
		if HasValidOAuthEnv(cfg) {
			defaultProvider = model.ProviderMSGraph
			clientID = cfg.OAuthClientID
			authority = cfg.OAuthAuthority
			clientSecret = cfg.OAuthClientSecret
		}

		defaultTenant := model.Tenant{
			Slug:                "default",
			Name:                "MeetEasier",
			Domains:             "",
			IsDefault:           true,
			ProviderType:        defaultProvider,
			OAuthClientID:       clientID,
			OAuthAuthority:      authority,
			OAuthClientSecret:   clientSecret,
			ExchangeDomain:      cfg.Domain,
			MaxDays:             cfg.MaxDays,
			MaxRoomLists:        cfg.MaxRoomLists,
			MaxRooms:            cfg.MaxRooms,
			MaxItems:            cfg.MaxItems,
			PrivacyMode:         model.PrivacyNormal,
			ThemePreset:         model.PresetStandard,
			Theme:               model.DefaultStandardTheme(),
			PollIntervalSeconds: cfg.PollIntervalSec,
		}
		if err := s.CreateTenant(ctx, &defaultTenant); err != nil {
			return fmt.Errorf("seed default tenant: %w", err)
		}

		slog.Info("Seeded initial default tenant", "provider", defaultProvider)
	}

	var adminCount int
	if err := s.db.QueryRowContext(ctx, `SELECT COUNT(*) FROM admin_users`).Scan(&adminCount); err != nil {
		return fmt.Errorf("count admin_users: %w", err)
	}
	if adminCount == 0 {
		username := strings.TrimSpace(cfg.AdminUsername)
		if username == "" {
			username = "admin"
		}
		password := cfg.AdminPassword
		if password == "" {
			password = "admin123"
		}
		hash, err := bcrypt.GenerateFromPassword([]byte(password), bcrypt.DefaultCost)
		if err != nil {
			return fmt.Errorf("hash admin password: %w", err)
		}
		if _, err := s.db.ExecContext(ctx,
			`INSERT INTO admin_users (username, password_hash) VALUES (?, ?)`,
			username, string(hash),
		); err != nil {
			return fmt.Errorf("seed admin user: %w", err)
		}
		slog.Info("Seeded initial admin user", "username", username)
	}

	return nil
}

// GetOrCreateSessionSecret returns a persistent HMAC signing secret.
func (s *Store) GetOrCreateSessionSecret(ctx context.Context, configured string) ([]byte, error) {
	if strings.TrimSpace(configured) != "" {
		return []byte(configured), nil
	}

	var existing string
	err := s.db.QueryRowContext(ctx, `SELECT value FROM app_settings WHERE key = 'session_secret'`).Scan(&existing)
	if err == nil && existing != "" {
		return []byte(existing), nil
	}
	if err != nil && !errors.Is(err, sql.ErrNoRows) {
		return nil, fmt.Errorf("read session_secret: %w", err)
	}

	buf := make([]byte, 32)
	if _, err := rand.Read(buf); err != nil {
		return nil, fmt.Errorf("generate session_secret: %w", err)
	}
	secret := hex.EncodeToString(buf)
	if _, err := s.db.ExecContext(ctx, `INSERT OR REPLACE INTO app_settings (key, value) VALUES ('session_secret', ?)`, secret); err != nil {
		return nil, fmt.Errorf("save session_secret: %w", err)
	}
	return []byte(secret), nil
}

// ListTenants returns all configured tenants ordered by default first, then name.
func (s *Store) ListTenants(ctx context.Context) ([]model.Tenant, error) {
	rows, err := s.db.QueryContext(ctx, `
		SELECT id, slug, name, domains, is_default, provider_type,
		       oauth_client_id, oauth_authority, oauth_client_secret, exchange_domain,
		       max_days, max_room_lists, max_rooms, max_items, room_blacklist,
		       privacy_mode, theme_preset, theme_json, custom_css, poll_interval_seconds,
		       created_at, updated_at
		FROM tenants
		ORDER BY is_default DESC, id ASC
	`)
	if err != nil {
		return nil, fmt.Errorf("list tenants: %w", err)
	}
	defer rows.Close()

	var tenants []model.Tenant
	for rows.Next() {
		t, err := scanTenant(rows)
		if err != nil {
			return nil, err
		}
		tenants = append(tenants, t)
	}
	return tenants, rows.Err()
}

// GetTenantByID fetches a tenant by numeric ID.
func (s *Store) GetTenantByID(ctx context.Context, id int64) (model.Tenant, error) {
	row := s.db.QueryRowContext(ctx, `
		SELECT id, slug, name, domains, is_default, provider_type,
		       oauth_client_id, oauth_authority, oauth_client_secret, exchange_domain,
		       max_days, max_room_lists, max_rooms, max_items, room_blacklist,
		       privacy_mode, theme_preset, theme_json, custom_css, poll_interval_seconds,
		       created_at, updated_at
		FROM tenants WHERE id = ?
	`, id)
	t, err := scanTenant(row)
	if errors.Is(err, sql.ErrNoRows) {
		return model.Tenant{}, ErrNotFound
	}
	return t, err
}

// GetTenantBySlug fetches a tenant by slug.
func (s *Store) GetTenantBySlug(ctx context.Context, slug string) (model.Tenant, error) {
	row := s.db.QueryRowContext(ctx, `
		SELECT id, slug, name, domains, is_default, provider_type,
		       oauth_client_id, oauth_authority, oauth_client_secret, exchange_domain,
		       max_days, max_room_lists, max_rooms, max_items, room_blacklist,
		       privacy_mode, theme_preset, theme_json, custom_css, poll_interval_seconds,
		       created_at, updated_at
		FROM tenants WHERE LOWER(slug) = LOWER(?)
	`, strings.TrimSpace(slug))
	t, err := scanTenant(row)
	if errors.Is(err, sql.ErrNoRows) {
		return model.Tenant{}, ErrNotFound
	}
	return t, err
}

// GetDefaultTenant returns the tenant marked as default, or the first tenant.
func (s *Store) GetDefaultTenant(ctx context.Context) (model.Tenant, error) {
	row := s.db.QueryRowContext(ctx, `
		SELECT id, slug, name, domains, is_default, provider_type,
		       oauth_client_id, oauth_authority, oauth_client_secret, exchange_domain,
		       max_days, max_room_lists, max_rooms, max_items, room_blacklist,
		       privacy_mode, theme_preset, theme_json, custom_css, poll_interval_seconds,
		       created_at, updated_at
		FROM tenants
		ORDER BY is_default DESC, id ASC
		LIMIT 1
	`)
	t, err := scanTenant(row)
	if errors.Is(err, sql.ErrNoRows) {
		return model.Tenant{}, ErrNotFound
	}
	return t, err
}

// ResolveTenantByHost matches an incoming HTTP Host header against configured tenant domains.
func (s *Store) ResolveTenantByHost(ctx context.Context, hostHeader string) (model.Tenant, error) {
	tenants, err := s.ListTenants(ctx)
	if err != nil {
		return model.Tenant{}, err
	}
	if len(tenants) == 0 {
		return model.Tenant{}, ErrNotFound
	}

	host := normalizeHost(hostHeader)
	if host != "" {
		// Exact domain match first
		for _, t := range tenants {
			for _, d := range t.DomainList() {
				if d == host {
					return t, nil
				}
			}
		}
		// Wildcard match (*.example.com)
		for _, t := range tenants {
			for _, d := range t.DomainList() {
				if strings.HasPrefix(d, "*.") {
					suffix := strings.TrimPrefix(d, "*")
					if strings.HasSuffix(host, suffix) && len(host) > len(suffix) {
						return t, nil
					}
				}
			}
		}
	}

	// Fallback to default tenant
	for _, t := range tenants {
		if t.IsDefault {
			return t, nil
		}
	}
	return tenants[0], nil
}

// CreateTenant inserts a new tenant into SQLite.
func (s *Store) CreateTenant(ctx context.Context, t *model.Tenant) error {
	t.Slug = model.Slugify(t.Slug)
	if t.Slug == "" {
		t.Slug = model.Slugify(t.Name)
	}
	if t.Slug == "" {
		return errors.New("tenant slug is required")
	}
	if t.MaxDays <= 0 {
		t.MaxDays = 10
	}
	if t.MaxRoomLists <= 0 {
		t.MaxRoomLists = 10
	}
	if t.MaxRooms <= 0 {
		t.MaxRooms = 20
	}
	if t.MaxItems <= 0 {
		t.MaxItems = 6
	}
	if t.PollIntervalSeconds < 10 {
		t.PollIntervalSeconds = 60
	}
	if t.ProviderType == "" {
		t.ProviderType = model.ProviderMock
	}
	if t.PrivacyMode == "" {
		t.PrivacyMode = model.PrivacyNormal
	}
	if t.ThemePreset == "" {
		t.ThemePreset = model.PresetStandard
	}

	themeJSON, err := t.MarshalTheme()
	if err != nil {
		return fmt.Errorf("marshal theme: %w", err)
	}

	tx, err := s.db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	defer func() { _ = tx.Rollback() }()

	if t.IsDefault {
		if _, err := tx.ExecContext(ctx, `UPDATE tenants SET is_default = 0`); err != nil {
			return err
		}
	}

	now := time.Now().UTC()
	res, err := tx.ExecContext(ctx, `
		INSERT INTO tenants (
			slug, name, domains, is_default, provider_type,
			oauth_client_id, oauth_authority, oauth_client_secret, exchange_domain,
			max_days, max_room_lists, max_rooms, max_items, room_blacklist,
			privacy_mode, theme_preset, theme_json, custom_css, poll_interval_seconds,
			created_at, updated_at
		) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
	`,
		t.Slug, t.Name, t.Domains, boolToInt(t.IsDefault), t.ProviderType,
		t.OAuthClientID, t.OAuthAuthority, t.OAuthClientSecret, t.ExchangeDomain,
		t.MaxDays, t.MaxRoomLists, t.MaxRooms, t.MaxItems, t.RoomBlacklist,
		t.PrivacyMode, t.ThemePreset, themeJSON, t.CustomCSS, t.PollIntervalSeconds,
		now, now,
	)
	if err != nil {
		if strings.Contains(strings.ToLower(err.Error()), "unique") {
			return ErrSlugExists
		}
		return fmt.Errorf("insert tenant: %w", err)
	}

	id, err := res.LastInsertId()
	if err != nil {
		return err
	}
	t.ID = id
	t.HasClientSecret = t.OAuthClientSecret != ""
	t.CreatedAt = now
	t.UpdatedAt = now

	return tx.Commit()
}

// UpdateTenant updates an existing tenant. If keepSecretWhenEmpty is true and
// t.OAuthClientSecret is empty, the stored secret is preserved.
func (s *Store) UpdateTenant(ctx context.Context, t *model.Tenant, keepSecretWhenEmpty bool) error {
	t.Slug = model.Slugify(t.Slug)
	if t.Slug == "" {
		return errors.New("tenant slug is required")
	}
	if t.MaxDays <= 0 {
		t.MaxDays = 10
	}
	if t.MaxRoomLists <= 0 {
		t.MaxRoomLists = 10
	}
	if t.MaxRooms <= 0 {
		t.MaxRooms = 20
	}
	if t.MaxItems <= 0 {
		t.MaxItems = 6
	}
	if t.PollIntervalSeconds < 10 {
		t.PollIntervalSeconds = 60
	}

	themeJSON, err := t.MarshalTheme()
	if err != nil {
		return fmt.Errorf("marshal theme: %w", err)
	}

	tx, err := s.db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	defer func() { _ = tx.Rollback() }()

	if keepSecretWhenEmpty && t.OAuthClientSecret == "" {
		var existingSecret string
		err := tx.QueryRowContext(ctx, `SELECT oauth_client_secret FROM tenants WHERE id = ?`, t.ID).Scan(&existingSecret)
		if errors.Is(err, sql.ErrNoRows) {
			return ErrNotFound
		}
		if err != nil {
			return err
		}
		t.OAuthClientSecret = existingSecret
	}

	if t.IsDefault {
		if _, err := tx.ExecContext(ctx, `UPDATE tenants SET is_default = 0 WHERE id != ?`, t.ID); err != nil {
			return err
		}
	}

	now := time.Now().UTC()
	res, err := tx.ExecContext(ctx, `
		UPDATE tenants SET
			slug = ?, name = ?, domains = ?, is_default = ?, provider_type = ?,
			oauth_client_id = ?, oauth_authority = ?, oauth_client_secret = ?, exchange_domain = ?,
			max_days = ?, max_room_lists = ?, max_rooms = ?, max_items = ?, room_blacklist = ?,
			privacy_mode = ?, theme_preset = ?, theme_json = ?, custom_css = ?, poll_interval_seconds = ?,
			updated_at = ?
		WHERE id = ?
	`,
		t.Slug, t.Name, t.Domains, boolToInt(t.IsDefault), t.ProviderType,
		t.OAuthClientID, t.OAuthAuthority, t.OAuthClientSecret, t.ExchangeDomain,
		t.MaxDays, t.MaxRoomLists, t.MaxRooms, t.MaxItems, t.RoomBlacklist,
		t.PrivacyMode, t.ThemePreset, themeJSON, t.CustomCSS, t.PollIntervalSeconds,
		now, t.ID,
	)
	if err != nil {
		if strings.Contains(strings.ToLower(err.Error()), "unique") {
			return ErrSlugExists
		}
		return fmt.Errorf("update tenant: %w", err)
	}
	n, _ := res.RowsAffected()
	if n == 0 {
		return ErrNotFound
	}
	t.HasClientSecret = t.OAuthClientSecret != ""
	t.UpdatedAt = now

	return tx.Commit()
}

// DeleteTenant removes a tenant and its assets unless it is the last tenant.
func (s *Store) DeleteTenant(ctx context.Context, id int64) error {
	tx, err := s.db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	defer func() { _ = tx.Rollback() }()

	var count int
	if err := tx.QueryRowContext(ctx, `SELECT COUNT(*) FROM tenants`).Scan(&count); err != nil {
		return err
	}
	if count <= 1 {
		return ErrCannotDeleteLast
	}

	var slug string
	var isDefault int
	err = tx.QueryRowContext(ctx, `SELECT slug, is_default FROM tenants WHERE id = ?`, id).Scan(&slug, &isDefault)
	if errors.Is(err, sql.ErrNoRows) {
		return ErrNotFound
	}
	if err != nil {
		return err
	}

	if _, err := tx.ExecContext(ctx, `DELETE FROM tenant_assets WHERE tenant_slug = ?`, slug); err != nil {
		return err
	}
	if _, err := tx.ExecContext(ctx, `DELETE FROM tenants WHERE id = ?`, id); err != nil {
		return err
	}

	if isDefault == 1 {
		if _, err := tx.ExecContext(ctx, `
			UPDATE tenants SET is_default = 1
			WHERE id = (SELECT id FROM tenants ORDER BY id ASC LIMIT 1)
		`); err != nil {
			return err
		}
	}

	return tx.Commit()
}

// GetAdminByUsername looks up an admin user by username.
func (s *Store) GetAdminByUsername(ctx context.Context, username string) (model.AdminUser, error) {
	var u model.AdminUser
	err := s.db.QueryRowContext(ctx, `
		SELECT id, username, password_hash, created_at
		FROM admin_users WHERE LOWER(username) = LOWER(?)
	`, strings.TrimSpace(username)).Scan(&u.ID, &u.Username, &u.PasswordHash, &u.CreatedAt)
	if errors.Is(err, sql.ErrNoRows) {
		return model.AdminUser{}, ErrNotFound
	}
	return u, err
}

// VerifyAdmin checks username and password against stored bcrypt hash.
func (s *Store) VerifyAdmin(ctx context.Context, username, password string) (model.AdminUser, error) {
	u, err := s.GetAdminByUsername(ctx, username)
	if err != nil {
		return model.AdminUser{}, err
	}
	if err := bcrypt.CompareHashAndPassword([]byte(u.PasswordHash), []byte(password)); err != nil {
		return model.AdminUser{}, ErrNotFound
	}
	return u, nil
}

// UpdateAdminPassword updates or creates the password hash for an admin user.
func (s *Store) UpdateAdminPassword(ctx context.Context, username, newPassword string) error {
	username = strings.TrimSpace(username)
	if username == "" {
		return errors.New("username is required")
	}
	if len(newPassword) < 6 {
		return errors.New("password must be at least 6 characters")
	}
	hash, err := bcrypt.GenerateFromPassword([]byte(newPassword), bcrypt.DefaultCost)
	if err != nil {
		return err
	}

	res, err := s.db.ExecContext(ctx, `UPDATE admin_users SET password_hash = ? WHERE LOWER(username) = LOWER(?)`, string(hash), username)
	if err != nil {
		return err
	}
	n, _ := res.RowsAffected()
	if n == 0 {
		_, err = s.db.ExecContext(ctx, `INSERT INTO admin_users (username, password_hash) VALUES (?, ?)`, username, string(hash))
		return err
	}
	return nil
}

// UpsertAsset stores or replaces an uploaded image asset for a tenant.
func (s *Store) UpsertAsset(ctx context.Context, slug, key, contentType string, data []byte) error {
	_, err := s.db.ExecContext(ctx, `
		INSERT INTO tenant_assets (tenant_slug, asset_key, content_type, data, updated_at)
		VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
		ON CONFLICT(tenant_slug, asset_key) DO UPDATE SET
			content_type = excluded.content_type,
			data = excluded.data,
			updated_at = CURRENT_TIMESTAMP
	`, slug, key, contentType, data)
	return err
}

// GetAsset retrieves a stored image asset for a tenant.
func (s *Store) GetAsset(ctx context.Context, slug, key string) (model.TenantAsset, error) {
	var a model.TenantAsset
	err := s.db.QueryRowContext(ctx, `
		SELECT tenant_slug, asset_key, content_type, data, updated_at
		FROM tenant_assets WHERE LOWER(tenant_slug) = LOWER(?) AND asset_key = ?
	`, slug, key).Scan(&a.TenantSlug, &a.AssetKey, &a.ContentType, &a.Data, &a.UpdatedAt)
	if errors.Is(err, sql.ErrNoRows) {
		return model.TenantAsset{}, ErrNotFound
	}
	return a, err
}

// ListAssetsByTenant retrieves all uploaded image assets belonging to a tenant.
func (s *Store) ListAssetsByTenant(ctx context.Context, slug string) ([]model.TenantAsset, error) {
	rows, err := s.db.QueryContext(ctx, `
		SELECT tenant_slug, asset_key, content_type, data, updated_at
		FROM tenant_assets WHERE LOWER(tenant_slug) = LOWER(?)
		ORDER BY asset_key ASC
	`, slug)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var assets []model.TenantAsset
	for rows.Next() {
		var a model.TenantAsset
		if err := rows.Scan(&a.TenantSlug, &a.AssetKey, &a.ContentType, &a.Data, &a.UpdatedAt); err != nil {
			return nil, err
		}
		assets = append(assets, a)
	}
	return assets, rows.Err()
}

type rowScanner interface {
	Scan(dest ...any) error
}

func scanTenant(r rowScanner) (model.Tenant, error) {
	var t model.Tenant
	var isDefaultInt int
	var themeJSON string
	err := r.Scan(
		&t.ID, &t.Slug, &t.Name, &t.Domains, &isDefaultInt, &t.ProviderType,
		&t.OAuthClientID, &t.OAuthAuthority, &t.OAuthClientSecret, &t.ExchangeDomain,
		&t.MaxDays, &t.MaxRoomLists, &t.MaxRooms, &t.MaxItems, &t.RoomBlacklist,
		&t.PrivacyMode, &t.ThemePreset, &themeJSON, &t.CustomCSS, &t.PollIntervalSeconds,
		&t.CreatedAt, &t.UpdatedAt,
	)
	if err != nil {
		return model.Tenant{}, err
	}
	t.IsDefault = isDefaultInt == 1
	t.HasClientSecret = t.OAuthClientSecret != ""

	t.Theme = model.ThemeForPreset(t.ThemePreset)
	if strings.TrimSpace(themeJSON) != "" && themeJSON != "{}" {
		_ = json.Unmarshal([]byte(themeJSON), &t.Theme)
	}
	return t, nil
}

func boolToInt(b bool) int {
	if b {
		return 1
	}
	return 0
}

func normalizeHost(hostHeader string) string {
	h := strings.TrimSpace(strings.ToLower(hostHeader))
	if h == "" {
		return ""
	}
	// Handle IPv6 [::1]:8080 or host:port
	if strings.HasPrefix(h, "[") {
		if idx := strings.Index(h, "]"); idx != -1 {
			return h[:idx+1]
		}
	}
	if idx := strings.Index(h, ":"); idx != -1 {
		h = h[:idx]
	}
	return h
}

