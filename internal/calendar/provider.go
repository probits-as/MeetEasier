package calendar

import (
	"context"
	"log/slog"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/probits-as/meeteasier/internal/model"
)

// Provider fetches rooms and roomlists for a given tenant.
type Provider interface {
	FetchRooms(ctx context.Context, tenant model.Tenant) ([]model.Room, error)
}

// TenantStore defines the store methods needed by the calendar manager.
type TenantStore interface {
	ListTenants(ctx context.Context) ([]model.Tenant, error)
}

type tenantCacheEntry struct {
	tenant      model.Tenant
	rooms       []model.Room
	roomlists   []string
	lastSync    time.Time
	nextSync    time.Time
	lastError   string
	isStale     bool
	syncing     bool
	initialized bool
}

// Manager maintains background calendar polling and stale-while-revalidate caches per tenant.
type Manager struct {
	store   TenantStore
	msgraph Provider
	mock    Provider

	mu      sync.RWMutex
	entries map[int64]*tenantCacheEntry
	wakeCh  chan struct{}
}

// NewManager creates a new calendar Manager.
func NewManager(store TenantStore, msgraph, mock Provider) *Manager {
	if msgraph == nil {
		msgraph = NewMSGraphProvider()
	}
	if mock == nil {
		mock = NewMockProvider()
	}
	return &Manager{
		store:   store,
		msgraph: msgraph,
		mock:    mock,
		entries: make(map[int64]*tenantCacheEntry),
		wakeCh:  make(chan struct{}, 1),
	}
}

// Start loads initial tenants and runs the background polling loop until ctx is cancelled.
func (m *Manager) Start(ctx context.Context) {
	if err := m.ReloadTenants(ctx); err != nil {
		slog.Error("Initial tenant load failed", "error", err)
	}

	// Perform initial non-blocking sync for all loaded tenants
	m.syncDue(ctx, true)

	ticker := time.NewTicker(5 * time.Second)
	defer ticker.Stop()

	for {
		select {
		case <-ctx.Done():
			return
		case <-ticker.C:
			m.syncDue(ctx, false)
		case <-m.wakeCh:
			m.syncDue(ctx, false)
		}
	}
}

// ReloadTenants refreshes the list of active tenants from SQLite without discarding cached rooms.
func (m *Manager) ReloadTenants(ctx context.Context) error {
	tenants, err := m.store.ListTenants(ctx)
	if err != nil {
		return err
	}

	m.mu.Lock()
	defer m.mu.Unlock()

	seen := make(map[int64]struct{}, len(tenants))
	for _, t := range tenants {
		seen[t.ID] = struct{}{}
		if existing, ok := m.entries[t.ID]; ok {
			providerChanged := existing.tenant.ProviderType != t.ProviderType ||
				existing.tenant.OAuthClientID != t.OAuthClientID ||
				existing.tenant.OAuthAuthority != t.OAuthAuthority ||
				existing.tenant.OAuthClientSecret != t.OAuthClientSecret ||
				existing.tenant.PrivacyMode != t.PrivacyMode ||
				existing.tenant.RoomBlacklist != t.RoomBlacklist
			existing.tenant = t
			if providerChanged {
				existing.nextSync = time.Time{}
			}
		} else {
			m.entries[t.ID] = &tenantCacheEntry{
				tenant: t,
			}
		}
	}

	for id := range m.entries {
		if _, ok := seen[id]; !ok {
			delete(m.entries, id)
		}
	}

	select {
	case m.wakeCh <- struct{}{}:
	default:
	}

	return nil
}

// SyncTenantNow forces an immediate synchronous refresh for a specific tenant.
func (m *Manager) SyncTenantNow(ctx context.Context, tenant model.Tenant) model.TenantStatus {
	m.mu.Lock()
	entry, ok := m.entries[tenant.ID]
	if !ok {
		entry = &tenantCacheEntry{tenant: tenant}
		m.entries[tenant.ID] = entry
	} else {
		entry.tenant = tenant
	}
	entry.syncing = true
	m.mu.Unlock()

	m.syncSingle(ctx, tenant)
	return m.GetStatus(tenant)
}

// GetStatus returns the cached rooms, roomlists, and sync metadata for a tenant.
// If the tenant has not yet synced (e.g. first request in tests or before initial poll finishes),
// for mock provider it populates synchronously on demand.
func (m *Manager) GetStatus(tenant model.Tenant) model.TenantStatus {
	m.mu.RLock()
	entry, ok := m.entries[tenant.ID]
	if !ok || (!entry.initialized && tenant.ProviderType == model.ProviderMock) {
		m.mu.RUnlock()
		// Populate mock synchronously so UI/tests never wait on background tick
		if tenant.ProviderType == model.ProviderMock {
			m.syncSingle(context.Background(), tenant)
			return m.GetStatus(tenant)
		}
		return model.TenantStatus{
			TenantID:   tenant.ID,
			TenantSlug: tenant.Slug,
		}
	}

	rooms := make([]model.Room, len(entry.rooms))
	copy(rooms, entry.rooms)
	roomlists := make([]string, len(entry.roomlists))
	copy(roomlists, entry.roomlists)

	status := model.TenantStatus{
		TenantID:    tenant.ID,
		TenantSlug:  tenant.Slug,
		Rooms:       rooms,
		Roomlists:   roomlists,
		LastSync:    entry.lastSync,
		LastError:   entry.lastError,
		IsStale:     entry.isStale,
		Syncing:     entry.syncing,
		Initialized: entry.initialized,
	}
	m.mu.RUnlock()
	return status
}

// GetRoomByAlias returns a single room matching the alias from cache.
func (m *Manager) GetRoomByAlias(tenant model.Tenant, alias string) (model.Room, bool) {
	status := m.GetStatus(tenant)
	target := strings.ToLower(strings.TrimSpace(alias))
	for _, r := range status.Rooms {
		if strings.ToLower(r.RoomAlias) == target {
			return r, true
		}
	}
	return model.Room{}, false
}

func (m *Manager) syncDue(ctx context.Context, forceAll bool) {
	now := time.Now()
	m.mu.RLock()
	var due []model.Tenant
	for _, entry := range m.entries {
		if entry.syncing {
			continue
		}
		if forceAll || entry.nextSync.IsZero() || !now.Before(entry.nextSync) {
			due = append(due, entry.tenant)
		}
	}
	m.mu.RUnlock()

	for _, t := range due {
		if ctx.Err() != nil {
			return
		}
		go m.syncSingle(ctx, t)
	}
}

func (m *Manager) syncSingle(ctx context.Context, tenant model.Tenant) {
	m.mu.Lock()
	entry, ok := m.entries[tenant.ID]
	if !ok {
		entry = &tenantCacheEntry{tenant: tenant}
		m.entries[tenant.ID] = entry
	}
	entry.syncing = true
	m.mu.Unlock()

	provider := m.providerFor(tenant)
	fetchCtx, cancel := context.WithTimeout(ctx, 25*time.Second)
	defer cancel()

	rooms, err := provider.FetchRooms(fetchCtx, tenant)

	interval := time.Duration(tenant.PollIntervalSeconds) * time.Second
	if interval < 10*time.Second {
		interval = 60 * time.Second
	}

	m.mu.Lock()
	defer m.mu.Unlock()

	entry, ok = m.entries[tenant.ID]
	if !ok {
		return
	}
	entry.syncing = false
	entry.nextSync = time.Now().Add(interval)

	if err != nil {
		entry.lastError = err.Error()
		entry.isStale = len(entry.rooms) > 0
		entry.initialized = true
		slog.Warn("Calendar sync failed; preserving cached data if available",
			"tenant", tenant.Slug,
			"provider", tenant.ProviderType,
			"cached_rooms", len(entry.rooms),
			"error", err,
		)
		return
	}

	roomlists := extractRoomlists(rooms)
	entry.rooms = rooms
	entry.roomlists = roomlists
	entry.lastSync = time.Now()
	entry.lastError = ""
	entry.isStale = false
	entry.initialized = true

	slog.Debug("Calendar sync completed",
		"tenant", tenant.Slug,
		"provider", tenant.ProviderType,
		"rooms", len(rooms),
		"roomlists", len(roomlists),
	)
}

func (m *Manager) providerFor(t model.Tenant) Provider {
	if strings.EqualFold(t.ProviderType, model.ProviderMSGraph) {
		return m.msgraph
	}
	return m.mock
}

func extractRoomlists(rooms []model.Room) []string {
	seen := make(map[string]struct{})
	var lists []string
	for _, r := range rooms {
		if r.Roomlist == "" {
			continue
		}
		if _, ok := seen[r.Roomlist]; !ok {
			seen[r.Roomlist] = struct{}{}
			lists = append(lists, r.Roomlist)
		}
	}
	sort.Slice(lists, func(i, j int) bool {
		return strings.ToLower(lists[i]) < strings.ToLower(lists[j])
	})
	return lists
}

func applyPrivacy(privacyMode, rawSubject, sensitivity, organizer string) (string, bool) {
	isPrivate := !strings.EqualFold(strings.TrimSpace(sensitivity), "normal") && strings.TrimSpace(sensitivity) != ""

	switch privacyMode {
	case model.PrivacyOrganizerAsSubject:
		if strings.TrimSpace(organizer) != "" {
			return organizer, true
		}
		return "Private", true
	case model.PrivacyHideSubject:
		return "Private", true
	default: // PrivacyNormal
		if isPrivate {
			return "Private", true
		}
		if strings.TrimSpace(rawSubject) == "" {
			return "(No title)", false
		}
		return rawSubject, false
	}
}

