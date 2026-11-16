package server

import (
	"log/slog"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/go-chi/chi/v5"
	"github.com/go-chi/chi/v5/middleware"

	"github.com/probits-as/meeteasier/internal/calendar"
	"github.com/probits-as/meeteasier/internal/config"
	"github.com/probits-as/meeteasier/internal/i18n"
	"github.com/probits-as/meeteasier/internal/model"
	"github.com/probits-as/meeteasier/internal/store"
)

// Server wires HTTP routes, templates, tenant resolution, and admin auth.
type Server struct {
	store         *store.Store
	calMgr        *calendar.Manager
	sessionSecret []byte
	staticDir     string

	cfgMu         sync.RWMutex
	adminAPIToken string
	defaultCfg    config.AppConfig

	startedAt    time.Time
	buildVersion string
	buildCommit  string
	metricsMu    sync.Mutex
	httpReqs     map[string]uint64
	bookings     map[string]uint64
}

// New creates a new HTTP Server instance.
func New(st *store.Store, calMgr *calendar.Manager, sessionSecret []byte, adminAPIToken, staticDir string) *Server {
	if staticDir == "" {
		staticDir = "static"
	}
	return &Server{
		store:         st,
		calMgr:        calMgr,
		sessionSecret: sessionSecret,
		adminAPIToken: strings.TrimSpace(adminAPIToken),
		staticDir:     staticDir,
		startedAt:     time.Now(),
		buildVersion:  "dev",
		buildCommit:   "none",
		httpReqs:      make(map[string]uint64),
		bookings:      make(map[string]uint64),
	}
}

// WithDefaultConfig attaches AppConfig so HTTP/API tenant imports can inherit OAuth env defaults.
func (s *Server) WithDefaultConfig(cfg config.AppConfig) *Server {
	s.cfgMu.Lock()
	s.defaultCfg = cfg
	s.cfgMu.Unlock()
	return s
}

// UpdateRuntimeConfig updates hot-reloadable configuration fields thread-safely.
func (s *Server) UpdateRuntimeConfig(cfg config.AppConfig) {
	s.cfgMu.Lock()
	s.adminAPIToken = strings.TrimSpace(cfg.AdminAPIToken)
	s.defaultCfg = cfg
	s.cfgMu.Unlock()
}

func (s *Server) getAdminAPIToken() string {
	s.cfgMu.RLock()
	tok := s.adminAPIToken
	s.cfgMu.RUnlock()
	return tok
}

func (s *Server) getDefaultConfig() config.AppConfig {
	s.cfgMu.RLock()
	cfg := s.defaultCfg
	s.cfgMu.RUnlock()
	return cfg
}

// Router builds the Chi router with all public, API, static, and admin routes.
func (s *Server) Router() http.Handler {
	r := chi.NewRouter()

	r.Use(middleware.RequestID)
	r.Use(middleware.RealIP)
	r.Use(s.slogMiddleware)
	r.Use(middleware.Recoverer)

	// Health, readiness & Prometheus metrics routes
	r.Get("/healthz", s.handleHealthz)
	r.Get("/livez", s.handleHealthz)
	r.Get("/readyz", s.handleReadyz)
	r.Get("/metrics", s.handleMetrics)

	// Public display routes
	r.Get("/", s.handleFlightboardPage)
	r.Get("/partials/flightboard", s.handleFlightboardPartial)
	r.Get("/single-room/{alias}", s.handleSingleRoomPage)
	r.Get("/partials/single-room/{alias}", s.handleSingleRoomPartial)
	r.Post("/single-room/{alias}/book", s.handleSingleRoomBook)
	r.Get("/tenant-assets/{slug}/{key}", s.handleTenantAsset)

	// JSON API routes (backward-compatible + health check + authorized admin API)
	r.Route("/api", func(api chi.Router) {
		api.Get("/rooms", s.handleAPIRooms)
		api.Get("/roomlists", s.handleAPIRoomlists)
		api.Get("/heartbeat", s.handleAPIHeartbeat)
		api.Get("/healthz", s.handleHealthz)
		api.Get("/livez", s.handleHealthz)
		api.Get("/readyz", s.handleReadyz)
		api.Post("/roombooking", s.handleAPIRoomBooking)

		api.Route("/admin", func(admAPI chi.Router) {
			admAPI.Use(s.requireAdminAPI)
			admAPI.Get("/tenants", s.handleAPIAdminListTenants)
			admAPI.Post("/tenants", s.handleAPIAdminUpsertTenant)
			admAPI.Get("/tenants/{slug}/theme", s.handleAPIAdminExportTheme)
			admAPI.Put("/tenants/{slug}/theme", s.handleAPIAdminImportTheme)
			admAPI.Post("/tenants/{slug}/theme", s.handleAPIAdminImportTheme)
		})

		api.HandleFunc("/*", func(w http.ResponseWriter, r *http.Request) {
			writeJSON(w, http.StatusNotFound, map[string]string{"error": "API route not found"})
		})
	})

	// Admin authentication & management routes
	r.Get("/admin/login", s.handleAdminLoginPage)
	r.Post("/admin/login", s.handleAdminLoginSubmit)
	r.Post("/admin/logout", s.handleAdminLogout)

	r.Route("/admin", func(adm chi.Router) {
		adm.Use(s.requireAdmin)
		adm.Get("/", s.handleAdminDashboard)
		adm.Post("/language", s.handleAdminLanguageUpdate)
		adm.Post("/password", s.handleAdminPasswordUpdate)
		adm.Post("/tenants", s.handleAdminTenantCreate)
		adm.Post("/tenants/import", s.handleAdminTenantImportCreate)
		adm.Post("/tenants/{id}", s.handleAdminTenantUpdate)
		adm.Post("/tenants/{id}/preset", s.handleAdminTenantPreset)
		adm.Post("/tenants/{id}/theme-export", s.handleAdminThemeExport)
		adm.Get("/tenants/{id}/theme-export", s.handleAdminThemeExport)
		adm.Post("/tenants/{id}/theme-import", s.handleAdminThemeImport)
		adm.Post("/tenants/{id}/sync", s.handleAdminTenantSync)
		adm.Post("/tenants/{id}/delete", s.handleAdminTenantDelete)
	})

	// Static files
	s.mountStatic(r)

	return r
}

func (s *Server) mountStatic(r chi.Router) {
	if _, err := os.Stat(s.staticDir); err != nil {
		slog.Warn("Static directory not found", "dir", s.staticDir, "error", err)
		return
	}
	fs := http.FileServer(http.Dir(s.staticDir))
	for _, prefix := range []string{"/css/*", "/js/*", "/img/*", "/svgs/*"} {
		r.Get(prefix, func(w http.ResponseWriter, req *http.Request) {
			fs.ServeHTTP(w, req)
		})
	}
	r.Get("/favicon.ico", func(w http.ResponseWriter, req *http.Request) {
		http.ServeFile(w, req, filepath.Join(s.staticDir, "favicon.ico"))
	})
}

func (s *Server) resolveTenant(r *http.Request) (model.Tenant, bool, error) {
	if slug := strings.TrimSpace(r.URL.Query().Get("tenant")); slug != "" {
		t, err := s.store.GetTenantBySlug(r.Context(), slug)
		if err == nil {
			return t, true, nil
		}
		if err != store.ErrNotFound {
			return model.Tenant{}, false, err
		}
	}

	host := strings.TrimSpace(r.Header.Get("X-Forwarded-Host"))
	if host == "" {
		host = r.Host
	} else if idx := strings.Index(host, ","); idx != -1 {
		host = strings.TrimSpace(host[:idx])
	}

	t, err := s.store.ResolveTenantByHost(r.Context(), host)
	return t, false, err
}

func (s *Server) activeLanguage(r *http.Request) string {
	return s.store.GetAppLanguage(r.Context(), s.getDefaultConfig().Language)
}

func (s *Server) loginLanguage(r *http.Request) string {
	if q := strings.TrimSpace(r.URL.Query().Get("lang")); q != "" {
		return i18n.Normalize(q)
	}
	if c, err := r.Cookie("meeteasier_lang"); err == nil && strings.TrimSpace(c.Value) != "" {
		return i18n.Normalize(c.Value)
	}
	return s.activeLanguage(r)
}

func (s *Server) tenantLanguage(r *http.Request, t model.Tenant) string {
	if lang := i18n.NormalizeTenantLanguage(t.Language); lang != "" {
		return lang
	}
	return s.activeLanguage(r)
}

func (s *Server) slogMiddleware(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		start := time.Now()
		ww := middleware.NewWrapResponseWriter(w, r.ProtoMajor)
		next.ServeHTTP(ww, r)

		status := ww.Status()
		if status == 0 {
			status = http.StatusOK
		}
		s.recordHTTPMetric(r.Method, status)

		// Keep noisy polling/probe logs at debug level when healthy; normal requests at info
		lvl := slog.LevelInfo
		isNoisy := strings.HasPrefix(r.URL.Path, "/partials/") ||
			strings.HasPrefix(r.URL.Path, "/css/") ||
			strings.HasPrefix(r.URL.Path, "/js/") ||
			strings.HasPrefix(r.URL.Path, "/img/") ||
			r.URL.Path == "/healthz" ||
			r.URL.Path == "/livez" ||
			r.URL.Path == "/readyz" ||
			r.URL.Path == "/metrics" ||
			r.URL.Path == "/api/heartbeat" ||
			r.URL.Path == "/api/healthz" ||
			r.URL.Path == "/api/livez" ||
			r.URL.Path == "/api/readyz"
		if isNoisy {
			if status >= http.StatusBadRequest {
				lvl = slog.LevelWarn
			} else {
				lvl = slog.LevelDebug
			}
		}

		dur := time.Since(start)
		slog.Log(r.Context(), lvl, "http request",
			"method", r.Method,
			"path", r.URL.Path,
			"host", r.Host,
			"status", status,
			"duration", dur.Round(time.Millisecond),
			"duration_ms", dur.Milliseconds(),
			"bytes", ww.BytesWritten(),
			"req_id", middleware.GetReqID(r.Context()),
		)
	})
}

