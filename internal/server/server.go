package server

import (
	"log/slog"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/go-chi/chi/v5"
	"github.com/go-chi/chi/v5/middleware"

	"github.com/probits-as/meeteasier/internal/calendar"
	"github.com/probits-as/meeteasier/internal/config"
	"github.com/probits-as/meeteasier/internal/model"
	"github.com/probits-as/meeteasier/internal/store"
)

// Server wires HTTP routes, templates, tenant resolution, and admin auth.
type Server struct {
	store         *store.Store
	calMgr        *calendar.Manager
	sessionSecret []byte
	adminAPIToken string
	staticDir     string
	defaultCfg    config.AppConfig
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
	}
}

// WithDefaultConfig attaches AppConfig so HTTP/API tenant imports can inherit OAuth env defaults.
func (s *Server) WithDefaultConfig(cfg config.AppConfig) *Server {
	s.defaultCfg = cfg
	return s
}

// Router builds the Chi router with all public, API, static, and admin routes.
func (s *Server) Router() http.Handler {
	r := chi.NewRouter()

	r.Use(middleware.RequestID)
	r.Use(middleware.RealIP)
	r.Use(slogMiddleware)
	r.Use(middleware.Recoverer)

	// Public display routes
	r.Get("/", s.handleFlightboardPage)
	r.Get("/partials/flightboard", s.handleFlightboardPartial)
	r.Get("/single-room/{alias}", s.handleSingleRoomPage)
	r.Get("/partials/single-room/{alias}", s.handleSingleRoomPartial)
	r.Get("/tenant-assets/{slug}/{key}", s.handleTenantAsset)

	// JSON API routes (backward-compatible + health check + authorized admin API)
	r.Route("/api", func(api chi.Router) {
		api.Get("/rooms", s.handleAPIRooms)
		api.Get("/roomlists", s.handleAPIRoomlists)
		api.Get("/heartbeat", s.handleAPIHeartbeat)

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

func slogMiddleware(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		start := time.Now()
		ww := middleware.NewWrapResponseWriter(w, r.ProtoMajor)
		next.ServeHTTP(ww, r)

		// Keep noisy polling logs at debug level; normal requests at info
		lvl := slog.LevelInfo
		if strings.HasPrefix(r.URL.Path, "/partials/") ||
			strings.HasPrefix(r.URL.Path, "/css/") ||
			strings.HasPrefix(r.URL.Path, "/js/") ||
			strings.HasPrefix(r.URL.Path, "/img/") ||
			r.URL.Path == "/api/heartbeat" {
			lvl = slog.LevelDebug
		}

		slog.Log(r.Context(), lvl, "http request",
			"method", r.Method,
			"path", r.URL.Path,
			"host", r.Host,
			"status", ww.Status(),
			"duration", time.Since(start).Round(time.Millisecond),
			"req_id", middleware.GetReqID(r.Context()),
		)
	})
}

