package config

import (
	"fmt"
	"strings"

	"github.com/spf13/viper"
)

// AppConfig holds global server and bootstrap settings loaded via Viper.
type AppConfig struct {
	Port            int
	Host            string
	DBPath          string
	LogLevel        string
	AdminUsername   string
	AdminPassword   string
	AdminAPIToken   string
	SessionSecret   string
	PollIntervalSec int

	// Optional automated tenant/theme bootstrap on startup (Docker volumes / K8s ConfigMaps & Secrets)
	BootstrapThemeFile string
	BootstrapDir       string
	BootstrapThemeJSON string

	// Optional legacy/seed MS Graph environment variables for initial default tenant
	OAuthClientID     string
	OAuthAuthority    string
	OAuthClientSecret string
	Domain            string
	MaxDays           int
	MaxRoomLists      int
	MaxRooms          int
	MaxItems          int
}

// Load reads configuration from Viper (CLI flags, .env file, environment variables, defaults).
func Load() AppConfig {
	viper.SetDefault("port", 8080)
	viper.SetDefault("host", "0.0.0.0")
	viper.SetDefault("db_path", "./data/meeteasier.db")
	viper.SetDefault("log_level", "info")
	viper.SetDefault("admin_username", "admin")
	viper.SetDefault("admin_password", "admin123")
	viper.SetDefault("admin_api_token", "")
	viper.SetDefault("session_secret", "")
	viper.SetDefault("poll_interval", 60)
	viper.SetDefault("bootstrap_theme_file", "")
	viper.SetDefault("bootstrap_dir", "/app/bootstrap")
	viper.SetDefault("bootstrap_theme_json", "")

	viper.SetDefault("search_maxdays", 10)
	viper.SetDefault("search_maxroomlists", 10)
	viper.SetDefault("search_maxrooms", 20)
	viper.SetDefault("search_maxitems", 6)

	// Load .env if present (non-fatal if absent)
	viper.SetConfigFile(".env")
	viper.SetConfigType("env")
	_ = viper.ReadInConfig()

	viper.SetEnvKeyReplacer(strings.NewReplacer(".", "_", "-", "_"))
	viper.AutomaticEnv()

	cfg := AppConfig{
		Port:               viper.GetInt("port"),
		Host:               viper.GetString("host"),
		DBPath:             viper.GetString("db_path"),
		LogLevel:           viper.GetString("log_level"),
		AdminUsername:      viper.GetString("admin_username"),
		AdminPassword:      viper.GetString("admin_password"),
		AdminAPIToken:      strings.TrimSpace(viper.GetString("admin_api_token")),
		SessionSecret:      viper.GetString("session_secret"),
		PollIntervalSec:    viper.GetInt("poll_interval"),
		BootstrapThemeFile: strings.TrimSpace(viper.GetString("bootstrap_theme_file")),
		BootstrapDir:       strings.TrimSpace(viper.GetString("bootstrap_dir")),
		BootstrapThemeJSON: strings.TrimSpace(viper.GetString("bootstrap_theme_json")),
		OAuthClientID:      strings.TrimSpace(viper.GetString("oauth_client_id")),
		OAuthAuthority:     strings.TrimSpace(viper.GetString("oauth_authority")),
		OAuthClientSecret:  strings.TrimSpace(viper.GetString("oauth_client_secret")),
		Domain:             strings.TrimSpace(viper.GetString("domain")),
		MaxDays:            viper.GetInt("search_maxdays"),
		MaxRoomLists:       viper.GetInt("search_maxroomlists"),
		MaxRooms:           viper.GetInt("search_maxrooms"),
		MaxItems:           viper.GetInt("search_maxitems"),
	}

	if cfg.Port <= 0 {
		cfg.Port = 8080
	}
	if cfg.PollIntervalSec < 10 {
		cfg.PollIntervalSec = 60
	}
	if cfg.MaxDays <= 0 {
		cfg.MaxDays = 10
	}
	if cfg.MaxRoomLists <= 0 {
		cfg.MaxRoomLists = 10
	}
	if cfg.MaxRooms <= 0 {
		cfg.MaxRooms = 20
	}
	if cfg.MaxItems <= 0 {
		cfg.MaxItems = 6
	}
	return cfg
}

// ListenAddr returns the host:port string for http.Server.
func (c AppConfig) ListenAddr() string {
	return fmt.Sprintf("%s:%d", c.Host, c.Port)
}

