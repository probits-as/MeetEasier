package config

import (
	"bufio"
	"bytes"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/fsnotify/fsnotify"
	"github.com/spf13/viper"
)

const EnvPrefix = "MEETEASIER"

var knownKeys = []string{
	"config_file",
	"port",
	"host",
	"db_path",
	"log_level",
	"static_dir",
	"admin_username",
	"admin_password",
	"admin_api_token",
	"session_secret",
	"poll_interval",
	"bootstrap_theme_file",
	"bootstrap_dir",
	"bootstrap_theme_json",
	"oauth_client_id",
	"oauth_authority",
	"oauth_client_secret",
	"domain",
	"search_maxdays",
	"search_maxroomlists",
	"search_maxrooms",
	"search_maxitems",
}

// AppConfig holds global server and bootstrap settings loaded via Viper.
// Precedence (highest to lowest): CLI flags > MEETEASIER_* env > unprefixed env > config file > defaults.
type AppConfig struct {
	ConfigFile      string
	Port            int
	Host            string
	DBPath          string
	LogLevel        string
	StaticDir       string
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

// Load reads configuration from the global Viper instance, logging any non-fatal fallback.
func Load() AppConfig {
	cfg, _ := LoadWithError()
	return cfg
}

// LoadWithError reads configuration from the global Viper instance and returns any explicit config error.
func LoadWithError() (AppConfig, error) {
	return LoadFromViper(viper.GetViper())
}

// LoadFromViper configures defaults, environment bindings, and config file discovery on the given Viper instance.
func LoadFromViper(v *viper.Viper) (AppConfig, error) {
	setupDefaultsAndEnv(v)

	if cfgPath := strings.TrimSpace(v.GetString("config_file")); cfgPath != "" {
		if abs, err := filepath.Abs(cfgPath); err == nil {
			cfgPath = abs
		}
		v.SetConfigFile(cfgPath)
		if isDotEnvPath(cfgPath) {
			v.SetConfigType("env")
			if err := readAndNormalizeDotEnv(v, cfgPath); err != nil {
				return AppConfig{}, fmt.Errorf("read config file %q: %w", cfgPath, err)
			}
		} else {
			if err := v.ReadInConfig(); err != nil {
				return AppConfig{}, fmt.Errorf("read config file %q: %w", cfgPath, err)
			}
		}
	} else {
		loaded, err := autoDiscoverConfig(v)
		if err != nil {
			return AppConfig{}, err
		}
		if !loaded {
			if absDotEnv, err := filepath.Abs(".env"); err == nil {
				if info, statErr := os.Stat(absDotEnv); statErr == nil && !info.IsDir() {
					v.SetConfigFile(absDotEnv)
					v.SetConfigType("env")
					_ = readAndNormalizeDotEnv(v, absDotEnv)
				}
			}
		}
	}

	return extractConfig(v), nil
}

// Watch starts watching the active Viper configuration file for changes (hot-reload)
// and invokes onChange with the freshly resolved AppConfig. Returns false if no config file is active.
func Watch(onChange func(AppConfig)) bool {
	return WatchViper(viper.GetViper(), onChange)
}

// WatchViper starts watching the configuration file loaded on v and triggers onChange on updates.
func WatchViper(v *viper.Viper, onChange func(AppConfig)) bool {
	cfgFile := v.ConfigFileUsed()
	if cfgFile == "" {
		return false
	}

	v.OnConfigChange(func(e fsnotify.Event) {
		if isDotEnvPath(v.ConfigFileUsed()) {
			_ = readAndNormalizeDotEnv(v, v.ConfigFileUsed())
		}
		if onChange != nil {
			onChange(extractConfig(v))
		}
	})
	v.WatchConfig()
	return true
}

func setupDefaultsAndEnv(v *viper.Viper) {
	v.SetDefault("config_file", "")
	v.SetDefault("port", 8080)
	v.SetDefault("host", "0.0.0.0")
	v.SetDefault("db_path", "./data/meeteasier.db")
	v.SetDefault("log_level", "info")
	v.SetDefault("static_dir", "static")
	v.SetDefault("admin_username", "admin")
	v.SetDefault("admin_password", "admin123")
	v.SetDefault("admin_api_token", "")
	v.SetDefault("session_secret", "")
	v.SetDefault("poll_interval", 60)
	v.SetDefault("bootstrap_theme_file", "")
	v.SetDefault("bootstrap_dir", "/app/bootstrap")
	v.SetDefault("bootstrap_theme_json", "")

	v.SetDefault("search_maxdays", 10)
	v.SetDefault("search_maxroomlists", 10)
	v.SetDefault("search_maxrooms", 20)
	v.SetDefault("search_maxitems", 6)

	v.SetEnvPrefix(EnvPrefix)
	v.SetEnvKeyReplacer(strings.NewReplacer(".", "_", "-", "_"))
	v.AutomaticEnv()

	for _, key := range knownKeys {
		upper := strings.ToUpper(key)
		_ = v.BindEnv(key, EnvPrefix+"_"+upper, upper)
	}
}

func autoDiscoverConfig(v *viper.Viper) (bool, error) {
	v.AddConfigPath(".")
	v.AddConfigPath("/app/config")
	v.AddConfigPath("/etc/meeteasier")

	for _, name := range []string{"meeteasier", "config"} {
		v.SetConfigName(name)
		err := v.ReadInConfig()
		if err == nil {
			return true, nil
		}
		var notFound viper.ConfigFileNotFoundError
		if !errors.As(err, &notFound) {
			return false, fmt.Errorf("read config %q: %w", name, err)
		}
	}
	return false, nil
}

func extractConfig(v *viper.Viper) AppConfig {
	cfg := AppConfig{
		ConfigFile:         v.ConfigFileUsed(),
		Port:               v.GetInt("port"),
		Host:               v.GetString("host"),
		DBPath:             v.GetString("db_path"),
		LogLevel:           v.GetString("log_level"),
		StaticDir:          strings.TrimSpace(v.GetString("static_dir")),
		AdminUsername:      v.GetString("admin_username"),
		AdminPassword:      v.GetString("admin_password"),
		AdminAPIToken:      strings.TrimSpace(v.GetString("admin_api_token")),
		SessionSecret:      v.GetString("session_secret"),
		PollIntervalSec:    v.GetInt("poll_interval"),
		BootstrapThemeFile: strings.TrimSpace(v.GetString("bootstrap_theme_file")),
		BootstrapDir:       strings.TrimSpace(v.GetString("bootstrap_dir")),
		BootstrapThemeJSON: strings.TrimSpace(v.GetString("bootstrap_theme_json")),
		OAuthClientID:      strings.TrimSpace(v.GetString("oauth_client_id")),
		OAuthAuthority:     strings.TrimSpace(v.GetString("oauth_authority")),
		OAuthClientSecret:  strings.TrimSpace(v.GetString("oauth_client_secret")),
		Domain:             strings.TrimSpace(v.GetString("domain")),
		MaxDays:            v.GetInt("search_maxdays"),
		MaxRoomLists:       v.GetInt("search_maxroomlists"),
		MaxRooms:           v.GetInt("search_maxrooms"),
		MaxItems:           v.GetInt("search_maxitems"),
	}

	if cfg.Port <= 0 {
		cfg.Port = 8080
	}
	if cfg.StaticDir == "" {
		cfg.StaticDir = "static"
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

func isDotEnvPath(path string) bool {
	base := strings.ToLower(filepath.Base(path))
	return base == ".env" || strings.HasSuffix(base, ".env")
}

// readAndNormalizeDotEnv reads a .env file and normalizes MEETEASIER_* keys into Viper's config layer
// so that MEETEASIER_FOO=bar inside .env behaves identically to foo: bar in YAML (strictly below env & flags).
func readAndNormalizeDotEnv(v *viper.Viper, path string) error {
	raw, err := os.ReadFile(path)
	if err != nil {
		return err
	}

	var plainLines []string
	var prefixedLines []string
	prefixDot := EnvPrefix + "_"

	scanner := bufio.NewScanner(bytes.NewReader(raw))
	for scanner.Scan() {
		line := scanner.Text()
		trimmed := strings.TrimSpace(line)
		if trimmed == "" || strings.HasPrefix(trimmed, "#") {
			continue
		}
		if strings.HasPrefix(trimmed, "export ") {
			trimmed = strings.TrimSpace(strings.TrimPrefix(trimmed, "export "))
		}
		eq := strings.IndexByte(trimmed, '=')
		if eq <= 0 {
			continue
		}
		key := strings.TrimSpace(trimmed[:eq])
		val := trimmed[eq+1:]
		if strings.HasPrefix(strings.ToUpper(key), prefixDot) {
			strippedKey := key[len(prefixDot):]
			prefixedLines = append(prefixedLines, strippedKey+"="+val)
		} else {
			plainLines = append(plainLines, key+"="+val)
		}
	}
	if err := scanner.Err(); err != nil {
		return err
	}

	var buf bytes.Buffer
	for _, l := range plainLines {
		buf.WriteString(l)
		buf.WriteByte('\n')
	}
	for _, l := range prefixedLines {
		buf.WriteString(l)
		buf.WriteByte('\n')
	}
	return v.ReadConfig(&buf)
}

// ListenAddr returns the host:port string for http.Server.
func (c AppConfig) ListenAddr() string {
	return fmt.Sprintf("%s:%d", c.Host, c.Port)
}
