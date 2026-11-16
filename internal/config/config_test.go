package config

import (
	"os"
	"path/filepath"
	"testing"
	"time"

	"github.com/spf13/pflag"
	"github.com/spf13/viper"
)

func TestPrecedenceAndEnvPrefixAndHotReload(t *testing.T) {
	dir := t.TempDir()
	cfgPath := filepath.Join(dir, "meeteasier.yaml")
	initialYAML := `
port: 9000
log_level: warn
admin_username: yaml-admin
poll_interval: 120
`
	if err := os.WriteFile(cfgPath, []byte(initialYAML), 0o600); err != nil {
		t.Fatalf("write yaml: %v", err)
	}

	// Unprefixed fallback env and MEETEASIER_* prefixed env (prefixed must win over unprefixed,
	// both must win over YAML config, and CLI flags must win over env).
	t.Setenv("MEETEASIER_CONFIG_FILE", cfgPath)
	t.Setenv("ADMIN_USERNAME", "legacy-admin")
	t.Setenv("MEETEASIER_ADMIN_USERNAME", "prefixed-admin")
	t.Setenv("OAUTH_CLIENT_ID", "legacy-client-id")
	t.Setenv("PORT", "9100")

	flags := pflag.NewFlagSet("test", pflag.ContinueOnError)
	flags.Int("port", 8080, "")
	if err := flags.Parse([]string{"--port", "9200"}); err != nil {
		t.Fatalf("parse flags: %v", err)
	}

	v := viper.New()
	if err := v.BindPFlag("port", flags.Lookup("port")); err != nil {
		t.Fatalf("BindPFlag: %v", err)
	}

	cfg, err := LoadFromViper(v)
	if err != nil {
		t.Fatalf("LoadFromViper failed: %v", err)
	}

	if cfg.Port != 9200 {
		t.Fatalf("expected CLI flag port=9200 to override env and config file, got %d", cfg.Port)
	}
	if cfg.AdminUsername != "prefixed-admin" {
		t.Fatalf("expected MEETEASIER_ADMIN_USERNAME to override ADMIN_USERNAME and yaml, got %q", cfg.AdminUsername)
	}
	if cfg.OAuthClientID != "legacy-client-id" {
		t.Fatalf("expected unprefixed OAUTH_CLIENT_ID fallback to work, got %q", cfg.OAuthClientID)
	}
	if cfg.LogLevel != "warn" || cfg.PollIntervalSec != 120 {
		t.Fatalf("expected YAML values log_level=warn, poll_interval=120, got %q / %d", cfg.LogLevel, cfg.PollIntervalSec)
	}

	// Test hot-reload via WatchViper
	reloaded := make(chan AppConfig, 2)
	if ok := WatchViper(v, func(next AppConfig) {
		reloaded <- next
	}); !ok {
		t.Fatalf("expected WatchViper to start watching %s", cfgPath)
	}

	updatedYAML := `
port: 9000
log_level: debug
admin_username: yaml-admin
poll_interval: 120
`
	if err := os.WriteFile(cfgPath, []byte(updatedYAML), 0o600); err != nil {
		t.Fatalf("update yaml: %v", err)
	}

	select {
	case next := <-reloaded:
		if next.LogLevel != "debug" {
			t.Fatalf("expected hot-reloaded log_level=debug, got %q", next.LogLevel)
		}
	case <-time.After(3 * time.Second):
		t.Fatalf("timed out waiting for config hot-reload")
	}
}

func TestDotEnvPrefixedKeysDoNotOverrideRealEnv(t *testing.T) {
	dir := t.TempDir()
	envPath := filepath.Join(dir, "test.env")
	dotEnv := `
MEETEASIER_PORT=9090
MEETEASIER_ADMIN_USERNAME=dotenv-admin
MEETEASIER_LOG_LEVEL=debug
`
	if err := os.WriteFile(envPath, []byte(dotEnv), 0o600); err != nil {
		t.Fatalf("write .env: %v", err)
	}

	t.Setenv("MEETEASIER_CONFIG_FILE", envPath)
	t.Setenv("MEETEASIER_ADMIN_USERNAME", "real-env-admin")

	v := viper.New()
	cfg, err := LoadFromViper(v)
	if err != nil {
		t.Fatalf("LoadFromViper failed: %v", err)
	}
	if cfg.Port != 9090 || cfg.LogLevel != "debug" {
		t.Fatalf("expected MEETEASIER_* keys from .env to load, got port=%d log_level=%q", cfg.Port, cfg.LogLevel)
	}
	if cfg.AdminUsername != "real-env-admin" {
		t.Fatalf("expected process env to override .env value, got %q", cfg.AdminUsername)
	}
}


func TestLanguageConfigPrecedence(t *testing.T) {
	// 1. Built-in default when unset
	t.Setenv("MEETEASIER_LANGUAGE", "")
	t.Setenv("LANGUAGE", "")
	vDefault := viper.New()
	cfgDefault, err := LoadFromViper(vDefault)
	if err != nil {
		t.Fatalf("LoadFromViper default failed: %v", err)
	}
	if cfgDefault.Language != "en" || cfgDefault.LanguageExplicit {
		t.Fatalf("expected default Language=en LanguageExplicit=false, got %q explicit=%v", cfgDefault.Language, cfgDefault.LanguageExplicit)
	}

	// 2. Unchanged bound --lang flag does not mark LanguageExplicit
	fsUnchanged := pflag.NewFlagSet("test-unchanged", pflag.ContinueOnError)
	fsUnchanged.String("lang", "en", "")
	vUnchanged := viper.New()
	if err := vUnchanged.BindPFlag("language", fsUnchanged.Lookup("lang")); err != nil {
		t.Fatalf("BindPFlag unchanged: %v", err)
	}
	cfgUnchanged, err := LoadFromViper(vUnchanged)
	if err != nil {
		t.Fatalf("LoadFromViper unchanged failed: %v", err)
	}
	if cfgUnchanged.Language != "en" || cfgUnchanged.LanguageExplicit {
		t.Fatalf("expected unchanged flag Language=en LanguageExplicit=false, got %q explicit=%v", cfgUnchanged.Language, cfgUnchanged.LanguageExplicit)
	}

	// 3. Config file language is marked explicit
	tmpDir := t.TempDir()
	yamlPath := filepath.Join(tmpDir, "config.yaml")
	if err := os.WriteFile(yamlPath, []byte("language: nb\n"), 0o600); err != nil {
		t.Fatalf("write config.yaml: %v", err)
	}
	vFile := viper.New()
	vFile.Set("config_file", yamlPath)
	cfgFile, err := LoadFromViper(vFile)
	if err != nil {
		t.Fatalf("LoadFromViper file failed: %v", err)
	}
	if cfgFile.Language != "nb" || !cfgFile.LanguageExplicit {
		t.Fatalf("expected config file Language=nb LanguageExplicit=true, got %q explicit=%v", cfgFile.Language, cfgFile.LanguageExplicit)
	}

	// 4. Prefixed env MEETEASIER_LANGUAGE overrides config file
	t.Setenv("MEETEASIER_LANGUAGE", "en")
	vEnv := viper.New()
	vEnv.Set("config_file", yamlPath)
	cfgEnv, err := LoadFromViper(vEnv)
	if err != nil {
		t.Fatalf("LoadFromViper env failed: %v", err)
	}
	if cfgEnv.Language != "en" || !cfgEnv.LanguageExplicit {
		t.Fatalf("expected env Language=en LanguageExplicit=true, got %q explicit=%v", cfgEnv.Language, cfgEnv.LanguageExplicit)
	}

	// 5. Explicit --lang flag overrides MEETEASIER_LANGUAGE
	t.Setenv("MEETEASIER_LANGUAGE", "en")
	fs := pflag.NewFlagSet("test", pflag.ContinueOnError)
	fs.String("lang", "en", "")
	if err := fs.Parse([]string{"--lang=nb"}); err != nil {
		t.Fatalf("parse flags: %v", err)
	}
	vFlag := viper.New()
	if err := vFlag.BindPFlag("language", fs.Lookup("lang")); err != nil {
		t.Fatalf("BindPFlag: %v", err)
	}
	cfgFlag, err := LoadFromViper(vFlag)
	if err != nil {
		t.Fatalf("LoadFromViper flag failed: %v", err)
	}
	if cfgFlag.Language != "nb" || !cfgFlag.LanguageExplicit {
		t.Fatalf("expected flag Language=nb LanguageExplicit=true, got %q explicit=%v", cfgFlag.Language, cfgFlag.LanguageExplicit)
	}
}
