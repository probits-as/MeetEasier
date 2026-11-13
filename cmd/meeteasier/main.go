package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"log/slog"
	"net/http"
	"os"
	"os/signal"
	"strings"
	"syscall"
	"time"

	"github.com/spf13/cobra"
	"github.com/spf13/viper"

	"github.com/probits-as/meeteasier/internal/calendar"
	"github.com/probits-as/meeteasier/internal/config"
	"github.com/probits-as/meeteasier/internal/model"
	"github.com/probits-as/meeteasier/internal/server"
	"github.com/probits-as/meeteasier/internal/store"
)

func main() {
	rootCmd := newRootCmd()
	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func newRootCmd() *cobra.Command {
	rootCmd := &cobra.Command{
		Use:   "meeteasier",
		Short: "MeetEasier multi-tenant conference room display server",
		RunE:  runServe,
	}

	rootCmd.PersistentFlags().String("db-path", "./data/meeteasier.db", "Path to SQLite database file")
	rootCmd.PersistentFlags().String("log-level", "info", "Logging level (debug, info, warn, error)")
	_ = viper.BindPFlag("db_path", rootCmd.PersistentFlags().Lookup("db-path"))
	_ = viper.BindPFlag("log_level", rootCmd.PersistentFlags().Lookup("log-level"))

	serveCmd := &cobra.Command{
		Use:   "serve",
		Short: "Start the MeetEasier HTTP server",
		RunE:  runServe,
	}
	serveCmd.Flags().IntP("port", "p", 8080, "Port to listen on")
	serveCmd.Flags().String("host", "0.0.0.0", "Host interface to bind")
	serveCmd.Flags().String("static-dir", "static", "Static files directory")
	serveCmd.Flags().Int("poll-interval", 60, "Default calendar polling interval in seconds")
	serveCmd.Flags().String("bootstrap-file", "", "Path (or comma-separated paths) to JSON theme bundle(s) to import at startup")
	serveCmd.Flags().String("bootstrap-dir", "/app/bootstrap", "Directory scanned for *.json theme bundles at startup")
	_ = viper.BindPFlag("port", serveCmd.Flags().Lookup("port"))
	_ = viper.BindPFlag("host", serveCmd.Flags().Lookup("host"))
	_ = viper.BindPFlag("static_dir", serveCmd.Flags().Lookup("static-dir"))
	_ = viper.BindPFlag("poll_interval", serveCmd.Flags().Lookup("poll-interval"))
	_ = viper.BindPFlag("bootstrap_theme_file", serveCmd.Flags().Lookup("bootstrap-file"))
	_ = viper.BindPFlag("bootstrap_dir", serveCmd.Flags().Lookup("bootstrap-dir"))

	var importSlug string
	var importDefault bool
	importCmd := &cobra.Command{
		Use:   "import <file.json>",
		Short: "Import or update a tenant and theme from a JSON bundle",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg := config.Load()
			setupLogger(cfg.LogLevel)

			raw, err := os.ReadFile(args[0])
			if err != nil {
				return fmt.Errorf("read bundle file: %w", err)
			}
			var bundle model.ThemeBundle
			if err := json.Unmarshal(raw, &bundle); err != nil {
				return fmt.Errorf("parse JSON bundle: %w", err)
			}
			if strings.TrimSpace(importSlug) != "" {
				bundle.Slug = strings.TrimSpace(importSlug)
			}
			if cmd.Flags().Changed("default") {
				bundle.IsDefault = &importDefault
			}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			defer cancel()

			st, err := store.Open(ctx, cfg)
			if err != nil {
				return fmt.Errorf("open store: %w", err)
			}
			defer st.Close()

			t, created, err := st.UpsertTenantFromBundle(ctx, cfg, bundle)
			if err != nil {
				return fmt.Errorf("import bundle: %w", err)
			}
			slog.Info("Imported tenant bundle",
				"slug", t.Slug,
				"name", t.Name,
				"created", created,
				"is_default", t.IsDefault,
				"provider", t.ProviderType,
			)
			return nil
		},
	}
	importCmd.Flags().StringVar(&importSlug, "slug", "", "Override target tenant slug from JSON bundle")
	importCmd.Flags().BoolVar(&importDefault, "default", false, "Set imported tenant as default tenant")

	var exportOut string
	exportCmd := &cobra.Command{
		Use:   "export <slug>",
		Short: "Export a tenant theme and assets as a portable JSON bundle",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg := config.Load()
			setupLogger(cfg.LogLevel)

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			defer cancel()

			st, err := store.Open(ctx, cfg)
			if err != nil {
				return fmt.Errorf("open store: %w", err)
			}
			defer st.Close()

			t, err := st.GetTenantBySlug(ctx, args[0])
			if err != nil {
				return fmt.Errorf("get tenant %q: %w", args[0], err)
			}
			bundle, err := st.BuildThemeBundle(ctx, t)
			if err != nil {
				return fmt.Errorf("build theme bundle: %w", err)
			}

			out := os.Stdout
			if strings.TrimSpace(exportOut) != "" {
				f, err := os.Create(exportOut)
				if err != nil {
					return fmt.Errorf("create output file: %w", err)
				}
				defer f.Close()
				out = f
			}

			enc := json.NewEncoder(out)
			enc.SetIndent("", "  ")
			return enc.Encode(bundle)
		},
	}
	exportCmd.Flags().StringVarP(&exportOut, "output", "o", "", "Write JSON bundle to file instead of stdout")

	adminCmd := &cobra.Command{
		Use:   "admin",
		Short: "Manage MeetEasier admin accounts",
	}

	var resetUser, resetPass string
	resetPwCmd := &cobra.Command{
		Use:   "reset-password",
		Short: "Reset or create an admin user password",
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg := config.Load()
			setupLogger(cfg.LogLevel)

			ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
			defer cancel()

			st, err := store.Open(ctx, cfg)
			if err != nil {
				return err
			}
			defer st.Close()

			if resetUser == "" {
				resetUser = cfg.AdminUsername
			}
			if err := st.UpdateAdminPassword(ctx, resetUser, resetPass); err != nil {
				return err
			}
			slog.Info("Admin password updated", "username", resetUser)
			return nil
		},
	}
	resetPwCmd.Flags().StringVarP(&resetUser, "username", "u", "admin", "Admin username")
	resetPwCmd.Flags().StringVarP(&resetPass, "password", "p", "", "New password (min 6 chars)")
	_ = resetPwCmd.MarkFlagRequired("password")

	adminCmd.AddCommand(resetPwCmd)
	rootCmd.AddCommand(serveCmd, importCmd, exportCmd, adminCmd)
	return rootCmd
}

func runServe(cmd *cobra.Command, args []string) error {
	cfg := config.Load()
	setupLogger(cfg.LogLevel)

	ctx, stop := signal.NotifyContext(context.Background(), syscall.SIGINT, syscall.SIGTERM)
	defer stop()

	st, err := store.Open(ctx, cfg)
	if err != nil {
		return fmt.Errorf("open store: %w", err)
	}
	defer st.Close()

	sessionSecret, err := st.GetOrCreateSessionSecret(ctx, cfg.SessionSecret)
	if err != nil {
		return fmt.Errorf("session secret: %w", err)
	}

	calMgr := calendar.NewManager(st, calendar.NewMSGraphProvider(), calendar.NewMockProvider())
	go calMgr.Start(ctx)

	staticDir := viper.GetString("static_dir")
	if staticDir == "" {
		staticDir = "static"
	}

	srvHandler := server.New(st, calMgr, sessionSecret, cfg.AdminAPIToken, staticDir).WithDefaultConfig(cfg).Router()
	httpSrv := &http.Server{
		Addr:              cfg.ListenAddr(),
		Handler:           srvHandler,
		ReadHeaderTimeout: 10 * time.Second,
	}

	errCh := make(chan error, 1)
	go func() {
		slog.Info("Starting MeetEasier server",
			"addr", cfg.ListenAddr(),
			"db_path", cfg.DBPath,
			"static_dir", staticDir,
		)
		if err := httpSrv.ListenAndServe(); err != nil && !errors.Is(err, http.ErrServerClosed) {
			errCh <- err
		}
		close(errCh)
	}()

	select {
	case err := <-errCh:
		return err
	case <-ctx.Done():
		slog.Info("Shutting down MeetEasier server...")
	}

	shutdownCtx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := httpSrv.Shutdown(shutdownCtx); err != nil {
		return fmt.Errorf("graceful shutdown: %w", err)
	}
	slog.Info("Server stopped cleanly")
	return nil
}

func setupLogger(levelStr string) {
	var level slog.Level
	switch strings.ToLower(strings.TrimSpace(levelStr)) {
	case "debug":
		level = slog.LevelDebug
	case "warn", "warning":
		level = slog.LevelWarn
	case "error":
		level = slog.LevelError
	default:
		level = slog.LevelInfo
	}
	handler := slog.NewTextHandler(os.Stdout, &slog.HandlerOptions{
		Level: level,
	})
	slog.SetDefault(slog.New(handler))
}

