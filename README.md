# MeetEasier

A lightweight, multi-tenant conference room display server built in **Go** (`chi`, `slog`, `templ`, `htmx`, `viper`, `cobra`, and `SQLite`).

MeetEasier connects to **Microsoft Graph API** (or runs in an offline **Mock/Demo** mode) to display conference room availability on hallway flightboards and single-room tablets, with per-tenant custom domains, branding, colors, labels, and custom CSS managed from a built-in `/admin` interface or JSON import/export API.

## Features

- **Single Go binary + SQLite (`modernc.org/sqlite`)**: No Node.js, npm, or external database required.
- **Multi-Tenant & Custom Domains**: Route tenants by HTTP `Host` / `X-Forwarded-Host` (including `*.domain.com` wildcards) or `?tenant=<slug>` preview parameter.
- **Theme Presets & Full Customization**:
  - `standard` preset (classic MeetEasier flightboard + single-room layout)
  - `compact` preset (2-column current + next meeting cards, portrait support, clickable rows)
  - Custom colors, labels, uploaded logos/banners stored in SQLite, and per-tenant Custom CSS.
- **Portable Theme Bundles (JSON Export & Import)**: Export/import themes (including base64-encoded logos/assets) from the `/admin` UI or automated scripts via `/api/admin/*`.
- **Resilient Caching**: Background polling per tenant with stale-while-revalidate fallback if Microsoft Graph is temporarily unreachable.

## Quick Start

```sh
# Generate templ views and build
templ generate
go build -o bin/meeteasier ./cmd/meeteasier

# Run server (defaults to port 8080, SQLite at ./data/meeteasier.db)
./bin/meeteasier serve
```

Or using [Task](https://taskfile.dev):

```sh
task run
```

### Default Endpoints

- **Flightboard:** `http://localhost:8080/`
- **Single room view:** `http://localhost:8080/single-room/<room-alias>`
- **Admin UI:** `http://localhost:8080/admin` (default seeded credentials: `admin` / `admin123`, configurable via `MEETEASIER_ADMIN_USERNAME` / `MEETEASIER_ADMIN_PASSWORD`)
- **JSON API:** `/api/rooms`, `/api/roomlists`, `/api/heartbeat`

To reset or update the admin password from the CLI:

```sh
./bin/meeteasier admin reset-password -u admin -p "new-strong-password"
```

## Configuration, Precedence & Hot-Reload

MeetEasier loads settings with **Viper** using the following precedence (highest to lowest):

1. **CLI flags** (e.g. `--port`, `--db-path`, `--log-level`, `--lang`, `--config`)
2. **Prefixed environment variables (`MEETEASIER_*`)** (e.g. `MEETEASIER_PORT`, `MEETEASIER_LANGUAGE`, `MEETEASIER_ADMIN_USERNAME`, `MEETEASIER_OAUTH_CLIENT_ID`)
3. **Unprefixed environment variables** (backward-compatible fallback, e.g. `PORT`, `LANGUAGE`, `OAUTH_CLIENT_ID`)
4. **Configuration file** (`--config` / `MEETEASIER_CONFIG_FILE`, or auto-discovered `meeteasier.yaml` / `config.yaml` / `.env` in `.`, `/app/config`, or `/etc/meeteasier`)
5. **Built-in defaults** (English `en` interface language by default)

When a configuration file is loaded, MeetEasier watches it automatically (`fsnotify`) and hot-reloads runtime changes such as `log_level`, `language`, and `admin_api_token` without restarting the server (compatible with Kubernetes `ConfigMap` updates).

### Language & Localization

Supported languages: **English (`en`, default)** and **Norsk Bokmål (`nb`)**.

- **Admin UI & Login (`/admin`, `/admin/login`):** Switch interface language directly from the top-bar selector (persisted in SQLite) or via `--lang`, `MEETEASIER_LANGUAGE` (`LANGUAGE`), or `language` in the config file.
- **Per-Tenant Display Language:** Each tenant defaults to the global interface language (except legacy `compact` preset tenants, which default to `nb`), or can explicitly override `Display & booking language` (`language`) in `/admin` or Theme JSON.

## Theme Export / Import & Authorized Admin API

Themes can be exported and imported as self-contained JSON bundles (including base64-encoded image assets):

- **Admin UI:** Open `/admin`, select a tenant, and click **Export theme (JSON)** or expand **Import theme from JSON**.
- **Authorized API (`/api/admin/*`):** Authenticate using any of:
  1. `Authorization: Bearer <MEETEASIER_ADMIN_API_TOKEN>` (when configured)
  2. HTTP Basic Auth (`-u admin:password`)
  3. An active `/admin` session cookie

### Example: Create or update a tenant & theme from a JSON bundle

```sh
curl -X POST https://your-host/api/admin/tenants \
  -H "Authorization: Bearer $ADMIN_API_TOKEN" \
  -H "Content-Type: application/json" \
  -d @my-theme.json
```

### Example: Export / Import theme for an existing tenant

```sh
# Export
curl -s -u admin:yourpassword https://your-host/api/admin/tenants/default/theme > theme.json

# Import
curl -X PUT https://your-host/api/admin/tenants/default/theme \
  -u admin:yourpassword \
  -H "Content-Type: application/json" \
  -d @theme.json
```

## Microsoft 365 Setup (Graph API)

1. Register an application in Microsoft Entra ID (Azure AD) with **Application permissions**:
   - `Calendars.Read` (or `Calendars.ReadWrite` if interactive room booking on `/single-room/<alias>` is enabled)
   - `Place.Read.All`
   - `User.Read.All`
2. Grant admin consent for your organization.
3. Configure `OAuth Client ID`, `OAuth Authority` (`https://login.microsoftonline.com/<tenant-id>`), and `OAuth Client Secret` per tenant in `/admin` (or seed the initial default tenant via `.env` / environment variables from `.env.template`).

## Docker & Docker Compose

Pre-built multi-arch images (`linux/amd64` and `linux/arm64`) are published to GitHub Container Registry:
https://github.com/probits-as/MeetEasier/pkgs/container/meeteasier

### Quick Run (Mock / Demo mode)

```sh
docker run -d \
  --name meeteasier \
  -p 8080:8080 \
  -v $(pwd)/data:/app/data \
  ghcr.io/probits-as/meeteasier:latest
```

Then open `http://localhost:8080/` for the flightboard or `http://localhost:8080/admin` (`admin` / `admin123`) to configure tenants, domains, and Microsoft 365 credentials.

### Production Run (Persistent DB + Environment Variables + Optional Theme Bootstrap)

```sh
docker run -d \
  --name meeteasier \
  --restart unless-stopped \
  -p 8080:8080 \
  -v $(pwd)/data:/app/data \
  -v $(pwd)/my-theme.json:/app/bootstrap/my-theme.json:ro \
  -e MEETEASIER_ADMIN_USERNAME=admin \
  -e MEETEASIER_ADMIN_PASSWORD=change-me-strong-password \
  -e MEETEASIER_ADMIN_API_TOKEN=change-me-api-token \
  -e MEETEASIER_OAUTH_CLIENT_ID=<your-client-id> \
  -e MEETEASIER_OAUTH_AUTHORITY=https://login.microsoftonline.com/<tenant-id> \
  -e MEETEASIER_OAUTH_CLIENT_SECRET=<your-client-secret> \
  ghcr.io/probits-as/meeteasier:v2
```

### Docker Compose

```yaml
services:
  meeteasier:
    image: ghcr.io/probits-as/meeteasier:v2
    container_name: meeteasier
    restart: unless-stopped
    ports:
      - "8080:8080"
    volumes:
      - ./data:/app/data
      # Optional: auto-import theme/tenant JSON bundle on startup
      # - ./my-theme.json:/app/bootstrap/my-theme.json:ro
      # Optional: hot-reloaded configuration file
      # - ./meeteasier.yaml:/app/config/meeteasier.yaml:ro
    environment:
      MEETEASIER_ADMIN_USERNAME: admin
      MEETEASIER_ADMIN_PASSWORD: change-me-strong-password
      # MEETEASIER_ADMIN_API_TOKEN: change-me-api-token
      # MEETEASIER_OAUTH_CLIENT_ID: ""
      # MEETEASIER_OAUTH_AUTHORITY: ""
      # MEETEASIER_OAUTH_CLIENT_SECRET: ""
```

### Build Locally

```sh
docker build -t meeteasier .
docker run -p 8080:8080 -v $(pwd)/data:/app/data meeteasier
```

## Automated Provisioning (Docker & Kubernetes)

MeetEasier can automatically import or update tenants and themes at startup — ideal for GitOps deployments, ephemeral containers, or read-only infrastructure:

1. **Bootstrap Directory (`MEETEASIER_BOOTSTRAP_DIR`, default `/app/bootstrap`)**
   All `*.json` files mounted into `/app/bootstrap` are imported in alphabetical order before the server starts listening.
   - **Docker:** `-v /path/to/my-theme.json:/app/bootstrap/my-theme.json:ro`
   - **Kubernetes (`ConfigMap` / `Secret`):** Mount your JSON bundle under `/app/bootstrap`:
     ```yaml
     volumeMounts:
       - name: theme-bootstrap
         mountPath: /app/bootstrap
         readOnly: true
     volumes:
       - name: theme-bootstrap
         configMap:
           name: meeteasier-theme
     ```
2. **Explicit File (`MEETEASIER_BOOTSTRAP_THEME_FILE` or `--bootstrap-file`)**
   Specify one or more comma-separated file paths to import on startup.
3. **Environment Variable (`MEETEASIER_BOOTSTRAP_THEME_JSON`)**
   Pass raw JSON or a base64-encoded JSON theme bundle directly from a Kubernetes `Secret` / `ExternalSecret`.
4. **CLI Import / Export**
   Run standalone provisioning commands (for example in a Kubernetes `initContainer` or CI script):
   ```sh
   ./bin/meeteasier import ./my-theme.json --default
   ./bin/meeteasier export default -o ./backup-theme.json
   ```

When bootstrapping a tenant marked `"is_default": true` (or `"slug": "default"`), any unset `MEETEASIER_OAUTH_CLIENT_ID`, `MEETEASIER_OAUTH_AUTHORITY`, `MEETEASIER_OAUTH_CLIENT_SECRET`, `MEETEASIER_DOMAIN`, and `MEETEASIER_SEARCH_*` values (or their unprefixed equivalents) are automatically inherited from environment variables so secrets stay out of your theme JSON files.

## License

MeetEasier is licensed under the [GNU General Public License v3.0](LICENSE).
