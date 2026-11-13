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
- **Admin UI:** `http://localhost:8080/admin` (default seeded credentials: `admin` / `admin123`, configurable via `ADMIN_USERNAME` / `ADMIN_PASSWORD`)
- **JSON API:** `/api/rooms`, `/api/roomlists`, `/api/heartbeat`

To reset or update the admin password from the CLI:

```sh
./bin/meeteasier admin reset-password -u admin -p "new-strong-password"
```

## Theme Export / Import & Authorized Admin API

Themes can be exported and imported as self-contained JSON bundles (including base64-encoded image assets):

- **Admin UI:** Open `/admin`, select a tenant, and click **Eksporter tema (JSON)** or expand **Importer tema fra JSON**.
- **Authorized API (`/api/admin/*`):** Authenticate using any of:
  1. `Authorization: Bearer <ADMIN_API_TOKEN>` (when `ADMIN_API_TOKEN` is configured)
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
   - `Calendars.Read`
   - `Place.Read.All`
   - `User.Read.All`
2. Grant admin consent for your organization.
3. Configure `OAuth Client ID`, `OAuth Authority` (`https://login.microsoftonline.com/<tenant-id>`), and `OAuth Client Secret` per tenant in `/admin` (or seed the initial default tenant via `.env` / environment variables from `.env.template`).

## Docker

```sh
docker build -t meeteasier .
docker run -p 8080:8080 \
  -v $(pwd)/data:/app/data \
  -v $(pwd)/my-theme.json:/app/bootstrap/my-theme.json:ro \
  meeteasier
```

## Automated Provisioning (Docker & Kubernetes)

MeetEasier can automatically import or update tenants and themes at startup — ideal for GitOps deployments, ephemeral containers, or read-only infrastructure:

1. **Bootstrap Directory (`BOOTSTRAP_DIR`, default `/app/bootstrap`)**
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
2. **Explicit File (`BOOTSTRAP_THEME_FILE` or `--bootstrap-file`)**
   Specify one or more comma-separated file paths to import on startup.
3. **Environment Variable (`BOOTSTRAP_THEME_JSON`)**
   Pass raw JSON or a base64-encoded JSON theme bundle directly from a Kubernetes `Secret` / `ExternalSecret`.
4. **CLI Import / Export**
   Run standalone provisioning commands (for example in a Kubernetes `initContainer` or CI script):
   ```sh
   ./bin/meeteasier import ./my-theme.json --default
   ./bin/meeteasier export default -o ./backup-theme.json
   ```

When bootstrapping a tenant marked `"is_default": true` (or `"slug": "default"`), any unset `OAUTH_CLIENT_ID`, `OAUTH_AUTHORITY`, `OAUTH_CLIENT_SECRET`, `DOMAIN`, and `SEARCH_*` values are automatically inherited from environment variables so secrets stay out of your theme JSON files.

## License

MeetEasier is licensed under the [GNU General Public License v3.0](LICENSE).
