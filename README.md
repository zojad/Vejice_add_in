# Vejice - Comma Checker for Microsoft Word

A Word add-in for Slovenian comma checking (missing commas and incorrect comma placement).

Supported on:
- Word Desktop (Windows and Mac)
- Word Online (Microsoft 365)
- Linux is supported for Word Online testing only (Word Desktop flow is not available on Linux).

## Test On Your Device

Prerequisites:
- Node.js 18+ (Node 20 LTS recommended)
- npm
- Git
- Microsoft Word Desktop and/or Word Online account

1. Clone and install:

```bash
git clone https://github.com/zojad/Vejice_add_in.git
cd Vejice_add_in
npm install
```

2. If `.env` is missing, create it from the template:

PowerShell:
```powershell
Copy-Item .env.example .env
```

macOS/Linux:
```bash
cp .env.example .env
```

3. Review and update `.env` as needed.

Real API mode:
```env
VEJICE_API_KEY=your-api-key
VEJICE_USE_MOCK=false
VEJICE_API_URL=https://127.0.0.1:4001/api/postavi_vejice
VEJICE_LEMMAS_URL=https://127.0.0.1:4001/lemmas
```

No API key (local mock smoke test):
```env
VEJICE_USE_MOCK=true
```

4. Install local HTTPS certificates:

```bash
npx office-addin-dev-certs install
```

5. Start local services (dev server + both proxies):

```bash
npm run start:web:manual
```

Keep this terminal running.

## Desktop Word Test Flow

**Requires Word Desktop on Windows or macOS. If you are on Linux, use the Word Online Test Flow below.**

1. In a second terminal (same repo), run:

```bash
npm start
```

2. Word opens with the add-in sideloaded.
3. Open any document with Slovenian text.
4. Go to `Home tab -> CJVT Vejice` and open the task pane.
5. Click `Preveri vejice`.

## Word Online Test Flow (Manual Sideload)

Before testing, import/trust the generated dev certificate in the browser/environment you use for Word Online.

1. Keep `npm run start:web:manual` running.
2. Open Word Online.
3. Upload manifest: `Insert -> Add-ins -> My add-ins -> Upload my add-in`.
4. Select `docs/manifest.web.xml`.
5. Open the add-in and run `Preveri vejice`.

Alternative automatic web debugging:
```bash
npm run start:web
```

## Stop Commands

- Stop Desktop debugging:
```bash
npm run stop
```
- Stop Web debugging (if you used `npm run start:web`):
```bash
npm run stop:web
```
- Stop manual dev services (`start:web:manual`): `Ctrl+C` in that terminal.

## Available Commands

| Command | Purpose |
|---------|---------|
| `npm install` | Install dependencies |
| `npm run build:dev` | Development build with sourcemaps |
| `npm run build` | Production build (minified) |
| `npm run watch` | Watch for file changes and rebuild |
| `npm run dev-server` | Start webpack dev server only |
| `npm start` | Sideload in Word Desktop |
| `npm run start:web:manual` | Start dev server + `proxy:lemmas` + `proxy:vejice` |
| `npm run start:web` | Automatic Word Online debugging |
| `npm run stop` | Stop Word Desktop debugging |
| `npm run stop:web` | Stop Word Online debugging |
| `npm run validate` | Validate `docs/manifest.dev.xml` |
| `npm run validate:web` | Validate `docs/manifest.web.xml` |
| `npm run validate:prod` | Validate `docs/manifest.prod.xml` |
| `npm run validate:web:prod` | Validate `docs/manifest.web.prod.xml` |
| `npm run lint:fix` | Fix code style issues |

## Troubleshooting

Certificate issues:

```bash
npx office-addin-dev-certs uninstall
npx office-addin-dev-certs install
```

- For Word Online testing, import/trust the generated dev certificate in the browser/environment you are using.
- Do this even if no browser warning is shown; add-ins can fail with a generic load error when the certificate is not trusted correctly.

Cannot connect to `https://localhost:4001`:
- Make sure `npm run start:web:manual` is running.
- Confirm port `4001` is not blocked by firewall or another process.

API errors (`403`, timeout):
- Verify `VEJICE_API_KEY` in `.env`.
- Ensure both proxies are running (included in `start:web:manual`).
- For smoke testing without credentials, set `VEJICE_USE_MOCK=true`.

Add-in not loading:
- Re-upload `docs/manifest.web.xml` (Word Online).
- Run manifest validation:
```bash
npm run validate
npm run validate:web
```
- Check browser/Office console logs.

## Building For Production

```bash
npm run build
```

Update `docs/manifest.prod.xml` and `docs/manifest.web.prod.xml` with your production domain, then deploy `docs/`.

## License
