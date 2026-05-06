# Vejice Word Add-in

A Microsoft Word add-in for reviewing Slovenian comma placement.

This fork is rebuilt to be safer and easier to evaluate:

- safe-by-default service configuration
- no committed local env state
- explicit separation between mock, trusted remote, and unsafe legacy services
- basic automated tests for pure logic

## Current scope

The add-in can run in two modes:

- `mock` mode: default, no document text leaves the add-in
- `remote` mode: opt-in, for either a trusted internal API or the legacy unsafe API

For anchoring, the default is `synthetic` mode, which avoids the remote lemmatizer unless you explicitly enable it.

## Safety model

This repo assumes the existing external comma-correction API and lemmatizer may be unsafe.

Because of that:

- `VEJICE_CORRECTION_PROVIDER=mock` is the default
- `VEJICE_LEMMATIZER_PROVIDER=synthetic` is the default
- remote correction is blocked unless you mark it trusted or explicitly allow the unsafe legacy API
- remote lemmatization is blocked unless you mark it trusted or explicitly allow the unsafe legacy lemmatizer

Examples:

Trusted internal correction API:

```env
VEJICE_CORRECTION_PROVIDER=remote
VEJICE_REMOTE_CORRECTION_TRUSTED=true
VEJICE_API_URL=https://your-domain.example/api/postavi_vejice
```

Unsafe legacy correction API:

```env
VEJICE_CORRECTION_PROVIDER=remote
VEJICE_ALLOW_UNSAFE_REMOTE_API=true
VEJICE_API_URL=https://127.0.0.1:4001/api/postavi_vejice
VEJICE_API_KEY=replace-me
```

Trusted internal lemmatizer:

```env
VEJICE_LEMMATIZER_PROVIDER=remote
VEJICE_REMOTE_LEMMATIZER_TRUSTED=true
VEJICE_LEMMAS_URL=https://your-domain.example/lemmas
```

Unsafe legacy lemmatizer:

```env
VEJICE_LEMMATIZER_PROVIDER=remote
VEJICE_ALLOW_UNSAFE_LEMMATIZER=true
VEJICE_LEMMAS_URL=https://127.0.0.1:4001/lemmas
```

## Prerequisites

- Node.js 20+
- npm
- Microsoft Word Desktop or Word Online

## Setup

1. Install dependencies:

```bash
npm install
```

2. Internal testing can run with no local secrets and no remote services. You can start immediately with the built-in safe defaults.

3. Optional: create a local config file only if you want to override the defaults:

```bash
cp .env.example .env.local
```

4. Keep the safe defaults unless you intentionally want remote services.

5. If you want Office/browser trust to be installed system-wide, install Office dev certificates manually:

```bash
npx office-addin-dev-certs install
```

If you skip this step, the dev server will still start with a local self-signed HTTPS certificate and you can trust it manually for Word Online testing.

## Internal testing

Fastest safe path:

```bash
npm run start:internal
```

This starts only the HTTPS dev server in mock mode. No document text leaves the add-in.

If you want the full local stack for later remote-service testing, start the dev server and local proxies:

```bash
npm run start:web:manual
```

Desktop Word:

```bash
npm start
```

Word Online:

1. Keep `npm run start:web:manual` running.
2. Open Word Online.
3. Upload `docs/manifest.web.xml`.
4. Open the add-in and run the checker.

## Verification

Run the internal submission checks:

```bash
npm run verify:internal
```

Additional helper:

```bash
npm run check:encoding
```

## Scripts

| Command | Purpose |
| --- | --- |
| `npm test` | Run unit tests for pure logic and service config |
| `npm run build` | Production build |
| `npm run build:dev` | Development build |
| `npm run watch` | Watch mode |
| `npm run dev-server` | Start webpack dev server only |
| `npm run start:internal` | Start the add-in in safe mock mode for internal testing |
| `npm start` | Sideload into desktop Word |
| `npm run start:web` | Start automatic web debugging |
| `npm run start:web:manual` | Start local dev server and proxies |
| `npm run stop` | Stop desktop debugging |
| `npm run stop:web` | Stop web debugging |
| `npm run validate` | Validate desktop manifest |
| `npm run validate:web` | Validate web manifest |
| `npm run verify:internal` | Run tests, encoding check, build, and manifest validation |
| `npm run lint` | Run Office add-in lint checks |
| `npm run lint:fix` | Apply lint fixes |
| `npm run check:encoding` | Scan source files for mojibake and oversized lines |

## Repo notes

- `.env` and `.env.local` are intentionally ignored and should stay local.
- `docs/` is used as the web build output for the add-in.
- `dist/` is not part of the intended tracked source workflow anymore.

## Next step if you replace the unsafe API

If you later build a safe internal correction API, you should only need to:

1. point `VEJICE_API_URL` at the new endpoint
2. set `VEJICE_CORRECTION_PROVIDER=remote`
3. set `VEJICE_REMOTE_CORRECTION_TRUSTED=true`

If you also replace the lemmatizer, do the equivalent with:

1. `VEJICE_LEMMAS_URL`
2. `VEJICE_LEMMATIZER_PROVIDER=remote`
3. `VEJICE_REMOTE_LEMMATIZER_TRUSTED=true`
