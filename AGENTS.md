# Repository Guidelines

## Project Structure & Module Organization
Keep Worker source logic in `src/index.ts`. Cloudflare bindings, cron, and deployment settings live in `wrangler.jsonc`; generated binding types live in `worker-configuration.d.ts`. Operational notes belong in `docs/cloudflare_worker.md`, and the user-facing overview lives in `docs/readme.md`.

Runtime backup artifacts such as `data/`, `downloads/`, `.wrangler/`, dependency folders, and logs are ignored and should stay out of commits.

## Build, Test, and Development Commands
Install dependencies from the lockfile with `npm install`. Run `npm run check` before deploying to type-check the Worker. Use `npm run dev` for local HTTP testing, `npm run dev:scheduled` to exercise the scheduled handler locally, and `npm run deploy` to publish.

After changing Cloudflare bindings or config, regenerate local binding types with `npx wrangler types`.

## Coding Style & Naming Conventions
Write TypeScript ESM using the existing strict compiler settings. Keep parsing, fetch, storage, and route helpers small and descriptive. Prefer platform APIs available in Workers over adding server-side dependencies. Keep D1 table names and R2 object keys stable because remote data has already been seeded.

## Testing Guidelines
Use `npm run check` as the required smoke test for code changes. For behavior changes, also run the relevant local route with `npm run dev` or the scheduled route with `npm run dev:scheduled`. When checking deployed state, use `/state` and `/sync-latest` on the Worker URL.

## Commit & Pull Request Guidelines
Use a single-line, imperative subject no longer than 72 characters, optionally prefixed with a scope such as `feat:` or `chore:`. Pull requests should summarize Worker behavior changes, list verification commands, and call out any D1/R2 data impact.

## Data Handling Tips
The current deployment stores structured rows in D1 and downloaded attachments in R2. Do not commit local data backups, logs, or generated runtime state. Preserve the existing request throttling and remote availability handling when extending the crawler.
