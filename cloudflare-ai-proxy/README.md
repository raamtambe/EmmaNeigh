# EmmaNeigh Cloudflare AI Proxy

This folder contains a Cloudflare Worker that provides the managed AI backend for EmmaNeigh without requiring Firebase Functions.

The Worker holds the provider API key and exposes the same endpoint contract the desktop app already expects:

- `POST /session`
- `GET|POST /health`
- `POST /prompt`
- `POST /agent`

## Why Use This

Use this path when:

- users should not enter their own API keys
- you do not want to run a local model
- you want a free / low-friction managed AI backend for beta use

The Worker keeps the LLM provider key in Cloudflare secrets and returns short-lived session tokens to the desktop app.

## Provider Support

Current implementation:

- `Groq` backend
- default model: `qwen/qwen3-32b`

The endpoint contract is model/provider-agnostic, so the Worker can later be adapted to Anthropic, OpenAI, or another provider.

## Files

- `src/index.js`: Worker implementation
- `wrangler.toml`: Worker configuration
- `.dev.vars.example`: local secret/env example
- `desktop-managed-config.example.json`: EmmaNeigh managed AI config example

## Required Cloudflare Secrets

Set these with Wrangler:

- `GROQ_API_KEY`
- `SESSION_SIGNING_KEY`

## Optional Cloudflare Vars

These can be placed in `wrangler.toml` or set through the Cloudflare dashboard:

- `MANAGED_AI_PROVIDER`
- `MANAGED_AI_MODEL`
- `MANAGED_AI_SESSION_TTL_SECONDS`
- `MANAGED_AI_ALLOW_MODEL_OVERRIDE`
- `ALLOWED_EMAILS`
- `ALLOWED_DOMAINS`
- `ALLOWED_USERS`
- `BLOCKED_EMAILS`
- `BLOCKED_DOMAINS`
- `BLOCKED_USERS`
- `ADMIN_EMAILS`
- `ADMIN_DOMAINS`
- `ADMIN_USERS`
- `ALLOWED_AI_PROVIDERS`
- `BLOCK_ALL`
- `ACCESS_POLICY_MESSAGE`

## Local Development

1. Install dependencies:

```bash
cd cloudflare-ai-proxy
npm install
```

2. Copy `.dev.vars.example` to `.dev.vars` and fill in values.

3. Run locally:

```bash
npm run dev
```

## Deploy

```bash
cd cloudflare-ai-proxy
npx wrangler login
npx wrangler secret put GROQ_API_KEY
npx wrangler secret put SESSION_SIGNING_KEY
npx wrangler deploy
```

After deploy, Wrangler will print a `workers.dev` URL.

## EmmaNeigh Desktop Config

Create a `telemetry-config.json` for the desktop app using the example in `desktop-managed-config.example.json`.

Example:

```json
{
  "allowedAiProviders": ["managed"],
  "managedAi": {
    "provider": "managed",
    "model": "qwen/qwen3-32b",
    "sessionUrl": "https://your-worker.workers.dev/session",
    "promptUrl": "https://your-worker.workers.dev/prompt",
    "agentUrl": "https://your-worker.workers.dev/agent",
    "healthUrl": "https://your-worker.workers.dev/health"
  }
}
```

For local testing, save that file as:

- `desktop-app/telemetry-config.json`

For packaged builds, ship the same file in the app resources directory so EmmaNeigh can read it at startup.

## Important Caveat

This removes Firebase from the AI path, but it does not magically create verified user identity.

Today, EmmaNeigh still uses email-only identity in the desktop app. The Worker improves key handling by issuing short-lived session tokens, but a later enterprise version should still move to verified email or SSO if identity assurance matters.
