# EmmaNeigh Managed Deployment Backend

This folder contains the minimum backend scaffolding to move EmmaNeigh away from user-managed local trust and into a managed deployment model.

## What It Provides

1. `telemetryIngest`
   - receives usage, feedback, and prompt telemetry from the desktop app
   - writes events to Firestore using the Admin SDK

2. `accessPolicy`
   - returns allow/block policy and admin-role data to the desktop app
   - supports a global policy doc plus identity-specific override docs

3. `managedAiSession`, `managedAiHealth`, `managedAiPrompt`, `managedAiAgent`
   - provide a server-held AI integration path
   - keep the Groq API key on the backend, not in end-user machines
   - issue short-lived managed AI session tokens after backend policy checks
   - allow the desktop app to use a managed remote model without asking users for provider tokens

4. locked-down Firestore rules
   - client access is denied by default
   - backend writes happen through the Admin SDK

## Firestore Collections

Create these collections in Firestore:

- `emmaneigh_policy`
  - doc: `global`
  - fields can include:
    - `allowed_emails`
    - `allowed_domains`
    - `allowed_users`
    - `blocked_emails`
    - `blocked_domains`
    - `blocked_users`
    - `admin_emails`
    - `admin_domains`
    - `admin_users`
    - `message`
    - `allowed_ai_providers`
    - `block_all`

- `emmaneigh_policy_overrides`
  - optional docs:
    - `email:<url-encoded-email>`
    - `username:<url-encoded-username>`
    - `userid:<url-encoded-userid>`
  - same field shape as the global policy doc

- `emmaneigh_events`
  - backend writes telemetry events here automatically

## Required Secrets

Create these Firebase Functions secrets:

- `EMMANEIGH_TELEMETRY_TOKEN`
- `EMMANEIGH_ACCESS_POLICY_TOKEN`
- `EMMANEIGH_SESSION_SIGNING_KEY`
- `EMMANEIGH_GROQ_API_KEY`

`EMMANEIGH_SESSION_SIGNING_KEY` signs the short-lived managed AI session tokens returned by `managedAiSession`. Unlike the older shared-token model, this secret does not ship in the desktop app bundle.

## Local Setup

1. Create a Firebase project
2. Enable Firestore
3. Copy `.firebaserc.example` to `.firebaserc` and set your project id
4. In `deployment-backend/functions` run `npm install`
5. Set the functions secrets:
   - `firebase functions:secrets:set EMMANEIGH_TELEMETRY_TOKEN`
   - `firebase functions:secrets:set EMMANEIGH_ACCESS_POLICY_TOKEN`
   - `firebase functions:secrets:set EMMANEIGH_SESSION_SIGNING_KEY`
   - `firebase functions:secrets:set EMMANEIGH_GROQ_API_KEY`
6. Deploy:
   - `firebase deploy --only functions,firestore:rules`

## Desktop App Managed Config

After deploy, create a real `telemetry-config.json` from `telemetry-config.example.json` and ship it with the desktop app or inject the same values through environment variables.

Recommended config shape:

```json
{
  "telemetryIngestUrl": "https://.../telemetryIngest",
  "accessPolicyUrl": "https://.../accessPolicy",
  "accessPolicyFailClosed": true,
  "allowedAiProviders": ["managed", "anthropic", "openai"],
  "managedAi": {
    "provider": "managed",
    "model": "qwen/qwen3-32b",
    "sessionUrl": "https://.../managedAiSession",
    "promptUrl": "https://.../managedAiPrompt",
    "agentUrl": "https://.../managedAiAgent",
    "healthUrl": "https://.../managedAiHealth"
  },
  "firebaseConfig": {
    "apiKey": "...",
    "authDomain": "...",
    "projectId": "...",
    "storageBucket": "...",
    "messagingSenderId": "...",
    "appId": "..."
  }
}
```

The managed AI block no longer needs a long-lived shared client token. EmmaNeigh signs in with the user’s email identity, requests a short-lived managed AI session from the backend, and keeps that token only in memory.

## Accounts To Create

For the current EmmaNeigh enterprise path, the admin should provision:

1. Firebase project
2. Firestore database
3. Firebase Functions / Google Cloud project hosting
4. Secret Manager-backed function secrets
5. Optional later: Microsoft Entra ID tenant/app registration for SSO
6. Optional later: Apple developer + Windows code-signing setup for release hardening

## Identity Provider Note

This backend does not yet implement Microsoft Entra SSO. The right next step is to use Entra for administrator authentication and eventually user authentication, while keeping the app's current email identity only as a lightweight transitional mode.

## Managed AI Note

The current managed AI backend is implemented against Groq first, with the desktop app talking only to the backend endpoints. That gives you a free remote starting point without putting provider keys on user machines.

The app-side contract is model/provider-agnostic, so later you can swap the backend implementation to Anthropic, OpenAI, or another provider without changing the desktop UX. The current implementation now issues short-lived backend session tokens for AI access instead of shipping a long-lived shared AI token in the app bundle.

Important limitation: the desktop app still uses email-only identity today. That means the session token handling is now materially better, but true identity assurance still requires verified email or SSO.
