# Backend-to-n8n Integration Contract

## Purpose and authority

This document defines the backend/n8n boundary. Website and Android clients call only the backend.

- PostgreSQL is authoritative for users, files, approvals, prices, bookings, payments, jobs, and accepted results.
- Backend owns authentication, authorization, validation, state transitions, job creation, and result acceptance.
- n8n owns temporary workflow execution and explicitly authorized external side effects.
- Stripe is authoritative for payment transactions after backend webhook verification.
- Google Calendar is authoritative for the external event; backend stores the returned event ID/state.
- Human decisions remain separate from AI recommendations.

Executable OpenAPI/JSON Schemas and fixtures under `packages/contracts/` are the implementation source of truth. Keep them consistent with this overview.

CV-related fields and visibility follow `cv-profile-schema.md`. The approved structured profile is authoritative; the uploaded CV and n8n result are supporting evidence/proposals.

## Transport and service authentication

Use HTTPS and an approved service-authentication method such as signed JWT or HMAC. Require:

- Separate keys per environment with rotation support.
- Body-integrity verification.
- Timestamp, short acceptance window, nonce/idempotency key, and replay rejection.
- Audience/issuer validation where applicable.
- Allowlisted trigger/callback origins, rate limits, and payload-size limits.

Never expose internal endpoints, webhook URLs, tokens, or keys to browser/mobile code.

## Trigger and acknowledgement

Backend creates the database job first, then sends a versioned asynchronous trigger containing:

| Field | Purpose |
| --- | --- |
| `schemaVersion` | Contract version |
| `jobId` / `jobType` | Authoritative job identity and workflow route |
| `jobVersion` | Optimistic concurrency/version check |
| `idempotencyKey` | Duplicate and side-effect protection |
| `correlationId` | Safe tracing identifier |
| `requestedAt` / `expiresAt` | Freshness and deadline |
| `actor` | Audit context; not authorization by itself |
| `resource` | Internal IDs and short-lived references |
| `parameters` | Small workflow-specific settings |

n8n validates the request before accepting it. Return a fast acknowledgement with `jobId`, accepted schema/workflow version, and acceptance time; long work completes by callback. Validation/authentication failures return a safe rejection and must not start processing.

## Job states

| State | Meaning |
| --- | --- |
| `queued` | Backend created job; no valid n8n acceptance yet |
| `accepted` | n8n accepted the job |
| `processing` | Work is active |
| `needs_review` | Human correction/approval is required |
| `succeeded` | Backend validated and stored the result |
| `failed_retryable` | Transient failure permits controlled retry |
| `failed_permanent` | Validation, authorization, policy, or unsupported-input failure |
| `cancelled` | Authorized cancellation completed |
| `expired` | Deadline passed without acceptable completion |

Allowed transitions must be encoded and tested. A callback may not revive a terminal job or update a different job/version. Backend alone commits final application state.

## Callback

Every callback includes:

- Schema, job, workflow, and correlation identifiers/versions.
- Idempotency key and expected `jobVersion`.
- Status, start/completion timestamps, result, safe warnings, and normalized error.

Backend must verify authentication, freshness, IDs, versions, current state, and result schema inside one concurrency-safe transaction. Acknowledge only after persistence. Repeated identical callbacks are safe; conflicting callbacks are rejected and audited.

Normalized errors contain `code`, `category`, `retryable`, and a safe message. Never return or log stack traces, secrets, access URLs, full prompts, full CV text, payment data, or private provider responses.

## File references

Files stay in private object storage. A job receives only:

- `fileId`, MIME type, size, expected checksum.
- Short-lived, single-purpose read reference and expiry.

The reference must be bound to the expected file where supported, excluded from logs/callbacks, and unusable after expiry. n8n returns structured/derived data, not the original file.

## Workflow-specific contract summary

### `consultant.cv.ingest`

Input: `consultantId`, current `profileId`/`profileVersion`, `fileId`, access reference, checksum, and extraction-config version.

Output: schema-aligned proposed additions/corrections, missing required fields, conflicts with the current profile, expertise suggestions, evidence locations, per-field confidence, optional suggested tier, and parser/model versions.

Backend stores proposals separately from approved data and normally maps completion to `needs_review`. It rejects unknown, excluded-sensitive, or visibility-changing fields and never applies proposals automatically.

### `knowledge.ingest`

Input: `knowledgeSourceId`, `fileId`, access reference, checksum, valid `approvalId`, target collection, chunking/embedding versions.

Output: source/version, chunk count, embedding model/configuration, index references, and warnings.

Backend must revalidate approval and target collection before accepting the result.

### `consultant.match`

Input: `requirementId`, sanitized category/topics/language/time zone/availability/summary, eligible consultant IDs with approved structured-profile versions, and matching-config version.

Output: candidate IDs, scores, reasons, limitations, evidence references, and model/config versions.

Backend stores this as a recommendation pending human confirmation.

### `consultant.rag.query`

Input: `consultantId`, validated query, permitted collection IDs, retrieval-config version.

Output: consultant-facing answer, stable source citations, warnings/insufficient-evidence indicator, and model/config versions.

Backend returns it only to the requesting consultant and does not add query/answer to permanent RAG.

### `calendar.event.manage`

Input: `bookingId`, operation (`create|update|cancel`), calendar connection reference, existing event ID where applicable, timestamps/time zone, safe title, permitted participants, and safe description.

Output: provider/event ID, event status, timestamps, and safe join URL.

For creation, backend records verified consultant acceptance, slot hold/booking, and payment before creating the job. Never send raw OAuth tokens in the payload.

### `notification.send`

Input: approved template ID/version, recipient reference, channel, safe merge fields, and application event ID.

Output: provider message ID, acceptance time, and delivery-acceptance status.

Do not include CV content, confidential client information, private file links, prompts, or payment credentials.

## Payments remain outside n8n

Backend creates payment sessions, calculates amounts, verifies Stripe webhooks, and changes payment state. n8n must not calculate charges, infer success from redirects, mark transactions, issue refunds/payouts, or choose a tier/price.

After confirmed payment, backend may create calendar or notification jobs containing only the required application state.

## Idempotency, concurrency, and cancellation

- Same idempotency key and same payload returns/preserves the original processing/result.
- Same key with a different payload is rejected.
- Ingestion identity includes source ID/version/checksum/collection.
- Calendar and notification retries use stored provider IDs or provider idempotency.
- Callback updates use `jobVersion` or equivalent optimistic locking.
- Cancellation is honored before irreversible side effects; after a side effect, return its actual state and perform only an explicitly authorized compensating action.
- If the side effect succeeded but callback failed, retry only the callback.

## Timeouts and retries

- Use asynchronous callbacks for long-running work.
- Apply bounded exponential backoff only to transient failures.
- Do not retry validation, authorization, policy, unsupported-input, expired, or budget-limit failures.
- Jobs must reach a terminal or needs-review state; never remain processing indefinitely.
- UI receives safe queued/processing/retry/failure states through the backend.

## Versioning and change control

- Version job schemas, callbacks, workflows, prompts, models, and processing configurations.
- Compatible optional additions may use a minor schema version; removals/renames/behavior changes require a new major version.
- Backend and n8n reject unsupported major versions.
- Contract fixtures cover every supported job/version.
- Promote changes development -> staging -> production after contract, integration, and synthetic evaluation tests pass.

## Observability

Safe logs may contain job/type/correlation IDs, versions, timestamps/duration, state, safe error code, and non-sensitive provider reference.

Do not log CV text, RAG query/answer, client summary, credentials, tokens, signed URLs, complete prompts, payment data, or file contents. Configure retention and access before production.

## Acceptance tests

The boundary passes only when tests prove:

- Browser/mobile cannot reach n8n through supported interfaces.
- Unsigned, expired, replayed, oversized, and invalid-version/schema requests are rejected.
- Duplicate triggers/callbacks do not duplicate side effects or corrupt state.
- Invalid callbacks cannot alter unrelated, stale, or terminal jobs.
- Long jobs expose useful states; failure, review, retry, cancellation, and expiry are recoverable.
- File references expire and do not leak.
- Backend retains payment and approval authority.
- Client information cannot enter permanent RAG.
- Executable contract fixtures pass in development and staging.
