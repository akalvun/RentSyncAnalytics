# RentSync.Core v2 — production-grade rewrite

## What changed vs the v1 skeleton

**Telemetry / observability (the differentiator):**
- `ITelemetryService` + `TelemetryService`: buffered via a bounded `Channel<T>`
  (DropOldest), single background drain task, local JSONL sink. Never blocks
  the Excel UI thread, never crashes the host on sink failure.
- Correlation: every event carries a `SessionId`; timed operations carry an
  `OperationId`, duration, success flag, and custom properties (attempt count,
  record count).
- Privacy-aware error tracking: exception type + message only, full stack goes
  to the local Serilog log, not into telemetry.
- Structured logging: Serilog → compact JSON (CLEF), rolling daily files,
  14-day retention. Directly readable by Seq / ELK.

**REST client:**
- Exponential backoff with **full jitter**, honours `Retry-After` on 429/503.
- Fail-fast on non-transient errors (401/404/JSON) — no wasted retries.
- `IHttpClientFactory` registration (no socket exhaustion / stale DNS).
- Auth via `IAuthTokenProvider` seam — Windows Credential Manager
  implementation stays in the VSTO layer; tests inject a fake.

**Cache (DuckDB):**
- Transactional snapshot replace + `cache_meta` freshness timestamp.
- Bulk insert via DuckDB **Appender** instead of row-by-row INSERT.
- `IsFreshAsync` implements the cache-first strategy ("minimise download
  times" from the job ad).

**Everything else:**
- .NET 8, nullable enabled, `TreatWarningsAsErrors`, DI composition root
  (`AddRentSyncCore`), NuGet metadata + SemVer in the csproj.
- 12 xUnit tests: retry behaviour, fail-fast, telemetry assertions,
  DuckDB round-trip (`:memory:`), aggregations, lease-expiry window.

## Build & test (local, VS 2026 or CLI)

```
dotnet restore
dotnet build
dotnet test
```

No Office/Windows dependency in Core — tests run anywhere with .NET 8 SDK.
(The VSTO AddIn project is the next step and does require Windows + Office.)

## Next step

Rewrite `RentSync.AddIn` (.NET Framework 4.8) to consume this Core:
`ThisAddIn_Startup` builds the DI container, ribbon callbacks marshal async
results back to the UI thread, Credential Manager token provider,
`FlushAsync` on shutdown.
