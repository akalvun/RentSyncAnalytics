namespace RentSync.Core.Telemetry;

/// <summary>
/// Client-side telemetry for a desktop/Office add-in context.
///
/// Design goals (mirrors what a production Excel add-in needs):
///  - Offline-first: events are buffered locally and never block the UI thread.
///  - Privacy-aware: no PII; only feature names, timings, error types, anonymous session ids.
///  - Correlated: every event carries a SessionId, and operations carry an OperationId
///    so a slow REST call can be tied to the exact ribbon click that triggered it.
/// </summary>
public interface ITelemetryService : IAsyncDisposable
{
    /// <summary>Anonymous id for the current application session.</summary>
    Guid SessionId { get; }

    /// <summary>Record a discrete usage event, e.g. "Ribbon.RefreshClicked".</summary>
    void TrackEvent(string name, IReadOnlyDictionary<string, object?>? properties = null);

    /// <summary>Record a handled exception with its context. Never throws.</summary>
    void TrackError(Exception exception, string context,
        IReadOnlyDictionary<string, object?>? properties = null);

    /// <summary>
    /// Start a timed operation. Dispose the returned scope to record duration and outcome:
    /// <code>
    /// using var op = telemetry.StartOperation("Api.FetchRentRoll");
    /// ... work ...
    /// op.Succeed();   // omit -> recorded as failed
    /// </code>
    /// </summary>
    IOperationScope StartOperation(string name,
        IReadOnlyDictionary<string, object?>? properties = null);

    /// <summary>Flush buffered events to the local sink (call on add-in shutdown).</summary>
    Task FlushAsync(CancellationToken cancellationToken = default);
}

/// <summary>A timed, correlated unit of work.</summary>
public interface IOperationScope : IDisposable
{
    Guid OperationId { get; }
    void Succeed();
    void Fail(Exception? exception = null);
    void SetProperty(string key, object? value);
}
