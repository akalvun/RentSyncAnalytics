namespace RentSync.Core.Telemetry;

public enum TelemetryEventKind
{
    Usage,
    Error,
    Operation
}

/// <summary>
/// One structured telemetry record. Serialized as a JSON line into the local
/// telemetry buffer file, ready to be shipped to any backend (App Insights,
/// Seq, ELK) — the demo keeps it local by design.
/// </summary>
public sealed record TelemetryEvent
{
    public required TelemetryEventKind Kind { get; init; }
    public required string Name { get; init; }
    public required DateTimeOffset TimestampUtc { get; init; }
    public required Guid SessionId { get; init; }
    public Guid? OperationId { get; init; }
    public double? DurationMs { get; init; }
    public bool? Success { get; init; }
    public string? ExceptionType { get; init; }
    public string? ExceptionMessage { get; init; }
    public IReadOnlyDictionary<string, object?>? Properties { get; init; }
}
