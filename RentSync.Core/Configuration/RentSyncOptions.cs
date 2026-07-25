namespace RentSync.Core.Configuration;

public sealed class TelemetryOptions
{
    /// <summary>Folder for the local JSONL telemetry buffer.</summary>
    public string BufferDirectory { get; set; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RentSync", "telemetry");

    /// <summary>Max events held in memory before oldest are dropped.</summary>
    public int BufferCapacity { get; set; } = 10_000;
}

public sealed class ApiOptions
{
    /// <summary>Demo endpoint; the real product would point at the platform API.</summary>
    public string BaseUrl { get; set; } = "https://jsonplaceholder.typicode.com";

    public int MaxRetries { get; set; } = 3;

    /// <summary>Base delay for exponential backoff (doubles per attempt, plus jitter).</summary>
    public TimeSpan RetryBaseDelay { get; set; } = TimeSpan.FromMilliseconds(400);

    public TimeSpan RequestTimeout { get; set; } = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Demo mode: serve rent roll from a bundled local JSON file instead of the
    /// network. The real product would always call the API; this keeps the
    /// portfolio demo self-contained and offline-capable.
    /// </summary>
    public bool UseLocalSample { get; set; }

    /// <summary>Path to the bundled sample JSON when UseLocalSample is true.</summary>
    public string LocalSamplePath { get; set; } = "";
}

public sealed class CacheOptions
{
    /// <summary>DuckDB database file. ":memory:" is used by tests.</summary>
    public string DatabasePath { get; set; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RentSync", "cache.duckdb");

    /// <summary>Cached data older than this is considered stale.</summary>
    public TimeSpan Freshness { get; set; } = TimeSpan.FromMinutes(30);
}
