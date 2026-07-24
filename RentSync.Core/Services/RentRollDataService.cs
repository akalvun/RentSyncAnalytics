using Microsoft.Extensions.Logging;
using RentSync.Core.Models;
using RentSync.Core.Telemetry;

namespace RentSync.Core.Services;

public interface IRentRollDataService
{
    /// <summary>Cache-first: serve fresh cache, otherwise fetch, store, serve.</summary>
    Task<IReadOnlyList<RentRollRecord>> GetRentRollAsync(
        bool forceRefresh = false, CancellationToken cancellationToken = default);
}

/// <summary>
/// The "minimise download times" requirement as one testable class:
/// the add-in never talks to the API directly, it asks this service,
/// which decides between DuckDB cache and network.
/// </summary>
public sealed class RentRollDataService : IRentRollDataService
{
    private readonly IRestApiClient _api;
    private readonly ICacheService _cache;
    private readonly ITelemetryService _telemetry;
    private readonly ILogger<RentRollDataService> _logger;

    public RentRollDataService(IRestApiClient api, ICacheService cache,
        ITelemetryService telemetry, ILogger<RentRollDataService> logger)
    {
        _api = api;
        _cache = cache;
        _telemetry = telemetry;
        _logger = logger;
    }

    public async Task<IReadOnlyList<RentRollRecord>> GetRentRollAsync(
        bool forceRefresh = false, CancellationToken cancellationToken = default)
    {
        using var op = _telemetry.StartOperation("Data.GetRentRoll",
            new Dictionary<string, object?> { ["forceRefresh"] = forceRefresh });

        if (!forceRefresh && await _cache.IsFreshAsync(cancellationToken).ConfigureAwait(false))
        {
            var cached = await _cache.LoadAsync(cancellationToken).ConfigureAwait(false);
            op.SetProperty("source", "cache");
            op.Succeed();
            return cached;
        }

        try
        {
            var fetched = await _api.FetchRentRollAsync(cancellationToken).ConfigureAwait(false);
            await _cache.StoreAsync(fetched, cancellationToken).ConfigureAwait(false);
            op.SetProperty("source", "api");
            op.Succeed();
            return fetched;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            // Network down? Degrade gracefully to stale cache instead of failing the user.
            _logger.LogWarning(ex, "API fetch failed, falling back to stale cache");
            var stale = await _cache.LoadAsync(cancellationToken).ConfigureAwait(false);
            if (stale.Count == 0) throw; // nothing to fall back to
            op.SetProperty("source", "stale-cache");
            op.Succeed();
            return stale;
        }
    }
}
