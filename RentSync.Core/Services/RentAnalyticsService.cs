using RentSync.Core.Models;
using RentSync.Core.Telemetry;

namespace RentSync.Core.Services;

public sealed record PropertySummary(
    string PropertyId,
    string PropertyName,
    int UnitCount,
    decimal TotalMonthlyRent,
    decimal TotalAreaSqm,
    decimal AvgRentPerSqm);

public interface IRentAnalyticsService
{
    Task<IReadOnlyList<PropertySummary>> GetPropertySummariesAsync(
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<RentRollRecord>> GetExpiringLeasesAsync(
        DateOnly asOf, int withinMonths, CancellationToken cancellationToken = default);
}

/// <summary>
/// Read-model on top of the cache. Kept LINQ-based over ICacheService.LoadAsync
/// (rather than raw SQL) so the logic is trivially unit-testable with an
/// in-memory cache; heavy aggregations could be pushed down into DuckDB SQL
/// without changing this interface.
/// </summary>
public sealed class RentAnalyticsService : IRentAnalyticsService
{
    private readonly ICacheService _cache;
    private readonly ITelemetryService _telemetry;

    public RentAnalyticsService(ICacheService cache, ITelemetryService telemetry)
    {
        _cache = cache;
        _telemetry = telemetry;
    }

    public async Task<IReadOnlyList<PropertySummary>> GetPropertySummariesAsync(
        CancellationToken cancellationToken = default)
    {
        using var op = _telemetry.StartOperation("Analytics.PropertySummaries");
        var records = await _cache.LoadAsync(cancellationToken).ConfigureAwait(false);

        var result = records
            .GroupBy(r => (r.PropertyId, r.PropertyName))
            .Select(g =>
            {
                var totalArea = g.Sum(r => r.AreaSqm);
                var totalRent = g.Sum(r => r.MonthlyRent);
                return new PropertySummary(
                    g.Key.PropertyId,
                    g.Key.PropertyName,
                    g.Count(),
                    totalRent,
                    totalArea,
                    totalArea > 0 ? Math.Round(totalRent / totalArea, 2) : 0m);
            })
            .OrderBy(s => s.PropertyName)
            .ToList();

        op.SetProperty("propertyCount", result.Count);
        op.Succeed();
        return result;
    }

    public async Task<IReadOnlyList<RentRollRecord>> GetExpiringLeasesAsync(
        DateOnly asOf, int withinMonths, CancellationToken cancellationToken = default)
    {
        using var op = _telemetry.StartOperation("Analytics.ExpiringLeases",
            new Dictionary<string, object?> { ["withinMonths"] = withinMonths });

        var records = await _cache.LoadAsync(cancellationToken).ConfigureAwait(false);
        var result = records
            .Where(r => r.IsExpiringWithin(asOf, withinMonths))
            .OrderBy(r => r.LeaseEnd)
            .ToList();

        op.SetProperty("expiringCount", result.Count);
        op.Succeed();
        return result;
    }
}
