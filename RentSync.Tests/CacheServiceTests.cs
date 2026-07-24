using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using RentSync.Core.Configuration;
using RentSync.Core.Models;
using RentSync.Core.Services;
using Xunit;

namespace RentSync.Tests;

public class CacheServiceTests : IAsyncLifetime
{
    private readonly FakeTelemetryService _telemetry = new();
    private CacheService _cache = null!;

    public async Task InitializeAsync()
    {
        _cache = new CacheService(
            Options.Create(new CacheOptions
            {
                DatabasePath = ":memory:",
                Freshness = TimeSpan.FromMinutes(30)
            }),
            _telemetry,
            NullLogger<CacheService>.Instance);
        await _cache.InitializeAsync();
    }

    public Task DisposeAsync()
    {
        _cache.Dispose();
        return Task.CompletedTask;
    }

    private static RentRollRecord Record(string propertyId, string unit,
        decimal rent, decimal area, string leaseEnd) => new()
    {
        PropertyId = propertyId,
        PropertyName = propertyId == "P1" ? "Riverside Plaza" : "Harbour Point",
        UnitId = unit,
        TenantName = "Tenant " + unit,
        MonthlyRent = rent,
        AreaSqm = area,
        LeaseStart = new DateOnly(2024, 1, 1),
        LeaseEnd = DateOnly.Parse(leaseEnd)
    };

    [Fact]
    public async Task StoreAndLoad_RoundTripsRecords()
    {
        var input = new[]
        {
            Record("P1", "U1", 2500m, 125m, "2027-12-31"),
            Record("P1", "U2", 1800m, 90m, "2026-10-15"),
            Record("P2", "A1", 4200m, 300m, "2028-06-30")
        };

        await _cache.StoreAsync(input);
        var loaded = await _cache.LoadAsync();

        Assert.Equal(3, loaded.Count);
        Assert.Equal(input.OrderBy(r => r.PropertyName).ThenBy(r => r.UnitId), loaded);
    }

    [Fact]
    public async Task Store_ReplacesPreviousSnapshot()
    {
        await _cache.StoreAsync(new[] { Record("P1", "U1", 2500m, 125m, "2027-12-31") });
        await _cache.StoreAsync(new[] { Record("P2", "A1", 4200m, 300m, "2028-06-30") });

        var loaded = await _cache.LoadAsync();

        var single = Assert.Single(loaded);
        Assert.Equal("P2", single.PropertyId);
    }

    [Fact]
    public async Task IsFresh_FalseWhenEmpty_TrueAfterStore()
    {
        Assert.False(await _cache.IsFreshAsync());
        await _cache.StoreAsync(new[] { Record("P1", "U1", 2500m, 125m, "2027-12-31") });
        Assert.True(await _cache.IsFreshAsync());
    }

    [Fact]
    public async Task Analytics_PropertySummaries_AggregatesCorrectly()
    {
        await _cache.StoreAsync(new[]
        {
            Record("P1", "U1", 2500m, 125m, "2027-12-31"),
            Record("P1", "U2", 1800m, 90m, "2026-10-15"),
            Record("P2", "A1", 4200m, 300m, "2028-06-30")
        });
        var analytics = new RentAnalyticsService(_cache, _telemetry);

        var summaries = await analytics.GetPropertySummariesAsync();

        Assert.Equal(2, summaries.Count);
        var p1 = summaries.Single(s => s.PropertyId == "P1");
        Assert.Equal(2, p1.UnitCount);
        Assert.Equal(4300m, p1.TotalMonthlyRent);
        Assert.Equal(20.00m, p1.AvgRentPerSqm); // 4300 / 215
    }

    [Fact]
    public async Task Analytics_ExpiringLeases_FiltersByWindow()
    {
        await _cache.StoreAsync(new[]
        {
            Record("P1", "U1", 2500m, 125m, "2027-12-31"),
            Record("P1", "U2", 1800m, 90m, "2026-10-15"),
            Record("P2", "A1", 4200m, 300m, "2028-06-30")
        });
        var analytics = new RentAnalyticsService(_cache, _telemetry);

        var expiring = await analytics.GetExpiringLeasesAsync(
            asOf: new DateOnly(2026, 7, 1), withinMonths: 6);

        var single = Assert.Single(expiring);
        Assert.Equal("U2", single.UnitId);
    }
}
