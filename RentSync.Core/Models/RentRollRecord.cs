namespace RentSync.Core.Models;

/// <summary>
/// One row of a rent roll: a lease on a unit within a property.
/// Immutable record type — safe to cache, pass across threads, and compare in tests.
/// </summary>
public sealed record RentRollRecord
{
    public required string PropertyId { get; init; }
    public required string PropertyName { get; init; }
    public required string UnitId { get; init; }
    public required string TenantName { get; init; }

    /// <summary>Monthly contracted rent in the portfolio base currency.</summary>
    public required decimal MonthlyRent { get; init; }

    /// <summary>Leased area in square metres.</summary>
    public required decimal AreaSqm { get; init; }

    public required DateOnly LeaseStart { get; init; }
    public required DateOnly LeaseEnd { get; init; }

    /// <summary>Rent per square metre; 0 when area is unknown.</summary>
    public decimal RentPerSqm => AreaSqm > 0 ? Math.Round(MonthlyRent / AreaSqm, 2) : 0m;

    public bool IsExpiringWithin(DateOnly asOf, int months) =>
        LeaseEnd >= asOf && LeaseEnd <= asOf.AddMonths(months);
}
