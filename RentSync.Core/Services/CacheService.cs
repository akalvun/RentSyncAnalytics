using DuckDB.NET.Data;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using RentSync.Core.Configuration;
using RentSync.Core.Models;
using RentSync.Core.Telemetry;

namespace RentSync.Core.Services;

public interface ICacheService : IDisposable
{
    Task InitializeAsync(CancellationToken cancellationToken = default);
    Task StoreAsync(IReadOnlyList<RentRollRecord> records, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<RentRollRecord>> LoadAsync(CancellationToken cancellationToken = default);
    Task<bool> IsFreshAsync(CancellationToken cancellationToken = default);
}

/// <summary>
/// Local cache backed by DuckDB (in-process OLAP database).
///
/// Caching strategy: full snapshot replace inside a transaction, plus a
/// metadata table with the snapshot timestamp. The add-in reads from cache
/// when the snapshot is younger than CacheOptions.Freshness and only hits
/// the API when data is stale — that is the "minimise download times"
/// requirement in practice. DuckDB (rather than SQLite) so that analytical
/// aggregations run as SQL directly over the cache.
/// </summary>
public sealed class CacheService : ICacheService
{
    private readonly DuckDBConnection _connection;
    private readonly ITelemetryService _telemetry;
    private readonly ILogger<CacheService> _logger;
    private readonly TimeSpan _freshness;

    public CacheService(IOptions<CacheOptions> options, ITelemetryService telemetry,
        ILogger<CacheService> logger)
    {
        _telemetry = telemetry;
        _logger = logger;
        _freshness = options.Value.Freshness;

        var path = options.Value.DatabasePath;
        if (path != ":memory:")
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        _connection = new DuckDBConnection($"DataSource={path}");
    }

    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        await _connection.OpenAsync(cancellationToken).ConfigureAwait(false);

        using var cmd = _connection.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS rent_roll (
                property_id   VARCHAR NOT NULL,
                property_name VARCHAR NOT NULL,
                unit_id       VARCHAR NOT NULL,
                tenant_name   VARCHAR NOT NULL,
                monthly_rent  DECIMAL(18,2) NOT NULL,
                area_sqm      DECIMAL(18,2) NOT NULL,
                lease_start   DATE NOT NULL,
                lease_end     DATE NOT NULL
            );
            CREATE TABLE IF NOT EXISTS cache_meta (
                key VARCHAR PRIMARY KEY,
                value VARCHAR NOT NULL
            );
            """;
        await cmd.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
        _logger.LogInformation("Cache initialized at {Path}", _connection.ConnectionString);
    }

    public async Task StoreAsync(IReadOnlyList<RentRollRecord> records,
        CancellationToken cancellationToken = default)
    {
        using var op = _telemetry.StartOperation("Cache.Store",
            new Dictionary<string, object?> { ["recordCount"] = records.Count });
        try
        {
            using var tx = _connection.BeginTransaction();

            using (var clear = _connection.CreateCommand())
            {
                clear.Transaction = tx;
                clear.CommandText = "DELETE FROM rent_roll;";
                await clear.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            }

            // DuckDB appender = bulk insert path, orders of magnitude faster
            // than row-by-row INSERT for large rent rolls.
            using (var appender = _connection.CreateAppender("rent_roll"))
            {
                foreach (var r in records)
                {
                    appender.CreateRow()
                        .AppendValue(r.PropertyId)
                        .AppendValue(r.PropertyName)
                        .AppendValue(r.UnitId)
                        .AppendValue(r.TenantName)
                        .AppendValue(r.MonthlyRent)
                        .AppendValue(r.AreaSqm)
                        .AppendValue(r.LeaseStart.ToDateTime(TimeOnly.MinValue))
                        .AppendValue(r.LeaseEnd.ToDateTime(TimeOnly.MinValue))
                        .EndRow();
                }
            }

            using (var meta = _connection.CreateCommand())
            {
                meta.Transaction = tx;
                meta.CommandText = """
                    INSERT OR REPLACE INTO cache_meta (key, value)
                    VALUES ('snapshot_utc', $ts);
                    """;
                var p = meta.CreateParameter();
                p.ParameterName = "ts";
                p.Value = DateTimeOffset.UtcNow.ToString("O");
                meta.Parameters.Add(p);
                await meta.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            }

            tx.Commit();
            op.Succeed();
        }
        catch (Exception ex)
        {
            op.Fail(ex);
            _telemetry.TrackError(ex, "Cache.Store");
            throw;
        }
    }

    public async Task<IReadOnlyList<RentRollRecord>> LoadAsync(
        CancellationToken cancellationToken = default)
    {
        using var op = _telemetry.StartOperation("Cache.Load");
        try
        {
            var result = new List<RentRollRecord>();
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = """
                SELECT property_id, property_name, unit_id, tenant_name,
                       monthly_rent, area_sqm, lease_start, lease_end
                FROM rent_roll
                ORDER BY property_name, unit_id;
                """;

            using var reader = await cmd.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
            while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                result.Add(new RentRollRecord
                {
                    PropertyId = reader.GetString(0),
                    PropertyName = reader.GetString(1),
                    UnitId = reader.GetString(2),
                    TenantName = reader.GetString(3),
                    MonthlyRent = reader.GetDecimal(4),
                    AreaSqm = reader.GetDecimal(5),
                    LeaseStart = DateOnly.FromDateTime(reader.GetDateTime(6)),
                    LeaseEnd = DateOnly.FromDateTime(reader.GetDateTime(7))
                });
            }

            op.SetProperty("recordCount", result.Count);
            op.Succeed();
            return result;
        }
        catch (Exception ex)
        {
            op.Fail(ex);
            _telemetry.TrackError(ex, "Cache.Load");
            throw;
        }
    }

    public async Task<bool> IsFreshAsync(CancellationToken cancellationToken = default)
    {
        using var cmd = _connection.CreateCommand();
        cmd.CommandText = "SELECT value FROM cache_meta WHERE key = 'snapshot_utc';";
        var value = await cmd.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false);

        if (value is not string s || !DateTimeOffset.TryParse(s, out var snapshot))
            return false;

        return DateTimeOffset.UtcNow - snapshot < _freshness;
    }

    public void Dispose() => _connection.Dispose();
}
