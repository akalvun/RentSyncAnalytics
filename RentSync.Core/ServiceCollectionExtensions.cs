using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using RentSync.Core.Configuration;
using RentSync.Core.Services;
using RentSync.Core.Telemetry;
using Serilog;
using Serilog.Formatting.Compact;

namespace RentSync.Core;

/// <summary>
/// Single composition root for the Core services. The VSTO layer calls
/// AddRentSyncCore() from ThisAddIn_Startup; tests call it with overrides.
/// </summary>
public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddRentSyncCore(
        this IServiceCollection services,
        Action<ApiOptions>? configureApi = null,
        Action<CacheOptions>? configureCache = null,
        Action<TelemetryOptions>? configureTelemetry = null)
    {
        services.AddOptions<ApiOptions>();
        services.AddOptions<CacheOptions>();
        services.AddOptions<TelemetryOptions>();
        if (configureApi is not null) services.Configure(configureApi);
        if (configureCache is not null) services.Configure(configureCache);
        if (configureTelemetry is not null) services.Configure(configureTelemetry);

        services.AddSingleton<ITelemetryService, TelemetryService>();
        services.AddSingleton<ICacheService, CacheService>();
        services.AddSingleton<IRentAnalyticsService, RentAnalyticsService>();
        services.AddSingleton<IRentRollDataService, RentRollDataService>();
        services.AddSingleton<IAuthTokenProvider, AnonymousTokenProvider>();

        // IHttpClientFactory: correct handler lifetime management (avoids the
        // classic socket-exhaustion / stale-DNS pitfalls of new HttpClient()).
        services.AddHttpClient<IRestApiClient, RestApiClient>();

        return services;
    }

    /// <summary>
    /// Structured logging: compact JSON lines (CLEF format), rolling daily files.
    /// Readable by Seq/ELK out of the box; grep-able as plain text in a pinch.
    /// </summary>
    public static IServiceCollection AddRentSyncLogging(
        this IServiceCollection services, string logDirectory)
    {
        Directory.CreateDirectory(logDirectory);

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.WithProperty("app", "RentSync")
            .WriteTo.File(
                new CompactJsonFormatter(),
                Path.Combine(logDirectory, "rentsync-.clef"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 14)
            .CreateLogger();

        services.AddLogging(builder =>
        {
            builder.ClearProviders();
            builder.AddSerilog(Log.Logger, dispose: true);
        });

        return services;
    }
}
