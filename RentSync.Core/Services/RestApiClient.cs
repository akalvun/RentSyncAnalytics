using System.Net;
using System.Net.Http;               // implicit usings omit this on net48 targets
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using RentSync.Core.Configuration;
using RentSync.Core.Models;
using RentSync.Core.Telemetry;

namespace RentSync.Core.Services;

public interface IRestApiClient
{
    Task<IReadOnlyList<RentRollRecord>> FetchRentRollAsync(CancellationToken cancellationToken = default);
}

/// <summary>
/// Async REST client for the rent roll endpoint.
///
/// Retry policy: transient failures (5xx, 408, 429, network errors) are retried
/// up to MaxRetries with exponential backoff + full jitter — the jitter prevents
/// a fleet of add-ins from hammering a recovering server in lockstep.
/// Non-transient failures (401, 404, bad payload) fail fast: retrying them
/// only wastes the user's time.
///
/// Every request is a telemetry operation, so slow endpoints show up in the
/// usage data with exact timings and retry counts.
/// </summary>
public sealed class RestApiClient : IRestApiClient
{
    private readonly HttpClient _http;
    private readonly IAuthTokenProvider _tokenProvider;
    private readonly ITelemetryService _telemetry;
    private readonly ILogger<RestApiClient> _logger;
    private readonly ApiOptions _options;

    public RestApiClient(
        HttpClient http,
        IAuthTokenProvider tokenProvider,
        ITelemetryService telemetry,
        IOptions<ApiOptions> options,
        ILogger<RestApiClient> logger)
    {
        _http = http;
        _tokenProvider = tokenProvider;
        _telemetry = telemetry;
        _logger = logger;
        _options = options.Value;

        _http.BaseAddress = new Uri(_options.BaseUrl, UriKind.Absolute);
        _http.Timeout = _options.RequestTimeout;
    }

    public async Task<IReadOnlyList<RentRollRecord>> FetchRentRollAsync(
        CancellationToken cancellationToken = default)
    {
        using var op = _telemetry.StartOperation("Api.FetchRentRoll");
        try
        {
            var records = await SendWithRetryAsync(op, cancellationToken).ConfigureAwait(false);
            op.SetProperty("recordCount", records.Count);
            op.Succeed();
            return records;
        }
        catch (Exception ex)
        {
            op.Fail(ex);
            _telemetry.TrackError(ex, "Api.FetchRentRoll");
            throw;
        }
    }

    private async Task<IReadOnlyList<RentRollRecord>> SendWithRetryAsync(
        IOperationScope op, CancellationToken ct)
    {
        var attempt = 0;
        while (true)
        {
            attempt++;
            HttpResponseMessage response;

            // The try/catch guards ONLY the send. Both a genuine network
            // failure and EnsureSuccessStatusCode() throw HttpRequestException,
            // so if status handling sat inside this block the retry filter
            // could not tell "connection refused" from "401 Unauthorized"
            // and would retry non-transient responses.
            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Get, "rentroll");
                var token = await _tokenProvider.GetTokenAsync(ct).ConfigureAwait(false);
                if (!string.IsNullOrEmpty(token))
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                response = await _http
                    .SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct)
                    .ConfigureAwait(false);
            }
            catch (Exception ex) when (IsTransientNetworkFailure(ex) &&
                                       attempt <= _options.MaxRetries &&
                                       !ct.IsCancellationRequested)
            {
                _logger.LogWarning(ex,
                    "Transient network failure on attempt {Attempt}/{Max}, retrying",
                    attempt, _options.MaxRetries + 1);
                await DelayForRetryAsync(attempt, response: null, ct).ConfigureAwait(false);
                continue;
            }

            using (response)
            {
                if (IsTransient(response.StatusCode) && attempt <= _options.MaxRetries)
                {
                    _logger.LogWarning(
                        "Transient status {Status} on attempt {Attempt}/{Max}, retrying",
                        (int)response.StatusCode, attempt, _options.MaxRetries + 1);
                    await DelayForRetryAsync(attempt, response, ct).ConfigureAwait(false);
                    continue;
                }

                // Non-transient status, or retries exhausted: fail now.
                response.EnsureSuccessStatusCode();

                var records = await response.Content
                    .ReadFromJsonAsync<List<RentRollRecord>>(cancellationToken: ct)
                    .ConfigureAwait(false)
                    ?? throw new InvalidOperationException("Rent roll payload was empty.");

                op.SetProperty("attempts", attempt);
                return records;
            }
        }
    }

    private async Task DelayForRetryAsync(int attempt, HttpResponseMessage? response, CancellationToken ct)
    {
        // Honour Retry-After on 429/503 when the server provides one.
        var serverHint = response?.Headers.RetryAfter?.Delta;

        // Exponential backoff with full jitter: random(0, base * 2^(attempt-1)).
        // Computed in milliseconds because the TimeSpan * double operator
        // does not exist on .NET Framework.
        var capMs = _options.RetryBaseDelay.TotalMilliseconds * Math.Pow(2, attempt - 1);
        double sample;
        lock (JitterLock) sample = Jitter.NextDouble(); // Random.Shared needs net6+; net48-safe
        var jittered = TimeSpan.FromMilliseconds(sample * capMs);

        var delay = serverHint is { } hint && hint > jittered ? hint : jittered;
        _logger.LogDebug("Waiting {DelayMs:F0} ms before retry {Attempt}",
            delay.TotalMilliseconds, attempt + 1);
        await Task.Delay(delay, ct).ConfigureAwait(false);
    }

    private static readonly Random Jitter = new();
    private static readonly object JitterLock = new();

    /// <summary>429. HttpStatusCode.TooManyRequests does not exist on .NET Framework.</summary>
    private const int TooManyRequests = 429;

    private static bool IsTransient(HttpStatusCode status) =>
        (int)status >= 500 ||
        status == HttpStatusCode.RequestTimeout ||
        (int)status == TooManyRequests;

    /// <summary>
    /// True only for failures raised while attempting the send itself
    /// (DNS, connection refused, HttpClient timeout). Deliberately does NOT
    /// classify by exception type alone anywhere status codes are available,
    /// because EnsureSuccessStatusCode throws the same exception type.
    /// User-initiated cancellation is never transient.
    /// </summary>
    private static bool IsTransientNetworkFailure(Exception ex) =>
        ex is HttpRequestException ||
        (ex is TaskCanceledException tce && !tce.CancellationToken.IsCancellationRequested);
}
