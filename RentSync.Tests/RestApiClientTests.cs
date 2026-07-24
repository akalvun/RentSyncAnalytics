using System.Net;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using RentSync.Core.Configuration;
using RentSync.Core.Services;
using RentSync.Core.Telemetry;
using Xunit;

namespace RentSync.Tests;

public class RestApiClientTests
{
    private const string SampleJson = """
        [{
            "propertyId": "P1", "propertyName": "Riverside Plaza",
            "unitId": "U1", "tenantName": "Acme GmbH",
            "monthlyRent": 2500.00, "areaSqm": 125.0,
            "leaseStart": "2024-01-01", "leaseEnd": "2027-12-31"
        }]
        """;

    private static (RestApiClient client, ScriptedHttpHandler handler, FakeTelemetryService telemetry)
        Build(ScriptedHttpHandler handler)
    {
        var telemetry = new FakeTelemetryService();
        var options = Options.Create(new ApiOptions
        {
            BaseUrl = "https://example.test/",
            MaxRetries = 3,
            RetryBaseDelay = TimeSpan.FromMilliseconds(1) // keep tests fast
        });

        var client = new RestApiClient(
            new HttpClient(handler),
            new AnonymousTokenProvider(),
            telemetry,
            options,
            NullLogger<RestApiClient>.Instance);

        return (client, handler, telemetry);
    }

    [Fact]
    public async Task FetchRentRoll_Success_ParsesRecords()
    {
        var (client, _, _) = Build(new ScriptedHttpHandler()
            .Enqueue(HttpStatusCode.OK, SampleJson));

        var records = await client.FetchRentRollAsync();

        var record = Assert.Single(records);
        Assert.Equal("Riverside Plaza", record.PropertyName);
        Assert.Equal(20.00m, record.RentPerSqm);
    }

    [Fact]
    public async Task FetchRentRoll_TransientErrors_RetriesUntilSuccess()
    {
        var (client, handler, _) = Build(new ScriptedHttpHandler()
            .Enqueue(HttpStatusCode.ServiceUnavailable)
            .Enqueue(HttpStatusCode.TooManyRequests)
            .Enqueue(HttpStatusCode.OK, SampleJson));

        var records = await client.FetchRentRollAsync();

        Assert.Single(records);
        Assert.Equal(3, handler.CallCount);
    }

    [Fact]
    public async Task FetchRentRoll_NonTransientError_FailsFastWithoutRetry()
    {
        var (client, handler, telemetry) = Build(new ScriptedHttpHandler()
            .Enqueue(HttpStatusCode.Unauthorized));

        await Assert.ThrowsAsync<HttpRequestException>(() => client.FetchRentRollAsync());

        Assert.Equal(1, handler.CallCount); // 401 must NOT be retried
        Assert.Contains(telemetry.Events,
            e => e.Kind == TelemetryEventKind.Error && e.Name == "Api.FetchRentRoll");
    }

    [Fact]
    public async Task FetchRentRoll_ExhaustedRetries_Throws()
    {
        var handler = new ScriptedHttpHandler();
        for (var i = 0; i < 5; i++) handler.Enqueue(HttpStatusCode.InternalServerError);
        var (client, _, _) = Build(handler);

        await Assert.ThrowsAsync<HttpRequestException>(() => client.FetchRentRollAsync());
        Assert.Equal(4, handler.CallCount); // 1 initial + 3 retries
    }

    [Fact]
    public async Task FetchRentRoll_RecordsOperationTelemetryWithAttemptCount()
    {
        var (client, _, telemetry) = Build(new ScriptedHttpHandler()
            .Enqueue(HttpStatusCode.ServiceUnavailable)
            .Enqueue(HttpStatusCode.OK, SampleJson));

        await client.FetchRentRollAsync();

        var op = Assert.Single(telemetry.Events,
            e => e.Kind == TelemetryEventKind.Operation && e.Name == "Api.FetchRentRoll");
        Assert.True(op.Success);
        Assert.Equal(2, op.Properties!["attempts"]);
        Assert.Equal(1, op.Properties!["recordCount"]);
    }
}
