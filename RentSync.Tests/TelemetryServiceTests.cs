using System.Text.Json;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using RentSync.Core.Configuration;
using RentSync.Core.Telemetry;
using Xunit;

namespace RentSync.Tests;

public class TelemetryServiceTests
{
    [Fact]
    public async Task Events_AreWrittenAsJsonLines_AndFlushedOnDispose()
    {
        var dir = Path.Combine(Path.GetTempPath(), "rentsync-tests", Guid.NewGuid().ToString("N"));
        var options = Options.Create(new TelemetryOptions { BufferDirectory = dir });

        await using (var telemetry = new TelemetryService(options,
                         NullLogger<TelemetryService>.Instance))
        {
            telemetry.TrackEvent("Test.Click", new Dictionary<string, object?> { ["button"] = "Refresh" });

            using (var op = telemetry.StartOperation("Test.Operation"))
            {
                await Task.Delay(10);
                op.SetProperty("rows", 42);
                op.Succeed();
            }

            telemetry.TrackError(new InvalidOperationException("boom"), "Test.Error");
        } // DisposeAsync flushes

        var file = Directory.GetFiles(dir, "telemetry-*.jsonl").Single();
        var lines = await File.ReadAllLinesAsync(file);

        // SessionStarted + Click + Operation + Error = 4
        Assert.Equal(4, lines.Length);

        var events = lines.Select(l => JsonDocument.Parse(l).RootElement).ToList();

        var op2 = events.Single(e => e.GetProperty("name").GetString() == "Test.Operation");
        Assert.True(op2.GetProperty("durationMs").GetDouble() >= 5);
        Assert.True(op2.GetProperty("success").GetBoolean());

        var error = events.Single(e => e.GetProperty("name").GetString() == "Test.Error");
        Assert.Equal("System.InvalidOperationException",
            error.GetProperty("exceptionType").GetString());

        // Every event carries the same session id (correlation works).
        var sessionIds = events.Select(e => e.GetProperty("sessionId").GetString()).Distinct();
        Assert.Single(sessionIds);
    }
}
