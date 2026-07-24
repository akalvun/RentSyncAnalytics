using System.Net;
using RentSync.Core.Services;
using RentSync.Core.Telemetry;

namespace RentSync.Tests;

/// <summary>Records telemetry in memory so tests can assert on it.</summary>
public sealed class FakeTelemetryService : ITelemetryService
{
    public List<TelemetryEvent> Events { get; } = new();
    public Guid SessionId { get; } = Guid.NewGuid();

    public void TrackEvent(string name, IReadOnlyDictionary<string, object?>? properties = null) =>
        Events.Add(new TelemetryEvent
        {
            Kind = TelemetryEventKind.Usage, Name = name,
            TimestampUtc = DateTimeOffset.UtcNow, SessionId = SessionId, Properties = properties
        });

    public void TrackError(Exception exception, string context,
        IReadOnlyDictionary<string, object?>? properties = null) =>
        Events.Add(new TelemetryEvent
        {
            Kind = TelemetryEventKind.Error, Name = context,
            TimestampUtc = DateTimeOffset.UtcNow, SessionId = SessionId,
            ExceptionType = exception.GetType().FullName,
            ExceptionMessage = exception.Message
        });

    public IOperationScope StartOperation(string name,
        IReadOnlyDictionary<string, object?>? properties = null) => new Scope(this, name);

    public Task FlushAsync(CancellationToken cancellationToken = default) => Task.CompletedTask;
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;

    private sealed class Scope : IOperationScope
    {
        private readonly FakeTelemetryService _owner;
        private readonly string _name;
        private readonly Dictionary<string, object?> _props = new();
        private bool _success;

        public Guid OperationId { get; } = Guid.NewGuid();
        public Scope(FakeTelemetryService owner, string name) { _owner = owner; _name = name; }
        public void Succeed() => _success = true;
        public void Fail(Exception? exception = null) => _success = false;
        public void SetProperty(string key, object? value) => _props[key] = value;

        public void Dispose() => _owner.Events.Add(new TelemetryEvent
        {
            Kind = TelemetryEventKind.Operation, Name = _name,
            TimestampUtc = DateTimeOffset.UtcNow, SessionId = _owner.SessionId,
            OperationId = OperationId, Success = _success,
            Properties = _props.Count > 0 ? _props : null
        });
    }
}

/// <summary>Scripted HttpMessageHandler: returns queued responses in order.</summary>
public sealed class ScriptedHttpHandler : HttpMessageHandler
{
    private readonly Queue<Func<HttpResponseMessage>> _responses = new();
    public int CallCount { get; private set; }

    public ScriptedHttpHandler Enqueue(HttpStatusCode status, string? json = null)
    {
        _responses.Enqueue(() =>
        {
            var response = new HttpResponseMessage(status);
            if (json is not null)
                response.Content = new StringContent(json,
                    System.Text.Encoding.UTF8, "application/json");
            return response;
        });
        return this;
    }

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken cancellationToken)
    {
        CallCount++;
        if (_responses.Count == 0)
            throw new InvalidOperationException("No scripted response left.");
        return Task.FromResult(_responses.Dequeue()());
    }
}
