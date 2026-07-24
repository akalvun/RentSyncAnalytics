using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Channels;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using RentSync.Core.Configuration;

namespace RentSync.Core.Telemetry;

/// <summary>
/// Buffered, non-blocking telemetry writer.
///
/// Producer side (TrackEvent/TrackError/operation scopes) only enqueues into a
/// bounded channel — O(1), never touches disk, never blocks Excel's UI thread.
/// A single background task drains the channel and appends JSON lines to a
/// local buffer file. If the channel is full the oldest events are dropped
/// (telemetry must never take the host application down with it).
/// </summary>
public sealed class TelemetryService : ITelemetryService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly Channel<TelemetryEvent> _channel;
    private readonly Task _drainTask;
    private readonly CancellationTokenSource _cts = new();
    private readonly ILogger<TelemetryService> _logger;
    private readonly string _bufferFilePath;

    public Guid SessionId { get; } = Guid.NewGuid();

    public TelemetryService(IOptions<TelemetryOptions> options, ILogger<TelemetryService> logger)
    {
        _logger = logger;

        var opts = options.Value;
        Directory.CreateDirectory(opts.BufferDirectory);
        _bufferFilePath = Path.Combine(
            opts.BufferDirectory,
            $"telemetry-{DateTime.UtcNow:yyyyMMdd}.jsonl");

        _channel = Channel.CreateBounded<TelemetryEvent>(new BoundedChannelOptions(opts.BufferCapacity)
        {
            SingleReader = true,
            SingleWriter = false,
            FullMode = BoundedChannelFullMode.DropOldest
        });

        _drainTask = Task.Run(() => DrainAsync(_cts.Token));

        TrackEvent("Telemetry.SessionStarted", new Dictionary<string, object?>
        {
            ["os"] = Environment.OSVersion.VersionString,
            ["runtime"] = Environment.Version.ToString(),
            ["is64Bit"] = Environment.Is64BitProcess
        });
    }

    public void TrackEvent(string name, IReadOnlyDictionary<string, object?>? properties = null) =>
        Enqueue(new TelemetryEvent
        {
            Kind = TelemetryEventKind.Usage,
            Name = name,
            TimestampUtc = DateTimeOffset.UtcNow,
            SessionId = SessionId,
            Properties = properties
        });

    public void TrackError(Exception exception, string context,
        IReadOnlyDictionary<string, object?>? properties = null)
    {
        // Log the full exception locally for diagnostics...
        _logger.LogError(exception, "Handled error in {Context}", context);

        // ...but keep the telemetry record privacy-safe: type + message only, no stack PII.
        Enqueue(new TelemetryEvent
        {
            Kind = TelemetryEventKind.Error,
            Name = context,
            TimestampUtc = DateTimeOffset.UtcNow,
            SessionId = SessionId,
            ExceptionType = exception.GetType().FullName,
            ExceptionMessage = exception.Message,
            Properties = properties
        });
    }

    public IOperationScope StartOperation(string name,
        IReadOnlyDictionary<string, object?>? properties = null) =>
        new OperationScope(this, name, properties);

    public async Task FlushAsync(CancellationToken cancellationToken = default)
    {
        _channel.Writer.TryComplete();

        // Task.WaitAsync is .NET 6+; this is the portable equivalent.
        var cancelled = new TaskCompletionSource<bool>(
            TaskCreationOptions.RunContinuationsAsynchronously);

        using (cancellationToken.Register(() => cancelled.TrySetResult(true)))
        {
            var finished = await Task.WhenAny(_drainTask, cancelled.Task).ConfigureAwait(false);
            if (finished != _drainTask)
                throw new OperationCanceledException(cancellationToken);
        }

        await _drainTask.ConfigureAwait(false); // propagate any drain failure
    }

    public async ValueTask DisposeAsync()
    {
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            await FlushAsync(timeout.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            // Shutdown must not hang the host; drop whatever is left.
        }
        finally
        {
            _cts.Cancel();
            _cts.Dispose();
        }
    }

    private void Enqueue(TelemetryEvent evt)
    {
        // DropOldest mode: TryWrite only fails after Complete(); losing events
        // during shutdown is acceptable for telemetry.
        _channel.Writer.TryWrite(evt);
    }

    private async Task DrainAsync(CancellationToken ct)
    {
        try
        {
            // WaitToReadAsync + TryRead instead of ReadAllAsync: avoids the
            // IAsyncEnumerable dependency, which .NET Framework lacks.
            // AutoFlush is off and we flush once per drained batch rather than
            // per event - one syscall per burst instead of one per record.
            using var writer = new StreamWriter(_bufferFilePath, append: true)
            {
                AutoFlush = false
            };

            while (await _channel.Reader.WaitToReadAsync(ct).ConfigureAwait(false))
            {
                var wrote = false;

                while (_channel.Reader.TryRead(out var evt))
                {
                    try
                    {
                        writer.WriteLine(JsonSerializer.Serialize(evt, JsonOptions));
                        wrote = true;
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        // A failing telemetry sink must never crash the drain loop.
                        _logger.LogWarning(ex,
                            "Failed to persist telemetry event {EventName}", evt.Name);
                    }
                }

                if (wrote) writer.Flush();
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown path.
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Telemetry drain loop stopped unexpectedly");
        }
    }

    /// <summary>Timed, correlated operation. Thread-safe for SetProperty via lock.</summary>
    private sealed class OperationScope : IOperationScope
    {
        private readonly TelemetryService _owner;
        private readonly string _name;
        private readonly Dictionary<string, object?> _properties;
        private readonly Stopwatch _stopwatch = Stopwatch.StartNew();
        private readonly object _gate = new();
        private bool _success;
        private Exception? _exception;
        private bool _disposed;

        public Guid OperationId { get; } = Guid.NewGuid();

        public OperationScope(TelemetryService owner, string name,
            IReadOnlyDictionary<string, object?>? properties)
        {
            _owner = owner;
            _name = name;
            _properties = new Dictionary<string, object?>();

            // Dictionary(IEnumerable<KeyValuePair<,>>) is .NET Core 2.0+;
            // copy manually so the same code compiles on .NET Framework.
            if (properties is not null)
            {
                foreach (var kvp in properties)
                    _properties[kvp.Key] = kvp.Value;
            }
        }

        public void Succeed()
        {
            lock (_gate) _success = true;
        }

        public void Fail(Exception? exception = null)
        {
            lock (_gate)
            {
                _success = false;
                _exception = exception;
            }
        }

        public void SetProperty(string key, object? value)
        {
            lock (_gate) _properties[key] = value;
        }

        public void Dispose()
        {
            lock (_gate)
            {
                if (_disposed) return;
                _disposed = true;
            }

            _stopwatch.Stop();
            _owner.Enqueue(new TelemetryEvent
            {
                Kind = TelemetryEventKind.Operation,
                Name = _name,
                TimestampUtc = DateTimeOffset.UtcNow,
                SessionId = _owner.SessionId,
                OperationId = OperationId,
                DurationMs = _stopwatch.Elapsed.TotalMilliseconds,
                Success = _success,
                ExceptionType = _exception?.GetType().FullName,
                ExceptionMessage = _exception?.Message,
                Properties = _properties.Count > 0 ? _properties : null
            });
        }
    }
}
