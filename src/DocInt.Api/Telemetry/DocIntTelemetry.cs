using System.Diagnostics;
using System.Diagnostics.Metrics;

namespace DocInt.Api.Telemetry;

public sealed class DocIntTelemetry : IDisposable
{
    public const string SourceName = "EuGo.DocInt";
    public const string MeterName = "EuGo.DocInt";
    public const string PagesProcessedInstrument = "docint.pages_processed";

    private readonly Meter _meter;

    public DocIntTelemetry(IMeterFactory meterFactory)
    {
        _meter = meterFactory.Create(MeterName);
        PagesProcessed = _meter.CreateCounter<long>(PagesProcessedInstrument, unit: "pages",
            description: "Pages (DI kinds), sheets (xlsx) or images processed per file");
    }

    public ActivitySource ActivitySource { get; } = new(SourceName);

    public Counter<long> PagesProcessed { get; }

    public void Dispose()
    {
        ActivitySource.Dispose();
        _meter.Dispose();
    }
}
