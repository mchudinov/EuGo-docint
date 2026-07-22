using System.Diagnostics.Metrics;
using DocInt.Api.Telemetry;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Diagnostics.Metrics.Testing;

namespace DocInt.Tests;

public class TelemetryTests : IClassFixture<ContractTestFactory>
{
    private readonly ContractTestFactory _factory;

    public TelemetryTests(ContractTestFactory factory) => _factory = factory;

    [Fact]
    public async Task Pages_processed_metric_counts_sheets_with_kind_tag()
    {
        var meterFactory = _factory.Services.GetRequiredService<IMeterFactory>();
        using var collector = new MetricCollector<long>(
            meterFactory, DocIntTelemetry.MeterName, DocIntTelemetry.PagesProcessedInstrument);

        using var form = Multipart.Form(("bom.xlsx", Golden.Bytes("bom.xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        var response = await _factory.CreateClient().PostAsync("/v1/extract", form);
        response.EnsureSuccessStatusCode();

        var measurements = collector.GetMeasurementSnapshot();
        Assert.Equal(2, measurements.Sum(m => m.Value));   // BoM + Notes sheets
        Assert.All(measurements, m => Assert.Equal("xlsx", m.Tags["kind"]));
    }
}
