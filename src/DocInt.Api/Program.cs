using DocInt.Api.Configuration;
using Serilog;
using Serilog.Debugging;

Log.Logger = new LoggerConfiguration()
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .CreateBootstrapLogger();
SelfLog.Enable(Console.Error);

try
{
    Log.Information("DocInt host starting");
    var builder = WebApplication.CreateBuilder(args);

    builder.AddServiceDefaults();
    builder.AddDocIntOptions();

    // Serilog replaces the default logging providers; ServiceDefaults' OTel traces/metrics
    // stay active. Under Aspire (OTEL_EXPORTER_OTLP_ENDPOINT set) also ship logs via OTLP.
    var loggerConfiguration = new LoggerConfiguration()
        .ReadFrom.Configuration(builder.Configuration)
        .Enrich.FromLogContext();
    if (!string.IsNullOrWhiteSpace(builder.Configuration["OTEL_EXPORTER_OTLP_ENDPOINT"]))
    {
        loggerConfiguration.WriteTo.OpenTelemetry();
    }
    var logger = loggerConfiguration.CreateLogger();
    builder.Logging.ClearProviders();
    builder.Logging.AddSerilog(logger);

    builder.Services.AddOpenApi();

    var app = builder.Build();

    app.MapDefaultEndpoints();
    if (app.Environment.IsDevelopment())
    {
        app.MapOpenApi();
    }

    app.MapHealthChecks("/healthz");

    var version = typeof(Program).Assembly
        .GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
        .OfType<System.Reflection.AssemblyInformationalVersionAttribute>()
        .FirstOrDefault()?.InformationalVersion ?? "unknown";
    app.MapGet("/info", () => Results.Json(new
    {
        service = "EuGo-docint",
        version,
        endpoints = new[] { "/v1/extract", "/healthz", "/info" }
    }));

    app.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "DocInt host terminated unexpectedly");
}
finally
{
    Log.CloseAndFlush();
}

public partial class Program;
