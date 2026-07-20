# EuGo-docint Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the stateless EuGo-docint document-understanding service (T1–T6 of the approved design): `POST /v1/extract` → Markdown / typed cells / image descriptions / per-file errors, with Aspire, golden-file tests, env-gated live smoke, and multi-arch container CI.

**Architecture:** Single ASP.NET Core minimal-API project (`DocInt.Api`) behind an Aspire AppHost + ServiceDefaults, engines behind the `IExtractionEngine` seam kind-keyed by `EngineRouter`, Azure reached only through two adapter interfaces (`ILayoutAnalysisClient`, `IVisionChatClient`) that are faked in tests. Spec: `docs/superpowers/specs/2026-07-19-eugo-docint-design.md` (authoritative; read it first).

**Tech Stack:** .NET 10 (SDK 10.0.300) · Aspire.AppHost.Sdk/13.1.0 · Serilog · Azure.AI.DocumentIntelligence 1.0.0 · Azure.AI.OpenAI 2.3.0 · DocumentFormat.OpenXml 3.3.0 · xUnit 2.9.3 + WebApplicationFactory · SkiaSharp 3.119.0 (fixture tool only) · Docker buildx multi-arch.

## Global Constraints

- Gate before every merge, exactly: `dotnet restore src/DocInt.slnx` → `dotnet build --no-restore src/DocInt.slnx` → `dotnet test --no-build src/DocInt.slnx`. Never merge red.
- Workflow per task: `git checkout main && git pull`, branch `step-NN-<slug>`, TDD steps below, gate green, `git checkout main && git merge step-NN-<slug> && git branch -d step-NN-<slug> && git push origin main`.
- All projects: `net10.0`, `<Nullable>enable</Nullable>`, `<ImplicitUsings>enable</ImplicitUsings>`.
- Document-level contract only — no compliance semantics anywhere. Stateless — nothing written to disk or any store by the service. **No document content or extraction output in logs** (filenames/sizes/kinds/durations are fine).
- Per-file error codes (exact strings): `unsupported_type`, `too_large`, `empty_file`, `corrupt`, `timeout`, `engine_error`, `engine_unconfigured`. Request-level 400 only for: not multipart, zero files, > MaxFilesPerRequest, malformed hints JSON, request body over MaxRequestBytes.
- Purpose hints v1: `bom`, `photo` — advisory only; unknown value → per-file warning `unknown purpose hint '<x>' ignored`.
- Config keys: `DocumentIntelligence:Endpoint|ApiKey`, `AzureOpenAI:Endpoint|ApiKey|DeploymentNameVision` (default `gpt-4.1-mini`), `DocInt:MaxFileBytes` (52428800), `DocInt:MaxFilesPerRequest` (32), `DocInt:PerFileTimeoutSeconds` (100), `DocInt:MaxParallelism` (4). Endpoint+ApiKey → key credential; endpoint only → `DefaultAzureCredential`. Never secrets in source.
- Package versions are pinned in this plan. If `dotnet restore` reports a pinned version does not exist, resolve the latest stable of the same major via Context7/nuget.org, use it, and record the substitution in the commit message.
- HTTP port is **8090** everywhere (8089 belongs to sibling repos).
- JSON wire format: camelCase properties, enums as lowercase strings, nulls omitted (`DocIntJson.Options` is the single source of truth; always serialize responses with it).

## File Map (final state)

```text
src/DocInt.slnx
src/DocInt.Api/DocInt.Api.csproj · Program.cs · appsettings{,.Development}.json · Properties/launchSettings.json
src/DocInt.Api/Api/ExtractEndpoint.cs
src/DocInt.Api/Configuration/DocIntOptions.cs
src/DocInt.Api/Contracts/ExtractContracts.cs
src/DocInt.Api/Validation/FileKindDetector.cs · HintsParser.cs · MultipartExtractRequestReader.cs
src/DocInt.Api/Engines/FileItem.cs · IExtractionEngine.cs · Errors.cs · EngineRouter.cs · ExtractionService.cs
src/DocInt.Api/Engines/SpreadsheetEngine.cs · LayoutEngine.cs · VisionEngine.cs · VisionPrompt.cs
src/DocInt.Api/Engines/ILayoutAnalysisClient.cs · AzureLayoutAnalysisClient.cs · IVisionChatClient.cs · AzureVisionChatClient.cs
src/DocInt.Api/Telemetry/DocIntTelemetry.cs
src/AppHost/AppHost.csproj · Program.cs · Properties/launchSettings.json
src/ServiceDefaults/ServiceDefaults.csproj · Extensions.cs
tests/DocInt.Tests/DocInt.Tests.csproj · DocIntAppFactory.cs · ContractTestFactory.cs · TestSupport.cs
tests/DocInt.Tests/HealthEndpointsTests.cs · ContractsSerializationTests.cs · OptionsTests.cs
tests/DocInt.Tests/FileKindDetectorTests.cs · HintsParserTests.cs · MultipartReaderTests.cs
tests/DocInt.Tests/ExtractContractTests.cs · ExtractionServiceTests.cs · GoldenFixturesTests.cs
tests/DocInt.Tests/SpreadsheetEngineTests.cs · LayoutEngineTests.cs · VisionEngineTests.cs
tests/DocInt.Tests/TelemetryTests.cs · RedactionTests.cs · LiveSmokeTests.cs
tests/DocInt.Tests/golden/  (committed binaries: text.pdf scanned.pdf sample.docx sample.pptx sample.html bom.xlsx photo.png corrupt.xlsx unknown.bin)
tools/make-golden/make-golden.csproj · Program.cs · MiniPdf.cs · OfficeFixtures.cs · ImageFixtures.cs
Dockerfile · .github/workflows/ci.yml · README.md
```

---

### Task 1: Solution scaffold — ServiceDefaults, AppHost, DocInt.Api with `/healthz` + `/info`, tests, CI build-test job

**Files:**
- Create: `src/DocInt.slnx`, `src/ServiceDefaults/ServiceDefaults.csproj`, `src/ServiceDefaults/Extensions.cs`, `src/AppHost/AppHost.csproj`, `src/AppHost/Program.cs`, `src/AppHost/Properties/launchSettings.json`, `src/DocInt.Api/DocInt.Api.csproj`, `src/DocInt.Api/Program.cs`, `src/DocInt.Api/appsettings.json`, `src/DocInt.Api/appsettings.Development.json`, `src/DocInt.Api/Properties/launchSettings.json`, `tests/DocInt.Tests/DocInt.Tests.csproj`, `tests/DocInt.Tests/DocIntAppFactory.cs`, `tests/DocInt.Tests/HealthEndpointsTests.cs`, `.github/workflows/ci.yml`

**Interfaces:**
- Produces: `public partial class Program` (WebApplicationFactory anchor); `DocIntAppFactory : WebApplicationFactory<Program>` with `protected virtual void ConfigureFakes(IServiceCollection)`; endpoints `GET /healthz` (200 "Healthy"), `GET /info` (JSON `{service, version, endpoints[]}`).

- [ ] **Step 1: Branch**

```bash
cd /home/michael/Documents/projects/real/EuGo-docint
git checkout main && git pull
git checkout -b step-01-scaffold
```

- [ ] **Step 2: Create solution and project files**

```bash
dotnet new sln --format slnx -n DocInt -o src
mkdir -p src/ServiceDefaults src/AppHost/Properties src/DocInt.Api/Properties tests/DocInt.Tests
```

Create `src/ServiceDefaults/ServiceDefaults.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsAspireSharedProject>true</IsAspireSharedProject>
  </PropertyGroup>

  <ItemGroup>
    <FrameworkReference Include="Microsoft.AspNetCore.App" />

    <PackageReference Include="Microsoft.Extensions.Http.Resilience" Version="10.7.0" />
    <PackageReference Include="Microsoft.Extensions.ServiceDiscovery" Version="10.7.0" />
    <PackageReference Include="OpenTelemetry.Exporter.OpenTelemetryProtocol" Version="1.16.0" />
    <PackageReference Include="OpenTelemetry.Extensions.Hosting" Version="1.16.0" />
    <PackageReference Include="OpenTelemetry.Instrumentation.AspNetCore" Version="1.16.0" />
    <PackageReference Include="OpenTelemetry.Instrumentation.Http" Version="1.16.0" />
    <PackageReference Include="OpenTelemetry.Instrumentation.Runtime" Version="1.15.1" />
  </ItemGroup>

</Project>
```

Create `src/ServiceDefaults/Extensions.cs` — copy the file **verbatim** from `/home/michael/Documents/projects/real/EuGo-mcp/src/ServiceDefaults/Extensions.cs` (stock Aspire service defaults: `AddServiceDefaults`, `ConfigureOpenTelemetry`, `AddDefaultHealthChecks`, `MapDefaultEndpoints` with `/health` + `/alive` in Development only):

```bash
cp /home/michael/Documents/projects/real/EuGo-mcp/src/ServiceDefaults/Extensions.cs src/ServiceDefaults/Extensions.cs
```

Create `src/AppHost/AppHost.csproj`:

```xml
<Project Sdk="Aspire.AppHost.Sdk/13.1.0">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <UserSecretsId>9f4a2f10-3e7d-4f0a-9c68-2f60d0c1e702</UserSecretsId>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="..\DocInt.Api\DocInt.Api.csproj" />
  </ItemGroup>

</Project>
```

Create `src/AppHost/Program.cs`:

```csharp
var builder = DistributedApplication.CreateBuilder(args);

builder.AddProject<Projects.DocInt_Api>("docint");

builder.Build().Run();
```

Create `src/AppHost/Properties/launchSettings.json`:

```json
{
  "$schema": "https://json.schemastore.org/launchsettings.json",
  "profiles": {
    "http": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "applicationUrl": "http://localhost:15290",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development",
        "DOTNET_ENVIRONMENT": "Development",
        "ASPIRE_DASHBOARD_OTLP_ENDPOINT_URL": "http://localhost:19290",
        "ASPIRE_DASHBOARD_MCP_ENDPOINT_URL": "http://localhost:18290",
        "ASPIRE_RESOURCE_SERVICE_ENDPOINT_URL": "http://localhost:20290",
        "ASPIRE_ALLOW_UNSECURED_TRANSPORT": "true"
      }
    }
  }
}
```

Create `src/DocInt.Api/DocInt.Api.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk.Web">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <UserSecretsId>9f4a2f10-3e7d-4f0a-9c68-2f60d0c1e701</UserSecretsId>
    <DockerDefaultTargetOS>Linux</DockerDefaultTargetOS>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.AspNetCore.OpenApi" Version="10.0.9" />
    <PackageReference Include="Serilog" Version="4.3.1" />
    <PackageReference Include="Serilog.Extensions.Hosting" Version="10.0.0" />
    <PackageReference Include="Serilog.Settings.Configuration" Version="10.0.1" />
    <PackageReference Include="Serilog.Sinks.Console" Version="6.1.1" />
    <PackageReference Include="Serilog.Sinks.Debug" Version="3.0.0" />
    <PackageReference Include="Serilog.Sinks.OpenTelemetry" Version="4.2.0" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\ServiceDefaults\ServiceDefaults.csproj" />
  </ItemGroup>
</Project>
```

Create `src/DocInt.Api/appsettings.json`:

```json
{
  "Serilog": {
    "MinimumLevel": {
      "Default": "Information",
      "Override": {
        "Microsoft": "Error",
        "System": "Error"
      }
    },
    "WriteTo": [
      { "Name": "Console" },
      { "Name": "Debug" }
    ]
  },
  "Kestrel": {
    "EndPoints": {
      "Http": {
        "Url": "http://*:8090"
      }
    }
  },
  "AllowedHosts": "*"
}
```

Create `src/DocInt.Api/appsettings.Development.json`:

```json
{
  "Serilog": {
    "MinimumLevel": {
      "Default": "Debug"
    }
  }
}
```

Create `src/DocInt.Api/Properties/launchSettings.json`:

```json
{
  "$schema": "https://json.schemastore.org/launchsettings.json",
  "profiles": {
    "http": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": false,
      "applicationUrl": "http://localhost:8090",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      }
    }
  }
}
```

Create `src/DocInt.Api/Program.cs` — **deliberately without** `/healthz` and `/info` yet (the failing test comes first):

```csharp
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
```

Create `tests/DocInt.Tests/DocInt.Tests.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="18.7.0" />
    <PackageReference Include="Microsoft.AspNetCore.Mvc.Testing" Version="10.0.9" />
    <PackageReference Include="xunit" Version="2.9.3" />
    <PackageReference Include="xunit.runner.visualstudio" Version="3.1.5">
      <PrivateAssets>all</PrivateAssets>
      <IncludeAssets>runtime; build; native; contentfiles; analyzers; buildtransitive</IncludeAssets>
    </PackageReference>
  </ItemGroup>
  <ItemGroup>
    <Using Include="Xunit" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\..\src\DocInt.Api\DocInt.Api.csproj" />
  </ItemGroup>
</Project>
```

Create `tests/DocInt.Tests/DocIntAppFactory.cs`:

```csharp
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.DependencyInjection;

namespace DocInt.Tests;

/// <summary>
/// Hosts the real Program in-memory. Subclasses override ConfigureFakes to replace
/// Azure adapters / engines; the base factory runs the app exactly as configured,
/// which means no Azure endpoints are set (the engine_unconfigured path).
/// </summary>
public class DocIntAppFactory : WebApplicationFactory<Program>
{
    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        builder.ConfigureTestServices(ConfigureFakes);
    }

    protected virtual void ConfigureFakes(IServiceCollection services)
    {
    }
}
```

Add projects to the solution:

```bash
dotnet sln src/DocInt.slnx add src/DocInt.Api/DocInt.Api.csproj src/ServiceDefaults/ServiceDefaults.csproj src/AppHost/AppHost.csproj tests/DocInt.Tests/DocInt.Tests.csproj
```

- [ ] **Step 3: Write the failing test**

Create `tests/DocInt.Tests/HealthEndpointsTests.cs`:

```csharp
using System.Net;
using System.Text.Json;

namespace DocInt.Tests;

public class HealthEndpointsTests : IClassFixture<DocIntAppFactory>
{
    private readonly HttpClient _client;

    public HealthEndpointsTests(DocIntAppFactory factory) => _client = factory.CreateClient();

    [Fact]
    public async Task Healthz_returns_healthy()
    {
        var response = await _client.GetAsync("/healthz");
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal("Healthy", await response.Content.ReadAsStringAsync());
    }

    [Fact]
    public async Task Info_returns_service_metadata()
    {
        var response = await _client.GetAsync("/info");
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        Assert.Equal("EuGo-docint", doc.RootElement.GetProperty("service").GetString());
        Assert.False(string.IsNullOrWhiteSpace(doc.RootElement.GetProperty("version").GetString()));
        var endpoints = doc.RootElement.GetProperty("endpoints").EnumerateArray().Select(e => e.GetString()).ToArray();
        Assert.Contains("/healthz", endpoints);
        Assert.Contains("/info", endpoints);
    }
}
```

- [ ] **Step 4: Run test to verify it fails**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: build succeeds; both tests FAIL with 404 NotFound (endpoints not mapped).

- [ ] **Step 5: Map the endpoints**

In `src/DocInt.Api/Program.cs`, replace:

```csharp
    var app = builder.Build();

    app.MapDefaultEndpoints();
    if (app.Environment.IsDevelopment())
    {
        app.MapOpenApi();
    }

    app.Run();
```

with:

```csharp
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
```

- [ ] **Step 6: Run tests to verify they pass**

```bash
dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (2 tests).

- [ ] **Step 7: CI workflow (build-test job)**

Create `.github/workflows/ci.yml`:

```yaml
name: ci

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  build-test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: 10.0.x
      - run: dotnet restore src/DocInt.slnx
      - run: dotnet build --no-restore src/DocInt.slnx
      - run: dotnet test --no-build src/DocInt.slnx
```

- [ ] **Step 8: Smoke-run both entry points (manual check, no commit content)**

```bash
timeout 15 dotnet run --project src/DocInt.Api &
sleep 8 && curl -s http://localhost:8090/healthz && curl -s http://localhost:8090/info; wait
```

Expected: `Healthy` and the `/info` JSON. (AppHost run `dotnet run --project src/AppHost` is interactive — verify once manually if desired, not required for the gate.)

- [ ] **Step 9: Gate, commit, merge**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
git add -A && git commit -m "Step 1: Aspire solution scaffold with /healthz, /info and CI build-test job"
git checkout main && git merge step-01-scaffold && git branch -d step-01-scaffold && git push origin main
```

---

### Task 2: Wire contracts + configuration options

**Files:**
- Create: `src/DocInt.Api/Contracts/ExtractContracts.cs`, `src/DocInt.Api/Configuration/DocIntOptions.cs`
- Modify: `src/DocInt.Api/Program.cs`
- Test: `tests/DocInt.Tests/ContractsSerializationTests.cs`, `tests/DocInt.Tests/OptionsTests.cs`

**Interfaces:**
- Produces (wire types, used by every later task): `FileKind { Pdf, Docx, Pptx, Html, Xlsx, Image }` + `Name()` extension → `"pdf"`…; `FileError(string Code, string Message)`; `TableResult(string Name, string Markdown, IReadOnlyList<IReadOnlyList<object?>> Rows)`; `FileResult(string Name, FileKind? Kind, string? Markdown, IReadOnlyList<TableResult>? Tables, string? ImageDescription, IReadOnlyList<string> Warnings, FileError? Error)`; `ExtractResponse(IReadOnlyList<FileResult> Files)`; `ErrorCodes` + `PurposeHints` constants; `DocIntJson.Options`.
- Produces (options): `DocIntOptions { MaxFileBytes=52428800, MaxFilesPerRequest=32, PerFileTimeoutSeconds=100, MaxParallelism=4, MaxRequestBytes (computed) }`; `DocumentIntelligenceOptions { Endpoint?, ApiKey? }`; `AzureOpenAIOptions { Endpoint?, ApiKey?, DeploymentNameVision="gpt-4.1-mini" }`; `builder.AddDocIntOptions()`.

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-02-contracts
```

- [ ] **Step 2: Write the failing tests**

Create `tests/DocInt.Tests/ContractsSerializationTests.cs`:

```csharp
using System.Text.Json;
using DocInt.Api.Contracts;

namespace DocInt.Tests;

public class ContractsSerializationTests
{
    [Fact]
    public void Kind_serializes_lowercase_and_nulls_are_omitted()
    {
        var result = new FileResult("manual.pdf", FileKind.Pdf, "# md", null, null, ["w1"], null);
        var json = JsonSerializer.Serialize(new ExtractResponse([result]), DocIntJson.Options);
        Assert.Contains("\"kind\":\"pdf\"", json);
        Assert.Contains("\"warnings\":[\"w1\"]", json);
        Assert.DoesNotContain("tables", json);
        Assert.DoesNotContain("imageDescription", json);
        Assert.DoesNotContain("error", json);
    }

    [Fact]
    public void Typed_cells_serialize_as_native_json_scalars()
    {
        var table = new TableResult("Sheet1", "| a |",
            [["Part", 40m, 19.99m, true, null]]);
        var result = new FileResult("bom.xlsx", FileKind.Xlsx, "# md", [table], null, [], null);
        var json = JsonSerializer.Serialize(result, DocIntJson.Options);
        Assert.Contains("[\"Part\",40,19.99,true,null]", json);
    }

    [Fact]
    public void Error_result_serializes_code_and_message()
    {
        var result = new FileResult("x.bin", null, null, null, null, [],
            new FileError(ErrorCodes.UnsupportedType, "could not detect a supported file type for 'x.bin'"));
        var json = JsonSerializer.Serialize(result, DocIntJson.Options);
        Assert.Contains("\"code\":\"unsupported_type\"", json);
        Assert.DoesNotContain("\"kind\"", json);
    }

    [Fact]
    public void Kind_names_are_lowercase()
    {
        Assert.Equal("pdf", FileKind.Pdf.Name());
        Assert.Equal("xlsx", FileKind.Xlsx.Name());
        Assert.Equal("image", FileKind.Image.Name());
    }
}
```

Create `tests/DocInt.Tests/OptionsTests.cs`:

```csharp
using DocInt.Api.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace DocInt.Tests;

public class OptionsTests
{
    [Fact]
    public void Defaults_match_the_spec()
    {
        var o = new DocIntOptions();
        Assert.Equal(52_428_800, o.MaxFileBytes);
        Assert.Equal(32, o.MaxFilesPerRequest);
        Assert.Equal(100, o.PerFileTimeoutSeconds);
        Assert.Equal(4, o.MaxParallelism);
        Assert.Equal(52_428_800L * 32 + 1_048_576, o.MaxRequestBytes);
        Assert.Equal("gpt-4.1-mini", new AzureOpenAIOptions().DeploymentNameVision);
    }

    [Fact]
    public void Configuration_binds_into_options()
    {
        using var factory = new BindingFactory();
        var docint = factory.Services.GetRequiredService<IOptions<DocIntOptions>>().Value;
        var di = factory.Services.GetRequiredService<IOptions<DocumentIntelligenceOptions>>().Value;
        Assert.Equal(3, docint.MaxFilesPerRequest);
        Assert.Equal("https://di.example", di.Endpoint);
    }

    private sealed class BindingFactory : DocIntAppFactory
    {
        protected override void ConfigureWebHost(Microsoft.AspNetCore.Hosting.IWebHostBuilder builder)
        {
            builder.UseSetting("DocInt:MaxFilesPerRequest", "3");
            builder.UseSetting("DocumentIntelligence:Endpoint", "https://di.example");
            base.ConfigureWebHost(builder);
        }
    }
}
```

- [ ] **Step 3: Run tests to verify they fail**

```bash
dotnet build --no-restore src/DocInt.slnx
```

Expected: FAILS to compile — `DocInt.Api.Contracts` / `DocInt.Api.Configuration` do not exist.

- [ ] **Step 4: Implement contracts**

Create `src/DocInt.Api/Contracts/ExtractContracts.cs`:

```csharp
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DocInt.Api.Contracts;

public enum FileKind
{
    Pdf,
    Docx,
    Pptx,
    Html,
    Xlsx,
    Image
}

public static class FileKindExtensions
{
    public static string Name(this FileKind kind) => JsonNamingPolicy.CamelCase.ConvertName(kind.ToString());
}

public sealed record FileError(string Code, string Message);

public sealed record TableResult(string Name, string Markdown, IReadOnlyList<IReadOnlyList<object?>> Rows);

public sealed record FileResult(
    string Name,
    FileKind? Kind,
    string? Markdown,
    IReadOnlyList<TableResult>? Tables,
    string? ImageDescription,
    IReadOnlyList<string> Warnings,
    FileError? Error);

public sealed record ExtractResponse(IReadOnlyList<FileResult> Files);

public static class ErrorCodes
{
    public const string UnsupportedType = "unsupported_type";
    public const string TooLarge = "too_large";
    public const string EmptyFile = "empty_file";
    public const string Corrupt = "corrupt";
    public const string Timeout = "timeout";
    public const string EngineError = "engine_error";
    public const string EngineUnconfigured = "engine_unconfigured";
}

public static class PurposeHints
{
    public const string Bom = "bom";
    public const string Photo = "photo";

    public static readonly IReadOnlySet<string> Known =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { Bom, Photo };
}

/// <summary>The single wire-format definition: camelCase, enums lowercase, nulls omitted.</summary>
public static class DocIntJson
{
    public static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };
}
```

- [ ] **Step 5: Implement options**

Create `src/DocInt.Api/Configuration/DocIntOptions.cs`:

```csharp
namespace DocInt.Api.Configuration;

public sealed class DocIntOptions
{
    public const string SectionName = "DocInt";

    public long MaxFileBytes { get; set; } = 52_428_800;
    public int MaxFilesPerRequest { get; set; } = 32;
    public int PerFileTimeoutSeconds { get; set; } = 100;
    public int MaxParallelism { get; set; } = 4;

    /// <summary>Request-level cap: worst-case accepted payload plus multipart overhead.</summary>
    public long MaxRequestBytes => MaxFileBytes * MaxFilesPerRequest + 1_048_576;
}

public sealed class DocumentIntelligenceOptions
{
    public const string SectionName = "DocumentIntelligence";

    public string? Endpoint { get; set; }
    public string? ApiKey { get; set; }
}

public sealed class AzureOpenAIOptions
{
    public const string SectionName = "AzureOpenAI";

    public string? Endpoint { get; set; }
    public string? ApiKey { get; set; }
    public string DeploymentNameVision { get; set; } = "gpt-4.1-mini";
}

public static class OptionsExtensions
{
    public static WebApplicationBuilder AddDocIntOptions(this WebApplicationBuilder builder)
    {
        builder.Services.AddOptions<DocIntOptions>()
            .Bind(builder.Configuration.GetSection(DocIntOptions.SectionName))
            .Validate(o => o.MaxFileBytes > 0 && o.MaxFilesPerRequest > 0
                        && o.PerFileTimeoutSeconds > 0 && o.MaxParallelism > 0,
                "DocInt options must all be positive")
            .ValidateOnStart();
        builder.Services.Configure<DocumentIntelligenceOptions>(
            builder.Configuration.GetSection(DocumentIntelligenceOptions.SectionName));
        builder.Services.Configure<AzureOpenAIOptions>(
            builder.Configuration.GetSection(AzureOpenAIOptions.SectionName));
        return builder;
    }
}
```

In `src/DocInt.Api/Program.cs`: add `using DocInt.Api.Configuration;` as the first line of the file, and insert after `builder.AddServiceDefaults();`:

```csharp
    builder.AddDocIntOptions();
```

- [ ] **Step 6: Run tests to verify they pass**

```bash
dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (8 tests).

- [ ] **Step 7: Gate, commit, merge**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
git add -A && git commit -m "Step 2: wire contracts (ExtractResponse envelope, DocIntJson) and typed options"
git checkout main && git merge step-02-contracts && git branch -d step-02-contracts && git push origin main
```

---

### Task 3: Kind detection, hints parsing, multipart request reader

**Files:**
- Create: `src/DocInt.Api/Engines/FileItem.cs`, `src/DocInt.Api/Validation/FileKindDetector.cs`, `src/DocInt.Api/Validation/HintsParser.cs`, `src/DocInt.Api/Validation/MultipartExtractRequestReader.cs`
- Test: `tests/DocInt.Tests/FileKindDetectorTests.cs`, `tests/DocInt.Tests/HintsParserTests.cs`, `tests/DocInt.Tests/MultipartReaderTests.cs`, `tests/DocInt.Tests/TestSupport.cs`

**Interfaces:**
- Consumes: Task 2 contracts + options.
- Produces: `FileItem { Index, Name, ClaimedContentType, Bytes, Kind?, ImageMediaType?, Purpose?, Warnings, Error? }`; `DetectionResult(FileKind? Kind, string? ImageMediaType, string? Warning)`; `FileKindDetector.Detect(string fileName, string? contentType, ReadOnlySpan<byte> content)`; `HintsParser.Parse(string json) → Dictionary<string,string>`; `BadExtractRequestException(string detail)`; `MultipartExtractRequestReader.ReadAsync(HttpRequest, CancellationToken) → IReadOnlyList<FileItem>` (throws `BadExtractRequestException` for the 400 cases).

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-03-validation
```

- [ ] **Step 2: Write the failing detector tests**

Create `tests/DocInt.Tests/TestSupport.cs`:

```csharp
using System.Net.Http.Headers;
using System.Text;

namespace DocInt.Tests;

public static class TestBytes
{
    public static readonly byte[] Pdf = Encoding.ASCII.GetBytes("%PDF-1.4\n% tiny fake body\n");
    public static readonly byte[] Png = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D];
    public static readonly byte[] Jpeg = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10];
    public static readonly byte[] Zip = [0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
    public static readonly byte[] Garbage = [0x00, 0x11, 0x22, 0x33, 0x44, 0x55];
    public static readonly byte[] Html = Encoding.ASCII.GetBytes("  <!DOCTYPE html><html><body>x</body></html>");
}

public static class Multipart
{
    public static MultipartFormDataContent Form(params (string Name, byte[] Bytes, string ContentType)[] files)
    {
        var form = new MultipartFormDataContent();
        foreach (var (name, bytes, contentType) in files)
        {
            var part = new ByteArrayContent(bytes);
            part.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);
            form.Add(part, "files", name);
        }
        return form;
    }

    public static MultipartFormDataContent WithHints(this MultipartFormDataContent form, string json)
    {
        form.Add(new StringContent(json, Encoding.UTF8, "application/json"), "hints");
        return form;
    }
}
```

Create `tests/DocInt.Tests/FileKindDetectorTests.cs`:

```csharp
using DocInt.Api.Contracts;
using DocInt.Api.Validation;

namespace DocInt.Tests;

public class FileKindDetectorTests
{
    [Fact]
    public void Pdf_magic_wins_regardless_of_claims()
    {
        var r = FileKindDetector.Detect("doc.bin", "application/octet-stream", TestBytes.Pdf);
        Assert.Equal(FileKind.Pdf, r.Kind);
        Assert.Null(r.ImageMediaType);
    }

    [Fact]
    public void Png_and_jpeg_magic_detect_image_with_media_type()
    {
        Assert.Equal(("image/png", FileKind.Image),
            (FileKindDetector.Detect("a.png", null, TestBytes.Png).ImageMediaType,
             FileKindDetector.Detect("a.png", null, TestBytes.Png).Kind));
        Assert.Equal("image/jpeg", FileKindDetector.Detect("b.jpg", null, TestBytes.Jpeg).ImageMediaType);
    }

    [Fact]
    public void Zip_disambiguates_by_content_type_then_extension()
    {
        Assert.Equal(FileKind.Xlsx, FileKindDetector.Detect("f.bin",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", TestBytes.Zip).Kind);
        Assert.Equal(FileKind.Docx, FileKindDetector.Detect("f.docx", null, TestBytes.Zip).Kind);
        Assert.Equal(FileKind.Pptx, FileKindDetector.Detect("f.pptx", "application/octet-stream", TestBytes.Zip).Kind);
        Assert.Null(FileKindDetector.Detect("f.zip", null, TestBytes.Zip).Kind);
    }

    [Fact]
    public void Html_detected_by_content_type_extension_or_sniff()
    {
        Assert.Equal(FileKind.Html, FileKindDetector.Detect("p.bin", "text/html; charset=utf-8", "<p>x</p>"u8).Kind);
        Assert.Equal(FileKind.Html, FileKindDetector.Detect("p.html", null, "<p>x</p>"u8).Kind);
        Assert.Equal(FileKind.Html, FileKindDetector.Detect("p.bin", null, TestBytes.Html).Kind);
    }

    [Fact]
    public void Claimed_type_contradicting_bytes_trusts_bytes_and_warns()
    {
        var r = FileKindDetector.Detect("fake.pdf", "application/pdf", TestBytes.Png);
        Assert.Equal(FileKind.Image, r.Kind);
        Assert.NotNull(r.Warning);
        Assert.Contains("application/pdf", r.Warning);
    }

    [Fact]
    public void Matching_claim_produces_no_warning()
    {
        Assert.Null(FileKindDetector.Detect("a.pdf", "application/pdf", TestBytes.Pdf).Warning);
    }

    [Fact]
    public void Undetectable_returns_null_kind()
    {
        Assert.Null(FileKindDetector.Detect("x.bin", null, TestBytes.Garbage).Kind);
        var claimed = FileKindDetector.Detect("x.pdf", "application/pdf", TestBytes.Garbage);
        Assert.Null(claimed.Kind);
        Assert.NotNull(claimed.Warning);
    }
}
```

- [ ] **Step 3: Run to verify compile failure**

```bash
dotnet build --no-restore src/DocInt.slnx
```

Expected: FAILS — `DocInt.Api.Validation` does not exist.

- [ ] **Step 4: Implement FileItem + detector**

Create `src/DocInt.Api/Engines/FileItem.cs`:

```csharp
using DocInt.Api.Contracts;

namespace DocInt.Api.Engines;

/// <summary>One file of an extract request as it moves through validation → routing → engine.</summary>
public sealed class FileItem
{
    public required int Index { get; init; }
    public required string Name { get; init; }
    public string? ClaimedContentType { get; init; }
    public byte[] Bytes { get; init; } = [];
    public FileKind? Kind { get; set; }
    public string? ImageMediaType { get; set; }
    public string? Purpose { get; set; }
    public List<string> Warnings { get; } = [];
    public FileError? Error { get; set; }
}
```

Create `src/DocInt.Api/Validation/FileKindDetector.cs`:

```csharp
using DocInt.Api.Contracts;

namespace DocInt.Api.Validation;

public sealed record DetectionResult(FileKind? Kind, string? ImageMediaType, string? Warning);

public static class FileKindDetector
{
    public static DetectionResult Detect(string fileName, string? contentType, ReadOnlySpan<byte> content)
    {
        var claimed = ClaimedKind(contentType);
        var ext = Path.GetExtension(fileName).ToLowerInvariant();

        if (content.StartsWith("%PDF"u8))
            return Result(FileKind.Pdf, null, claimed, contentType);
        if (content.StartsWith((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]))
            return Result(FileKind.Image, "image/png", claimed, contentType);
        if (content.Length >= 3 && content[0] == 0xFF && content[1] == 0xD8 && content[2] == 0xFF)
            return Result(FileKind.Image, "image/jpeg", claimed, contentType);

        if (content.StartsWith("PK"u8))
        {
            var kind = claimed is FileKind.Docx or FileKind.Xlsx or FileKind.Pptx
                ? claimed
                : ext switch
                {
                    ".docx" => FileKind.Docx,
                    ".xlsx" => FileKind.Xlsx,
                    ".pptx" => FileKind.Pptx,
                    _ => (FileKind?)null
                };
            return kind is null
                ? new DetectionResult(null, null, MismatchWarning(claimed, null, contentType))
                : Result(kind.Value, null, claimed, contentType);
        }

        if (claimed == FileKind.Html || ext is ".html" or ".htm" || LooksLikeHtml(content))
            return Result(FileKind.Html, null, claimed, contentType);

        return new DetectionResult(null, null, MismatchWarning(claimed, null, contentType));
    }

    private static DetectionResult Result(FileKind kind, string? mediaType, FileKind? claimed, string? contentType) =>
        new(kind, mediaType, MismatchWarning(claimed, kind, contentType));

    private static string? MismatchWarning(FileKind? claimed, FileKind? detected, string? contentType) =>
        claimed is not null && claimed != detected
            ? $"claimed content type '{contentType}' does not match the file bytes"
            : null;

    private static FileKind? ClaimedKind(string? contentType)
    {
        if (string.IsNullOrWhiteSpace(contentType)) return null;
        var mediaType = contentType.Split(';')[0].Trim().ToLowerInvariant();
        return mediaType switch
        {
            "application/pdf" => FileKind.Pdf,
            "text/html" => FileKind.Html,
            "image/png" or "image/jpeg" or "image/jpg" => FileKind.Image,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => FileKind.Docx,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => FileKind.Xlsx,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" => FileKind.Pptx,
            _ => null
        };
    }

    private static bool LooksLikeHtml(ReadOnlySpan<byte> content)
    {
        var i = 0;
        while (i < content.Length && (content[i] == (byte)' ' || content[i] == (byte)'\t'
            || content[i] == (byte)'\r' || content[i] == (byte)'\n')) i++;
        var rest = content[i..];
        return StartsWithAscii(rest, "<!doctype") || StartsWithAscii(rest, "<html");
    }

    private static bool StartsWithAscii(ReadOnlySpan<byte> content, string prefix)
    {
        if (content.Length < prefix.Length) return false;
        for (var i = 0; i < prefix.Length; i++)
            if (char.ToLowerInvariant((char)content[i]) != prefix[i]) return false;
        return true;
    }
}
```

- [ ] **Step 5: Run detector tests to verify they pass**

```bash
dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx --filter "FullyQualifiedName~FileKindDetectorTests"
```

Expected: PASS (7 tests).

- [ ] **Step 6: Write the failing hints + reader tests**

Create `tests/DocInt.Tests/HintsParserTests.cs`:

```csharp
using DocInt.Api.Validation;

namespace DocInt.Tests;

public class HintsParserTests
{
    [Fact]
    public void Parses_filename_to_purpose_map()
    {
        var hints = HintsParser.Parse("""{"bom.xlsx":{"purpose":"bom"},"p.png":{"purpose":"photo"}}""");
        Assert.Equal("bom", hints["bom.xlsx"]);
        Assert.Equal("photo", hints["p.png"]);
    }

    [Fact]
    public void Entry_without_purpose_is_ignored()
    {
        var hints = HintsParser.Parse("""{"a.pdf":{}}""");
        Assert.Empty(hints);
    }

    [Fact]
    public void Invalid_json_throws_bad_request()
    {
        Assert.Throws<BadExtractRequestException>(() => HintsParser.Parse("not json"));
        Assert.Throws<BadExtractRequestException>(() => HintsParser.Parse("[1,2]"));
    }
}
```

Create `tests/DocInt.Tests/MultipartReaderTests.cs`:

```csharp
using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using DocInt.Api.Validation;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Options;

namespace DocInt.Tests;

public class MultipartReaderTests
{
    private static MultipartExtractRequestReader Reader(DocIntOptions? options = null) =>
        new(Microsoft.Extensions.Options.Options.Create(options ?? new DocIntOptions()));

    private static async Task<HttpRequest> RequestOf(MultipartFormDataContent form)
    {
        var context = new DefaultHttpContext();
        context.Request.ContentType = form.Headers.ContentType!.ToString();
        var stream = new MemoryStream();
        await form.CopyToAsync(stream);
        stream.Position = 0;
        context.Request.Body = stream;
        context.Request.ContentLength = stream.Length;
        return context.Request;
    }

    [Fact]
    public async Task Reads_files_with_kinds_in_order()
    {
        using var form = Multipart.Form(
            ("manual.pdf", TestBytes.Pdf, "application/pdf"),
            ("photo.png", TestBytes.Png, "image/png"));
        var items = await Reader().ReadAsync(await RequestOf(form), CancellationToken.None);
        Assert.Equal(2, items.Count);
        Assert.Equal(("manual.pdf", FileKind.Pdf, 0), (items[0].Name, items[0].Kind, items[0].Index));
        Assert.Equal(("photo.png", FileKind.Image, "image/png"), (items[1].Name, items[1].Kind, items[1].ImageMediaType));
        Assert.Equal(TestBytes.Pdf, items[0].Bytes);
    }

    [Fact]
    public async Task Known_hint_sets_purpose_unknown_hint_warns()
    {
        using var form = Multipart.Form(
            ("a.pdf", TestBytes.Pdf, "application/pdf"),
            ("b.pdf", TestBytes.Pdf, "application/pdf"))
            .WithHints("""{"a.pdf":{"purpose":"bom"},"b.pdf":{"purpose":"mystery"}}""");
        var items = await Reader().ReadAsync(await RequestOf(form), CancellationToken.None);
        Assert.Equal("bom", items[0].Purpose);
        Assert.Null(items[1].Purpose);
        Assert.Contains("unknown purpose hint 'mystery' ignored", items[1].Warnings);
    }

    [Fact]
    public async Task Oversized_file_gets_too_large_error()
    {
        using var form = Multipart.Form(("big.pdf", TestBytes.Pdf, "application/pdf"));
        var items = await Reader(new DocIntOptions { MaxFileBytes = 4 }).ReadAsync(await RequestOf(form), CancellationToken.None);
        Assert.Equal(ErrorCodes.TooLarge, items[0].Error!.Code);
    }

    [Fact]
    public async Task Empty_file_gets_empty_file_error()
    {
        using var form = Multipart.Form(("empty.pdf", [], "application/pdf"));
        var items = await Reader().ReadAsync(await RequestOf(form), CancellationToken.None);
        Assert.Equal(ErrorCodes.EmptyFile, items[0].Error!.Code);
    }

    [Fact]
    public async Task Undetectable_file_gets_unsupported_type_error()
    {
        using var form = Multipart.Form(("x.bin", TestBytes.Garbage, "application/octet-stream"));
        var items = await Reader().ReadAsync(await RequestOf(form), CancellationToken.None);
        Assert.Equal(ErrorCodes.UnsupportedType, items[0].Error!.Code);
    }

    [Fact]
    public async Task Malformed_requests_throw_bad_request()
    {
        // zero files
        using var empty = new MultipartFormDataContent();
        empty.Add(new StringContent("{}"), "hints");
        await Assert.ThrowsAsync<BadExtractRequestException>(async () =>
            await Reader().ReadAsync(await RequestOf(empty), CancellationToken.None));

        // too many files
        using var many = Multipart.Form(
            ("a.pdf", TestBytes.Pdf, "application/pdf"),
            ("b.pdf", TestBytes.Pdf, "application/pdf"),
            ("c.pdf", TestBytes.Pdf, "application/pdf"));
        await Assert.ThrowsAsync<BadExtractRequestException>(async () =>
            await Reader(new DocIntOptions { MaxFilesPerRequest = 2 }).ReadAsync(await RequestOf(many), CancellationToken.None));

        // not multipart
        var context = new DefaultHttpContext();
        context.Request.ContentType = "application/json";
        context.Request.Body = new MemoryStream("{}"u8.ToArray());
        await Assert.ThrowsAsync<BadExtractRequestException>(async () =>
            await Reader().ReadAsync(context.Request, CancellationToken.None));

        // malformed hints
        using var badHints = Multipart.Form(("a.pdf", TestBytes.Pdf, "application/pdf")).WithHints("nope");
        await Assert.ThrowsAsync<BadExtractRequestException>(async () =>
            await Reader().ReadAsync(await RequestOf(badHints), CancellationToken.None));
    }

    [Fact]
    public async Task Oversized_hints_part_throws_bad_request()
    {
        var padding = new string('a', 300_000);
        using var form = Multipart.Form(("a.pdf", TestBytes.Pdf, "application/pdf"))
            .WithHints("{\"a.pdf\":{\"purpose\":\"" + padding + "\"}}");
        await Assert.ThrowsAsync<BadExtractRequestException>(async () =>
            await Reader().ReadAsync(await RequestOf(form), CancellationToken.None));
    }

    [Fact]
    public async Task Declared_content_length_over_request_cap_throws_bad_request()
    {
        using var form = Multipart.Form(("a.pdf", TestBytes.Pdf, "application/pdf"));
        var request = await RequestOf(form);
        request.ContentLength = long.MaxValue;
        await Assert.ThrowsAsync<BadExtractRequestException>(async () =>
            await Reader().ReadAsync(request, CancellationToken.None));
    }
}
```

- [ ] **Step 7: Run to verify compile failure**

```bash
dotnet build --no-restore src/DocInt.slnx
```

Expected: FAILS — `HintsParser`, `MultipartExtractRequestReader`, `BadExtractRequestException` missing.

- [ ] **Step 8: Implement hints parser and reader**

Create `src/DocInt.Api/Validation/HintsParser.cs`:

```csharp
using System.Text.Json;
using DocInt.Api.Contracts;

namespace DocInt.Api.Validation;

public sealed class BadExtractRequestException(string detail) : Exception(detail);

public static class HintsParser
{
    public sealed record HintEntry(string? Purpose);

    /// <summary>Returns filename → raw purpose string. Purpose validity is applied later per file.</summary>
    public static Dictionary<string, string> Parse(string json)
    {
        try
        {
            var entries = JsonSerializer.Deserialize<Dictionary<string, HintEntry>>(json, DocIntJson.Options)
                ?? throw new BadExtractRequestException("hints must be a JSON object");
            return entries
                .Where(kv => !string.IsNullOrWhiteSpace(kv.Value?.Purpose))
                .ToDictionary(kv => kv.Key, kv => kv.Value!.Purpose!, StringComparer.Ordinal);
        }
        catch (JsonException)
        {
            throw new BadExtractRequestException("hints part is not a valid JSON object of {\"<filename>\":{\"purpose\":\"...\"}}");
        }
    }
}
```

Create `src/DocInt.Api/Validation/MultipartExtractRequestReader.cs`:

```csharp
using System.Text;
using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using DocInt.Api.Engines;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Options;
using Microsoft.Net.Http.Headers;

namespace DocInt.Api.Validation;

/// <summary>
/// Streams the multipart body: buffers each 'files' part in memory up to MaxFileBytes
/// (a too-large part is rejected at the cap, never fully buffered), collects the optional
/// 'hints' part (capped at MaxHintsBytes, same never-fully-buffered rule), detects kinds,
/// and applies purpose hints. Nothing touches disk.
/// </summary>
public sealed class MultipartExtractRequestReader(IOptions<DocIntOptions> options)
{
    /// <summary>Far beyond any legitimate hints object for &lt;=32 files; guards against unbounded buffering.</summary>
    private const int MaxHintsBytes = 262_144;

    private readonly DocIntOptions _options = options.Value;

    public async Task<IReadOnlyList<FileItem>> ReadAsync(HttpRequest request, CancellationToken ct)
    {
        if (request.ContentLength is { } declared && declared > _options.MaxRequestBytes)
            throw new BadExtractRequestException(
                $"request body of {declared} bytes exceeds the limit of {_options.MaxRequestBytes} bytes");

        if (!MediaTypeHeaderValue.TryParse(request.ContentType, out var mediaType)
            || !"multipart/form-data".Equals(mediaType.MediaType.Value, StringComparison.OrdinalIgnoreCase))
            throw new BadExtractRequestException("request must be multipart/form-data");
        var boundary = HeaderUtilities.RemoveQuotes(mediaType.Boundary).Value;
        if (string.IsNullOrWhiteSpace(boundary))
            throw new BadExtractRequestException("multipart boundary missing");

        var files = new List<FileItem>();
        string? hintsJson = null;
        var reader = new MultipartReader(boundary, request.Body);
        while (await reader.ReadNextSectionAsync(ct) is { } section)
        {
            if (!ContentDispositionHeaderValue.TryParse(section.ContentDisposition, out var disposition))
                continue;
            var partName = HeaderUtilities.RemoveQuotes(disposition.Name).Value;

            if (disposition.IsFileDisposition() && partName == "files")
            {
                if (files.Count >= _options.MaxFilesPerRequest)
                    throw new BadExtractRequestException(
                        $"more than {_options.MaxFilesPerRequest} files in one request");
                var fileName = HeaderUtilities.RemoveQuotes(disposition.FileName).Value;
                if (string.IsNullOrEmpty(fileName)) fileName = $"file-{files.Count}";
                var (bytes, tooLarge) = await BufferAsync(section.Body, _options.MaxFileBytes, ct);
                var item = new FileItem
                {
                    Index = files.Count,
                    Name = fileName,
                    ClaimedContentType = section.ContentType,
                    Bytes = bytes
                };
                if (tooLarge)
                    item.Error = new FileError(ErrorCodes.TooLarge,
                        $"file exceeds the per-file limit of {_options.MaxFileBytes} bytes");
                else if (bytes.Length == 0)
                    item.Error = new FileError(ErrorCodes.EmptyFile, "file is empty");
                else
                {
                    var detection = FileKindDetector.Detect(fileName, section.ContentType, bytes);
                    if (detection.Warning is not null) item.Warnings.Add(detection.Warning);
                    if (detection.Kind is null)
                        item.Error = new FileError(ErrorCodes.UnsupportedType,
                            $"could not detect a supported file type for '{fileName}'");
                    else
                    {
                        item.Kind = detection.Kind;
                        item.ImageMediaType = detection.ImageMediaType;
                    }
                }
                files.Add(item);
            }
            else if (disposition.IsFormDisposition() && partName == "hints")
            {
                var (bytes, tooLarge) = await BufferAsync(section.Body, MaxHintsBytes, ct);
                if (tooLarge)
                    throw new BadExtractRequestException(
                        $"hints part exceeds the limit of {MaxHintsBytes} bytes");
                hintsJson = Encoding.UTF8.GetString(bytes);
            }
        }

        if (files.Count == 0)
            throw new BadExtractRequestException("request contains no file parts named 'files'");
        if (hintsJson is not null)
            ApplyHints(files, HintsParser.Parse(hintsJson));
        return files;
    }

    private static void ApplyHints(List<FileItem> files, Dictionary<string, string> hints)
    {
        foreach (var file in files)
        {
            if (!hints.TryGetValue(file.Name, out var purpose)) continue;
            if (PurposeHints.Known.Contains(purpose))
                file.Purpose = purpose.ToLowerInvariant();
            else
                file.Warnings.Add($"unknown purpose hint '{purpose}' ignored");
        }
    }

    private static async Task<(byte[] Bytes, bool TooLarge)> BufferAsync(Stream body, long maxBytes, CancellationToken ct)
    {
        using var buffered = new MemoryStream();
        var buffer = new byte[81920];
        while (true)
        {
            var read = await body.ReadAsync(buffer, ct);
            if (read == 0) break;
            if (buffered.Length + read > maxBytes)
            {
                while (await body.ReadAsync(buffer, ct) != 0) { }
                return ([], true);
            }
            buffered.Write(buffer, 0, read);
        }
        return (buffered.ToArray(), false);
    }
}
```

- [ ] **Step 9: Run tests to verify they pass**

```bash
dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (all tests; 16 new in this task: 7 detector + 3 hints + 6 reader).

- [ ] **Step 10: Gate, commit, merge**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
git add -A && git commit -m "Step 3: kind detection, hints parsing, streaming multipart reader"
git checkout main && git merge step-03-validation && git branch -d step-03-validation && git push origin main
```

---

### Task 4: Engine seam, router, extraction service, `/v1/extract` endpoint (contract frozen against test fakes)

**Files:**
- Create: `src/DocInt.Api/Engines/IExtractionEngine.cs`, `src/DocInt.Api/Engines/Errors.cs`, `src/DocInt.Api/Engines/EngineRouter.cs`, `src/DocInt.Api/Engines/ExtractionService.cs`, `src/DocInt.Api/Api/ExtractEndpoint.cs`
- Modify: `src/DocInt.Api/Program.cs`
- Test: `tests/DocInt.Tests/ContractTestFactory.cs`, `tests/DocInt.Tests/ExtractContractTests.cs`, `tests/DocInt.Tests/ExtractionServiceTests.cs`

**Interfaces:**
- Consumes: `FileItem`, contracts, options, `MultipartExtractRequestReader`, `BadExtractRequestException`.
- Produces: `EngineOutcome(FileResult Result, int PagesProcessed)`; `IExtractionEngine { IReadOnlyCollection<FileKind> Kinds; Task<EngineOutcome> ExtractAsync(FileItem, CancellationToken) }`; `EngineUnconfiguredException(string)`; `Errors.For(FileItem, string code, string message) → EngineOutcome`; `EngineRouter.RouteAsync(FileItem, CancellationToken)` (owns per-file timeout + catch-all); `ExtractionService.ExtractAsync(IReadOnlyList<FileItem>, CancellationToken) → ExtractResponse` (bounded parallelism, request-order results, per-file log line); `POST /v1/extract` mapped by `app.MapExtract()`. Test-side: `ContractTestFactory` (fake engines per kind), `FakeEngine`.

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-04-endpoint
```

- [ ] **Step 2: Write the failing contract tests**

Create `tests/DocInt.Tests/ContractTestFactory.cs`:

```csharp
using DocInt.Api.Contracts;
using DocInt.Api.Engines;
using Microsoft.Extensions.DependencyInjection;

namespace DocInt.Tests;

/// <summary>A fake engine per kind group; replaced by real engines task by task.</summary>
public sealed class FakeEngine(IReadOnlyCollection<FileKind> kinds, Func<FileItem, EngineOutcome> produce)
    : IExtractionEngine
{
    public IReadOnlyCollection<FileKind> Kinds => kinds;

    public Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct) =>
        Task.FromResult(produce(file));

    public static EngineOutcome Markdown(FileItem f, string markdown, int pages) =>
        new(new FileResult(f.Name, f.Kind, markdown, null, null, f.Warnings.ToArray(), null), pages);

    public static EngineOutcome Image(FileItem f, string description) =>
        new(new FileResult(f.Name, f.Kind, null, null, description, f.Warnings.ToArray(), null), 1);
}

public class ContractTestFactory : DocIntAppFactory
{
    protected override void ConfigureFakes(IServiceCollection services)
    {
        services.AddSingleton<IExtractionEngine>(new FakeEngine(
            [FileKind.Pdf, FileKind.Docx, FileKind.Pptx, FileKind.Html],
            f => FakeEngine.Markdown(f, "LAYOUT_MD_SENTINEL", 3)));
        services.AddSingleton<IExtractionEngine>(new FakeEngine(
            [FileKind.Xlsx],
            f => FakeEngine.Markdown(f, "XLSX_MD_SENTINEL", 1)));
        services.AddSingleton<IExtractionEngine>(new FakeEngine(
            [FileKind.Image],
            f => FakeEngine.Image(f, "VISION_DESCRIPTION_SENTINEL")));
    }
}
```

Create `tests/DocInt.Tests/ExtractContractTests.cs`:

```csharp
using System.Net;
using System.Text.Json;
using DocInt.Api.Contracts;
using Microsoft.AspNetCore.Hosting;

namespace DocInt.Tests;

public class ExtractContractTests : IClassFixture<ContractTestFactory>
{
    private readonly ContractTestFactory _factory;

    public ExtractContractTests(ContractTestFactory factory) => _factory = factory;

    private static async Task<ExtractResponse> ReadResponse(HttpResponseMessage response)
    {
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var json = await response.Content.ReadAsStringAsync();
        return JsonSerializer.Deserialize<ExtractResponse>(json, DocIntJson.Options)!;
    }

    [Fact]
    public async Task Mixed_batch_returns_per_file_results_in_request_order()
    {
        using var form = Multipart.Form(
            ("manual.pdf", TestBytes.Pdf, "application/pdf"),
            ("junk.bin", TestBytes.Garbage, "application/octet-stream"),
            ("photo.png", TestBytes.Png, "image/png"));
        var result = await ReadResponse(await _factory.CreateClient().PostAsync("/v1/extract", form));

        Assert.Equal(3, result.Files.Count);
        Assert.Equal(["manual.pdf", "junk.bin", "photo.png"], result.Files.Select(f => f.Name).ToArray());
        Assert.Equal("LAYOUT_MD_SENTINEL", result.Files[0].Markdown);
        Assert.Equal(FileKind.Pdf, result.Files[0].Kind);
        Assert.Null(result.Files[0].Error);
        Assert.Equal(ErrorCodes.UnsupportedType, result.Files[1].Error!.Code);
        Assert.Equal("VISION_DESCRIPTION_SENTINEL", result.Files[2].ImageDescription);
    }

    [Fact]
    public async Task Unknown_hint_returns_warning_on_that_file()
    {
        using var form = Multipart.Form(("manual.pdf", TestBytes.Pdf, "application/pdf"))
            .WithHints("""{"manual.pdf":{"purpose":"mystery"}}""");
        var result = await ReadResponse(await _factory.CreateClient().PostAsync("/v1/extract", form));
        Assert.Contains("unknown purpose hint 'mystery' ignored", result.Files[0].Warnings);
    }

    [Fact]
    public async Task Malformed_requests_return_400_problem()
    {
        var client = _factory.CreateClient();

        var notMultipart = await client.PostAsync("/v1/extract",
            new StringContent("{}", System.Text.Encoding.UTF8, "application/json"));
        Assert.Equal(HttpStatusCode.BadRequest, notMultipart.StatusCode);

        using var noFiles = new MultipartFormDataContent();
        noFiles.Add(new StringContent("{}"), "hints");
        Assert.Equal(HttpStatusCode.BadRequest, (await client.PostAsync("/v1/extract", noFiles)).StatusCode);

        using var badHints = Multipart.Form(("a.pdf", TestBytes.Pdf, "application/pdf")).WithHints("nope");
        Assert.Equal(HttpStatusCode.BadRequest, (await client.PostAsync("/v1/extract", badHints)).StatusCode);
    }

    [Fact]
    public async Task Too_many_files_returns_400()
    {
        using var capped = new CappedFactory();
        using var form = Multipart.Form(
            ("a.pdf", TestBytes.Pdf, "application/pdf"),
            ("b.pdf", TestBytes.Pdf, "application/pdf"),
            ("c.pdf", TestBytes.Pdf, "application/pdf"));
        var response = await capped.CreateClient().PostAsync("/v1/extract", form);
        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task Oversized_file_gets_per_file_too_large_inside_200()
    {
        using var tiny = new TinyFileFactory();
        using var form = Multipart.Form(
            ("big.pdf", TestBytes.Pdf, "application/pdf"),
            ("photo.png", TestBytes.Png, "image/png"));
        var result = await ReadResponse(await tiny.CreateClient().PostAsync("/v1/extract", form));
        Assert.Equal(ErrorCodes.TooLarge, result.Files[0].Error!.Code);
        Assert.Null(result.Files[1].Error);
    }

    [Fact]
    public async Task Kind_without_registered_engine_gets_engine_error()
    {
        using var pdfOnly = new PdfOnlyEngineFactory();
        using var form = Multipart.Form(("photo.png", TestBytes.Png, "image/png"));
        var result = await ReadResponse(await pdfOnly.CreateClient().PostAsync("/v1/extract", form));
        Assert.Equal(ErrorCodes.EngineError, result.Files[0].Error!.Code);
        Assert.Contains("no engine registered", result.Files[0].Error!.Message);
    }

    private sealed class CappedFactory : ContractTestFactory
    {
        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.UseSetting("DocInt:MaxFilesPerRequest", "2");
            base.ConfigureWebHost(builder);
        }
    }

    private sealed class TinyFileFactory : ContractTestFactory
    {
        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.UseSetting("DocInt:MaxFileBytes", "10");
            base.ConfigureWebHost(builder);
        }
    }

    private sealed class PdfOnlyEngineFactory : DocIntAppFactory
    {
        protected override void ConfigureFakes(Microsoft.Extensions.DependencyInjection.IServiceCollection services)
        {
            services.AddSingleton<DocInt.Api.Engines.IExtractionEngine>(new FakeEngine(
                [FileKind.Pdf], f => FakeEngine.Markdown(f, "x", 1)));
        }
    }
}
```

Create `tests/DocInt.Tests/ExtractionServiceTests.cs`:

```csharp
using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using DocInt.Api.Engines;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DocInt.Tests;

public class ExtractionServiceTests
{
    private sealed class CountingEngine : IExtractionEngine
    {
        private int _inFlight;
        public int MaxObserved;
        public IReadOnlyCollection<FileKind> Kinds { get; } = [FileKind.Pdf];

        public async Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct)
        {
            var now = Interlocked.Increment(ref _inFlight);
            InterlockedExtensions.Max(ref MaxObserved, now);
            await Task.Delay(50, ct);
            Interlocked.Decrement(ref _inFlight);
            return FakeEngine.Markdown(file, $"md-{file.Index}", 1);
        }
    }

    private static class InterlockedExtensions
    {
        public static void Max(ref int location, int value)
        {
            int current;
            while (value > (current = Volatile.Read(ref location)))
                Interlocked.CompareExchange(ref location, value, current);
        }
    }

    [Fact]
    public async Task Respects_parallelism_cap_and_preserves_request_order()
    {
        var options = Microsoft.Extensions.Options.Options.Create(
            new DocIntOptions { MaxParallelism = 2 });
        var engine = new CountingEngine();
        var service = new ExtractionService(
            new EngineRouter([engine], options), options, NullLogger<ExtractionService>.Instance);

        var files = Enumerable.Range(0, 8).Select(i => new FileItem
        {
            Index = i, Name = $"f{i}.pdf", Kind = FileKind.Pdf, Bytes = TestBytes.Pdf
        }).ToArray();

        var response = await service.ExtractAsync(files, CancellationToken.None);

        Assert.Equal(8, response.Files.Count);
        Assert.All(Enumerable.Range(0, 8), i => Assert.Equal($"md-{i}", response.Files[i].Markdown));
        Assert.True(engine.MaxObserved <= 2, $"observed {engine.MaxObserved} concurrent extractions");
    }

    [Fact]
    public async Task Timeout_produces_per_file_timeout_error()
    {
        var options = Microsoft.Extensions.Options.Options.Create(
            new DocIntOptions { PerFileTimeoutSeconds = 1 });
        var hanging = new FakeHangingEngine();
        var service = new ExtractionService(
            new EngineRouter([hanging], options), options, NullLogger<ExtractionService>.Instance);

        var files = new[] { new FileItem { Index = 0, Name = "slow.pdf", Kind = FileKind.Pdf, Bytes = TestBytes.Pdf } };
        var response = await service.ExtractAsync(files, CancellationToken.None);

        Assert.Equal(ErrorCodes.Timeout, response.Files[0].Error!.Code);
    }

    private sealed class FakeHangingEngine : IExtractionEngine
    {
        public IReadOnlyCollection<FileKind> Kinds { get; } = [FileKind.Pdf];
        public async Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct)
        {
            await Task.Delay(TimeSpan.FromMinutes(5), ct);
            throw new UnreachableException();
        }
    }
}
```

Add `using System.Diagnostics;` to the top of `ExtractionServiceTests.cs` for `UnreachableException`.

- [ ] **Step 3: Run to verify compile failure**

```bash
dotnet build --no-restore src/DocInt.slnx
```

Expected: FAILS — `IExtractionEngine`, `EngineRouter`, `ExtractionService`, `EngineOutcome` missing.

- [ ] **Step 4: Implement the seam, router, service, endpoint**

Create `src/DocInt.Api/Engines/IExtractionEngine.cs`:

```csharp
using DocInt.Api.Contracts;

namespace DocInt.Api.Engines;

/// <summary>Result plus the pages/sheets/images processed (feeds the docint.pages_processed metric).</summary>
public sealed record EngineOutcome(FileResult Result, int PagesProcessed);

public interface IExtractionEngine
{
    IReadOnlyCollection<FileKind> Kinds { get; }
    Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct);
}
```

Create `src/DocInt.Api/Engines/Errors.cs`:

```csharp
using DocInt.Api.Contracts;

namespace DocInt.Api.Engines;

public sealed class EngineUnconfiguredException(string message) : Exception(message);

public static class Errors
{
    public static EngineOutcome For(FileItem file, string code, string message) =>
        new(new FileResult(file.Name, file.Kind, null, null, null, file.Warnings.ToArray(),
            new FileError(code, message)), 0);
}
```

Create `src/DocInt.Api/Engines/EngineRouter.cs`:

```csharp
using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using Microsoft.Extensions.Options;

namespace DocInt.Api.Engines;

/// <summary>Picks the engine by detected kind; owns the per-file timeout and the catch-all.</summary>
public sealed class EngineRouter
{
    private readonly Dictionary<FileKind, IExtractionEngine> _byKind = [];
    private readonly TimeSpan _timeout;

    public EngineRouter(IEnumerable<IExtractionEngine> engines, IOptions<DocIntOptions> options)
    {
        foreach (var engine in engines)
            foreach (var kind in engine.Kinds)
                _byKind[kind] = engine;
        _timeout = TimeSpan.FromSeconds(options.Value.PerFileTimeoutSeconds);
    }

    public async Task<EngineOutcome> RouteAsync(FileItem file, CancellationToken requestCt)
    {
        var kind = file.Kind!.Value;
        if (!_byKind.TryGetValue(kind, out var engine))
            return Errors.For(file, ErrorCodes.EngineError, $"no engine registered for kind '{kind.Name()}'");

        using var cts = CancellationTokenSource.CreateLinkedTokenSource(requestCt);
        cts.CancelAfter(_timeout);
        try
        {
            return await engine.ExtractAsync(file, cts.Token);
        }
        catch (OperationCanceledException) when (cts.IsCancellationRequested && !requestCt.IsCancellationRequested)
        {
            return Errors.For(file, ErrorCodes.Timeout,
                $"extraction exceeded the per-file timeout of {_timeout.TotalSeconds:0}s");
        }
        catch (EngineUnconfiguredException ex)
        {
            return Errors.For(file, ErrorCodes.EngineUnconfigured, ex.Message);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            // Exception message only — never document content.
            return Errors.For(file, ErrorCodes.EngineError, ex.Message);
        }
    }
}
```

Create `src/DocInt.Api/Engines/ExtractionService.cs`:

```csharp
using System.Diagnostics;
using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using Microsoft.Extensions.Options;

namespace DocInt.Api.Engines;

/// <summary>Runs a batch: bounded parallelism, request-order results, per-file log line (no content).</summary>
public sealed class ExtractionService(
    EngineRouter router,
    IOptions<DocIntOptions> options,
    ILogger<ExtractionService> logger)
{
    public async Task<ExtractResponse> ExtractAsync(IReadOnlyList<FileItem> files, CancellationToken ct)
    {
        var results = new FileResult[files.Count];
        await Parallel.ForEachAsync(files,
            new ParallelOptions { MaxDegreeOfParallelism = options.Value.MaxParallelism, CancellationToken = ct },
            async (file, token) =>
            {
                var started = Stopwatch.GetTimestamp();
                var outcome = file.Error is not null
                    ? new EngineOutcome(new FileResult(file.Name, file.Kind, null, null, null,
                        file.Warnings.ToArray(), file.Error), 0)
                    : await router.RouteAsync(file, token);
                results[file.Index] = outcome.Result;
                logger.LogInformation(
                    "Processed {FileName}: kind={Kind} sizeBytes={SizeBytes} outcome={Outcome} durationMs={DurationMs:0}",
                    file.Name, file.Kind?.Name() ?? "unknown", file.Bytes.Length,
                    outcome.Result.Error?.Code ?? "ok",
                    Stopwatch.GetElapsedTime(started).TotalMilliseconds);
            });
        return new ExtractResponse(results);
    }
}
```

Create `src/DocInt.Api/Api/ExtractEndpoint.cs`:

```csharp
using DocInt.Api.Contracts;
using DocInt.Api.Engines;
using DocInt.Api.Validation;

namespace DocInt.Api.Api;

public static class ExtractEndpoint
{
    public static IEndpointRouteBuilder MapExtract(this IEndpointRouteBuilder app)
    {
        app.MapPost("/v1/extract", Handle)
            .WithName("Extract")
            .WithSummary("Extract Markdown, typed tables and image descriptions from documents")
            .WithDescription("multipart/form-data: N file parts named 'files'; optional 'hints' part "
                + "with JSON {\"<filename>\":{\"purpose\":\"bom|photo\"}}. Well-formed requests always "
                + "return 200 with per-file success or error.")
            .Produces<ExtractResponse>(StatusCodes.Status200OK, "application/json")
            .ProducesProblem(StatusCodes.Status400BadRequest);
        return app;
    }

    private static async Task<IResult> Handle(
        HttpRequest request,
        MultipartExtractRequestReader reader,
        ExtractionService service,
        CancellationToken ct)
    {
        IReadOnlyList<FileItem> files;
        try
        {
            files = await reader.ReadAsync(request, ct);
        }
        catch (BadExtractRequestException ex)
        {
            return Results.Problem(title: "Malformed extract request", detail: ex.Message,
                statusCode: StatusCodes.Status400BadRequest);
        }
        var response = await service.ExtractAsync(files, ct);
        return Results.Json(response, DocIntJson.Options);
    }
}
```

In `src/DocInt.Api/Program.cs`:
1. Add to the using block at the top: `using DocInt.Api.Api;`, `using DocInt.Api.Engines;`, `using DocInt.Api.Validation;`
2. Insert after `builder.AddDocIntOptions();`:

```csharp
    builder.Services.AddSingleton<MultipartExtractRequestReader>();
    builder.Services.AddSingleton<EngineRouter>();
    builder.Services.AddSingleton<ExtractionService>();
```

3. Insert after `app.MapHealthChecks("/healthz");`:

```csharp
    app.MapExtract();
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (all; 8 new). The wire contract is now frozen.

- [ ] **Step 6: Gate, commit, merge**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
git add -A && git commit -m "Step 4: /v1/extract endpoint, engine seam, router with timeout, bounded-parallel service"
git checkout main && git merge step-04-endpoint && git branch -d step-04-endpoint && git push origin main
```

---

### Task 5: Golden fixture generator + committed fixtures

**Files:**
- Create: `tools/make-golden/make-golden.csproj`, `tools/make-golden/Program.cs`, `tools/make-golden/MiniPdf.cs`, `tools/make-golden/OfficeFixtures.cs`, `tools/make-golden/ImageFixtures.cs`, `tests/DocInt.Tests/golden/*` (9 committed binaries)
- Modify: `tests/DocInt.Tests/DocInt.Tests.csproj` (copy golden to output), `tests/DocInt.Tests/TestSupport.cs` (add `Golden` helper), `src/DocInt.slnx` (add tool project)
- Test: `tests/DocInt.Tests/GoldenFixturesTests.cs`

**Interfaces:**
- Produces: committed fixtures `text.pdf`, `scanned.pdf` (image-only page containing rendered text "OCR TARGET UV400"), `sample.docx` (contains "EuGo DOCX golden document"), `sample.pptx` (contains "EuGo PPTX golden slide"), `sample.html` (contains "EuGo HTML golden"), `bom.xlsx` (sheet "BoM": header row + `["M3 screw", 40, 19.99, =B2*C2 cached 799.6, TRUE, date 2026-07-19]` + `["Washer", 100, 0.125, cached 12.5, FALSE, 2026-01-02]`; sheet "Notes": one string cell), `photo.png` (sunglasses shapes + visible "UV400" text), `corrupt.xlsx` (PK header + garbage), `unknown.bin` (no known magic). Test helper `Golden.Bytes(string name) → byte[]`.
- The `text.pdf` body text is `EuGo text PDF golden - Part M3 screw UV400` — the redaction test (Task 9) greps logs for this exact string.

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-05-golden
```

- [ ] **Step 2: Write the failing fixtures test**

Append to `tests/DocInt.Tests/TestSupport.cs`:

```csharp
public static class Golden
{
    public static byte[] Bytes(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "golden", name));
}
```

Create `tests/DocInt.Tests/GoldenFixturesTests.cs`:

```csharp
using DocInt.Api.Contracts;
using DocInt.Api.Validation;

namespace DocInt.Tests;

public class GoldenFixturesTests
{
    [Theory]
    [InlineData("text.pdf", FileKind.Pdf)]
    [InlineData("scanned.pdf", FileKind.Pdf)]
    [InlineData("sample.docx", FileKind.Docx)]
    [InlineData("sample.pptx", FileKind.Pptx)]
    [InlineData("sample.html", FileKind.Html)]
    [InlineData("bom.xlsx", FileKind.Xlsx)]
    [InlineData("photo.png", FileKind.Image)]
    [InlineData("corrupt.xlsx", FileKind.Xlsx)]
    public void Fixture_exists_and_detects_as_expected_kind(string name, FileKind expected)
    {
        var bytes = Golden.Bytes(name);
        Assert.NotEmpty(bytes);
        Assert.Equal(expected, FileKindDetector.Detect(name, null, bytes).Kind);
    }

    [Fact]
    public void Unknown_bin_is_undetectable()
    {
        Assert.Null(FileKindDetector.Detect("unknown.bin", null, Golden.Bytes("unknown.bin")).Kind);
    }
}
```

Add the copy-to-output glob to `tests/DocInt.Tests/DocInt.Tests.csproj`, before `</Project>`:

```xml
  <ItemGroup>
    <Content Include="golden/**" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>
```

- [ ] **Step 3: Run to verify failure**

```bash
dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx --filter "FullyQualifiedName~GoldenFixturesTests"
```

Expected: FAIL — `FileNotFoundException`/`DirectoryNotFoundException` (fixtures don't exist yet).

- [ ] **Step 4: Build the generator tool**

Create `tools/make-golden/make-golden.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <RootNamespace>MakeGolden</RootNamespace>
    <AssemblyName>make-golden</AssemblyName>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
    <PackageReference Include="SkiaSharp" Version="3.119.0" />
    <PackageReference Include="SkiaSharp.NativeAssets.Linux" Version="3.119.0" />
  </ItemGroup>
</Project>
```

Create `tools/make-golden/MiniPdf.cs`:

```csharp
using System.Text;

namespace MakeGolden;

/// <summary>Deterministic minimal PDF writer — enough structure for Azure Document Intelligence.</summary>
public static class MiniPdf
{
    public static byte[] TextPdf(string text) => Build(
        Ascii("<< /Type /Catalog /Pages 2 0 R >>"),
        Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"),
        Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>"),
        StreamObject(Encoding.ASCII.GetBytes($"BT /F1 24 Tf 72 700 Td ({Escape(text)}) Tj ET")),
        Ascii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"));

    public static byte[] ImagePdf(byte[] jpeg, int width, int height) => Build(
        Ascii("<< /Type /Catalog /Pages 2 0 R >>"),
        Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"),
        Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>"),
        StreamObject(Encoding.ASCII.GetBytes("q 612 0 0 792 0 0 cm /Im0 Do Q")),
        StreamObject(jpeg,
            $"/Type /XObject /Subtype /Image /Width {width} /Height {height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode "));

    private static byte[] Ascii(string s) => Encoding.ASCII.GetBytes(s);

    private static string Escape(string s) =>
        s.Replace("\\", @"\\").Replace("(", @"\(").Replace(")", @"\)");

    private static byte[] StreamObject(byte[] data, string extraDict = "")
    {
        var head = Encoding.ASCII.GetBytes($"<< {extraDict}/Length {data.Length} >>\nstream\n");
        var tail = Encoding.ASCII.GetBytes("\nendstream");
        return [.. head, .. data, .. tail];
    }

    private static byte[] Build(params byte[][] objects)
    {
        using var ms = new MemoryStream();
        void WriteAscii(string s)
        {
            var bytes = Encoding.ASCII.GetBytes(s);
            ms.Write(bytes, 0, bytes.Length);
        }

        WriteAscii("%PDF-1.4\n");
        var offsets = new long[objects.Length];
        for (var i = 0; i < objects.Length; i++)
        {
            offsets[i] = ms.Position;
            WriteAscii($"{i + 1} 0 obj\n");
            ms.Write(objects[i], 0, objects[i].Length);
            WriteAscii("\nendobj\n");
        }
        var xrefPos = ms.Position;
        WriteAscii($"xref\n0 {objects.Length + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets)
            WriteAscii($"{offset:0000000000} 00000 n \n");
        WriteAscii($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xrefPos}\n%%EOF");
        return ms.ToArray();
    }
}
```

Create `tools/make-golden/OfficeFixtures.cs`:

```csharp
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace MakeGolden;

public static class OfficeFixtures
{
    public static byte[] Docx()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("EuGo DOCX golden document"))),
                new W.Paragraph(new W.Run(new W.Text("Bill of Materials evidence text for extraction")))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    public static byte[] BomXlsx()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new S.Workbook();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new S.Stylesheet(
                new S.Fonts(new S.Font()),
                new S.Fills(new S.Fill()),
                new S.Borders(new S.Border()),
                new S.CellFormats(
                    new S.CellFormat(),
                    new S.CellFormat { NumberFormatId = 14, ApplyNumberFormat = true }));

            var shared = workbookPart.AddNewPart<SharedStringTablePart>();
            shared.SharedStringTable = new S.SharedStringTable();
            int Str(string s)
            {
                var i = 0;
                foreach (var item in shared.SharedStringTable.Elements<S.SharedStringItem>())
                {
                    if (item.InnerText == s) return i;
                    i++;
                }
                shared.SharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(s)));
                return i;
            }

            var bomSheet = workbookPart.AddNewPart<WorksheetPart>();
            bomSheet.Worksheet = new S.Worksheet(new S.SheetData(
                Row(1, SharedCell("A1", Str("Part")), SharedCell("B1", Str("Qty")),
                       SharedCell("C1", Str("UnitPrice")), SharedCell("D1", Str("Total")),
                       SharedCell("E1", Str("InStock")), SharedCell("F1", Str("Updated"))),
                Row(2, SharedCell("A2", Str("M3 screw")), NumberCell("B2", "40"), NumberCell("C2", "19.99"),
                       FormulaCell("D2", "B2*C2", "799.6"), BoolCell("E2", true),
                       DateCell("F2", new DateTime(2026, 7, 19))),
                Row(3, SharedCell("A3", Str("Washer")), NumberCell("B3", "100"), NumberCell("C3", "0.125"),
                       FormulaCell("D3", "B3*C3", "12.5"), BoolCell("E3", false),
                       DateCell("F3", new DateTime(2026, 1, 2)))));

            var notesSheet = workbookPart.AddNewPart<WorksheetPart>();
            notesSheet.Worksheet = new S.Worksheet(new S.SheetData(
                Row(1, SharedCell("A1", Str("EuGo BoM golden notes")))));

            workbookPart.Workbook.AppendChild(new S.Sheets(
                new S.Sheet { Id = workbookPart.GetIdOfPart(bomSheet), SheetId = 1U, Name = "BoM" },
                new S.Sheet { Id = workbookPart.GetIdOfPart(notesSheet), SheetId = 2U, Name = "Notes" }));
            workbookPart.Workbook.Save();
        }
        return ms.ToArray();

        static S.Row Row(uint index, params S.Cell[] cells)
        {
            var row = new S.Row { RowIndex = index };
            row.Append(cells);
            return row;
        }
        static S.Cell SharedCell(string r, int sharedIndex) => new()
        { CellReference = r, DataType = S.CellValues.SharedString, CellValue = new S.CellValue(sharedIndex.ToString()) };
        static S.Cell NumberCell(string r, string number) => new()
        { CellReference = r, CellValue = new S.CellValue(number) };
        static S.Cell FormulaCell(string r, string formula, string cached) => new()
        { CellReference = r, CellFormula = new S.CellFormula(formula), CellValue = new S.CellValue(cached) };
        static S.Cell BoolCell(string r, bool value) => new()
        { CellReference = r, DataType = S.CellValues.Boolean, CellValue = new S.CellValue(value ? "1" : "0") };
        static S.Cell DateCell(string r, DateTime date) => new()
        { CellReference = r, StyleIndex = 1U, CellValue = new S.CellValue(date.ToOADate().ToString(CultureInfo.InvariantCulture)) };
    }

    public static byte[] Pptx(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
        {
            var presentationPart = doc.AddPresentationPart();
            presentationPart.Presentation = new P.Presentation();

            var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rIdMaster");
            masterPart.SlideMaster = new P.SlideMaster(
                new P.CommonSlideData(EmptyShapeTree()),
                new P.ColorMap
                {
                    Background1 = D.ColorSchemeIndexValues.Light1,
                    Text1 = D.ColorSchemeIndexValues.Dark1,
                    Background2 = D.ColorSchemeIndexValues.Light2,
                    Text2 = D.ColorSchemeIndexValues.Dark2,
                    Accent1 = D.ColorSchemeIndexValues.Accent1,
                    Accent2 = D.ColorSchemeIndexValues.Accent2,
                    Accent3 = D.ColorSchemeIndexValues.Accent3,
                    Accent4 = D.ColorSchemeIndexValues.Accent4,
                    Accent5 = D.ColorSchemeIndexValues.Accent5,
                    Accent6 = D.ColorSchemeIndexValues.Accent6,
                    Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                });

            var themePart = masterPart.AddNewPart<ThemePart>("rIdTheme");
            themePart.Theme = MinimalTheme();

            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rIdLayout");
            layoutPart.SlideLayout = new P.SlideLayout(
                new P.CommonSlideData(EmptyShapeTree()),
                new P.ColorMapOverride(new D.MasterColorMapping()));
            masterPart.SlideMaster.AppendChild(new P.SlideLayoutIdList(
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = "rIdLayout" }));

            var slidePart = presentationPart.AddNewPart<SlidePart>("rIdSlide");
            slidePart.AddPart(layoutPart);
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 2U, Name = "Title" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle(),
                            new D.Paragraph(new D.Run(new D.Text(text))))))),
                new P.ColorMapOverride(new D.MasterColorMapping()));

            presentationPart.Presentation.Append(
                new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = "rIdMaster" }),
                new P.SlideIdList(new P.SlideId { Id = 256U, RelationshipId = "rIdSlide" }),
                new P.SlideSize { Cx = 9144000, Cy = 6858000 },
                new P.NotesSize { Cx = 6858000, Cy = 9144000 });
            presentationPart.Presentation.Save();
        }
        return ms.ToArray();

        static P.ShapeTree EmptyShapeTree() => new(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties());
    }

    private static D.Theme MinimalTheme() => new(
        new D.ThemeElements(
            new D.ColorScheme(
                new D.Dark1Color(new D.SystemColor { Val = D.SystemColorValues.WindowText }),
                new D.Light1Color(new D.SystemColor { Val = D.SystemColorValues.Window }),
                new D.Dark2Color(new D.RgbColorModelHex { Val = "44546A" }),
                new D.Light2Color(new D.RgbColorModelHex { Val = "E7E6E6" }),
                new D.Accent1Color(new D.RgbColorModelHex { Val = "4472C4" }),
                new D.Accent2Color(new D.RgbColorModelHex { Val = "ED7D31" }),
                new D.Accent3Color(new D.RgbColorModelHex { Val = "A5A5A5" }),
                new D.Accent4Color(new D.RgbColorModelHex { Val = "FFC000" }),
                new D.Accent5Color(new D.RgbColorModelHex { Val = "5B9BD5" }),
                new D.Accent6Color(new D.RgbColorModelHex { Val = "70AD47" }),
                new D.Hyperlink(new D.RgbColorModelHex { Val = "0563C1" }),
                new D.FollowedHyperlinkColor(new D.RgbColorModelHex { Val = "954F72" }))
            { Name = "Office" },
            new D.FontScheme(
                new D.MajorFont(
                    new D.LatinFont { Typeface = "Calibri Light" },
                    new D.EastAsianFont { Typeface = "" },
                    new D.ComplexScriptFont { Typeface = "" }),
                new D.MinorFont(
                    new D.LatinFont { Typeface = "Calibri" },
                    new D.EastAsianFont { Typeface = "" },
                    new D.ComplexScriptFont { Typeface = "" }))
            { Name = "Office" },
            new D.FormatScheme(
                new D.FillStyleList(SolidPh(), SolidPh(), SolidPh()),
                new D.LineStyleList(
                    new D.Outline(SolidPh()),
                    new D.Outline(SolidPh()),
                    new D.Outline(SolidPh())),
                new D.EffectStyleList(
                    new D.EffectStyle(new D.EffectList()),
                    new D.EffectStyle(new D.EffectList()),
                    new D.EffectStyle(new D.EffectList())),
                new D.BackgroundFillStyleList(SolidPh(), SolidPh(), SolidPh()))
            { Name = "Office" }))
    { Name = "MinimalTheme" };

    private static D.SolidFill SolidPh() =>
        new(new D.SchemeColor { Val = D.SchemeColorValues.PhColor });
}
```

Create `tools/make-golden/ImageFixtures.cs`:

```csharp
using SkiaSharp;

namespace MakeGolden;

public static class ImageFixtures
{
    /// <summary>Synthetic sunglasses product photo with visible "UV400" text.</summary>
    public static byte[] SunglassesPng()
    {
        using var surface = SKSurface.Create(new SKImageInfo(640, 400));
        var canvas = surface.Canvas;
        canvas.Clear(new SKColor(230, 230, 235));

        using var frame = new SKPaint { Color = new SKColor(25, 25, 25), IsAntialias = true };
        using var lens = new SKPaint { Color = new SKColor(60, 45, 95), IsAntialias = true };
        canvas.DrawRect(new SKRect(130, 150, 510, 165), frame);
        canvas.DrawCircle(210, 200, 75, frame);
        canvas.DrawCircle(430, 200, 75, frame);
        canvas.DrawCircle(210, 200, 65, lens);
        canvas.DrawCircle(430, 200, 65, lens);

        using var ink = new SKPaint { Color = new SKColor(15, 15, 15), IsAntialias = true };
        using var big = new SKFont(SKTypeface.Default, 42);
        using var small = new SKFont(SKTypeface.Default, 24);
        canvas.DrawText("UV400", 265, 330, SKTextAlign.Left, big, ink);
        canvas.DrawText("Sunglasses model SG-1", 175, 370, SKTextAlign.Left, small, ink);

        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }

    /// <summary>A "scanned page": white JPEG with rendered text only — proves OCR in live smoke.</summary>
    public static (byte[] Jpeg, int Width, int Height) ScannedPageJpeg()
    {
        const int width = 1240;
        const int height = 1650;
        using var surface = SKSurface.Create(new SKImageInfo(width, height));
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.White);

        using var ink = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        using var font = new SKFont(SKTypeface.Default, 48);
        canvas.DrawText("EuGo scanned golden page", 100, 220, SKTextAlign.Left, font, ink);
        canvas.DrawText("OCR TARGET UV400", 100, 320, SKTextAlign.Left, font, ink);
        canvas.DrawText("This page exists to prove OCR works", 100, 420, SKTextAlign.Left, font, ink);

        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Jpeg, 85);
        return (data.ToArray(), width, height);
    }
}
```

Create `tools/make-golden/Program.cs`:

```csharp
using System.Text;
using MakeGolden;

var outputDir = args.Length > 0 ? args[0] : "tests/DocInt.Tests/golden";
Directory.CreateDirectory(outputDir);

void Save(string name, byte[] bytes)
{
    File.WriteAllBytes(Path.Combine(outputDir, name), bytes);
    Console.WriteLine($"{name}: {bytes.Length} bytes");
}

Save("text.pdf", MiniPdf.TextPdf("EuGo text PDF golden - Part M3 screw UV400"));

var (jpeg, width, height) = ImageFixtures.ScannedPageJpeg();
Save("scanned.pdf", MiniPdf.ImagePdf(jpeg, width, height));

Save("sample.docx", OfficeFixtures.Docx());
Save("sample.pptx", OfficeFixtures.Pptx("EuGo PPTX golden slide"));
Save("sample.html", Encoding.UTF8.GetBytes(
    "<!DOCTYPE html><html><body><h1>EuGo HTML golden</h1>"
    + "<table><tr><th>Part</th><th>Qty</th></tr><tr><td>M3 screw</td><td>40</td></tr></table>"
    + "</body></html>"));
Save("bom.xlsx", OfficeFixtures.BomXlsx());
Save("photo.png", ImageFixtures.SunglassesPng());
Save("corrupt.xlsx", [0x50, 0x4B, 0x03, 0x04, .. Enumerable.Repeat((byte)0xDE, 64)]);
Save("unknown.bin", [.. Enumerable.Range(1, 64).Select(i => (byte)i)]);

Console.WriteLine($"Golden fixtures written to {Path.GetFullPath(outputDir)}");
```

- [ ] **Step 5: Add the tool to the solution and generate the fixtures**

```bash
dotnet sln src/DocInt.slnx add tools/make-golden/make-golden.csproj
dotnet run --project tools/make-golden
ls -la tests/DocInt.Tests/golden/
```

Expected: 9 files listed with non-zero sizes. Spot-check: `file tests/DocInt.Tests/golden/*` should report PDF, Zip/OOXML, PNG, HTML, data.

- [ ] **Step 6: Run tests to verify they pass**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (9 new fixture tests).

- [ ] **Step 7: Gate, commit (binaries included), merge**

```bash
git add -A && git status --short   # verify the 9 golden binaries are staged
git commit -m "Step 5: make-golden fixture generator + committed golden fixtures"
git checkout main && git merge step-05-golden && git branch -d step-05-golden && git push origin main
```

---

### Task 6: SpreadsheetEngine (OpenXML typed cells + Markdown)

**Files:**
- Create: `src/DocInt.Api/Engines/SpreadsheetEngine.cs`
- Modify: `src/DocInt.Api/DocInt.Api.csproj` (add OpenXml), `src/DocInt.Api/Program.cs` (register engine), `tests/DocInt.Tests/ContractTestFactory.cs` (drop the xlsx fake)
- Test: `tests/DocInt.Tests/SpreadsheetEngineTests.cs`

**Interfaces:**
- Consumes: `IExtractionEngine`, `EngineOutcome`, `Errors`, `FileItem`, golden `bom.xlsx`/`corrupt.xlsx`.
- Produces: `SpreadsheetEngine : IExtractionEngine` (Kinds = [Xlsx]) — per sheet one `TableResult(sheetName, markdownTable, typedRows)`; `FileResult.Markdown` = all sheets as `## <name>` sections; `PagesProcessed` = sheet count. Typed cells: shared/inline strings → `string`, numbers → `decimal` (fallback `double` on overflow), booleans → `bool`, date-styled numbers (NumberFormatId 14–22, 45–47) → ISO-8601 `string` (`yyyy-MM-dd` when midnight), error cells → `null` + warning, formula without cached value → `null` + warning.

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-06-spreadsheet
```

- [ ] **Step 2: Write the failing tests**

Create `tests/DocInt.Tests/SpreadsheetEngineTests.cs`:

```csharp
using DocInt.Api.Contracts;
using DocInt.Api.Engines;

namespace DocInt.Tests;

public class SpreadsheetEngineTests
{
    private static async Task<EngineOutcome> Run(string fixture)
    {
        var engine = new SpreadsheetEngine();
        var item = new FileItem { Index = 0, Name = fixture, Kind = FileKind.Xlsx, Bytes = Golden.Bytes(fixture) };
        return await engine.ExtractAsync(item, CancellationToken.None);
    }

    [Fact]
    public async Task Bom_xlsx_yields_exact_typed_cells()
    {
        var outcome = await Run("bom.xlsx");
        var bom = outcome.Result.Tables![0];

        Assert.Equal("BoM", bom.Name);
        Assert.Equal(["Part", "Qty", "UnitPrice", "Total", "InStock", "Updated"],
            bom.Rows[0].Cast<string>().ToArray());
        Assert.Equal("M3 screw", bom.Rows[1][0]);
        Assert.Equal(40m, bom.Rows[1][1]);
        Assert.Equal(19.99m, bom.Rows[1][2]);       // exact — never re-parsed display text
        Assert.Equal(799.6m, bom.Rows[1][3]);       // formula cell: cached value
        Assert.Equal(true, bom.Rows[1][4]);
        Assert.Equal("2026-07-19", bom.Rows[1][5]); // date-styled number → ISO-8601
        Assert.Equal(0.125m, bom.Rows[2][2]);
        Assert.Equal(false, bom.Rows[2][4]);
        Assert.Null(outcome.Result.Error);
    }

    [Fact]
    public async Task Markdown_renders_both_sheets()
    {
        var outcome = await Run("bom.xlsx");
        Assert.Equal(2, outcome.Result.Tables!.Count);
        Assert.Equal("Notes", outcome.Result.Tables[1].Name);
        Assert.Contains("## BoM", outcome.Result.Markdown);
        Assert.Contains("## Notes", outcome.Result.Markdown);
        Assert.Contains("| M3 screw | 40 | 19.99 |", outcome.Result.Markdown);
        Assert.Equal(2, outcome.PagesProcessed);
    }

    [Fact]
    public async Task Corrupt_xlsx_maps_to_corrupt_error()
    {
        var outcome = await Run("corrupt.xlsx");
        Assert.Equal(ErrorCodes.Corrupt, outcome.Result.Error!.Code);
    }

    [Fact]
    public async Task Http_contract_returns_typed_json_numbers()
    {
        using var factory = new ContractTestFactory();
        using var form = Multipart.Form(("bom.xlsx", Golden.Bytes("bom.xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        var response = await factory.CreateClient().PostAsync("/v1/extract", form);
        var json = await response.Content.ReadAsStringAsync();
        Assert.Contains("\"kind\":\"xlsx\"", json);
        Assert.Contains("19.99", json);
        Assert.DoesNotContain("\"19.99\"", json);   // number, not string
        Assert.Contains("\"2026-07-19\"", json);
    }
}
```

Update `tests/DocInt.Tests/ContractTestFactory.cs` — remove the xlsx fake so the real engine serves xlsx. Replace the `ConfigureFakes` body with:

```csharp
    protected override void ConfigureFakes(IServiceCollection services)
    {
        services.AddSingleton<IExtractionEngine>(new FakeEngine(
            [FileKind.Pdf, FileKind.Docx, FileKind.Pptx, FileKind.Html],
            f => FakeEngine.Markdown(f, "LAYOUT_MD_SENTINEL", 3)));
        services.AddSingleton<IExtractionEngine>(new FakeEngine(
            [FileKind.Image],
            f => FakeEngine.Image(f, "VISION_DESCRIPTION_SENTINEL")));
    }
```

- [ ] **Step 3: Run to verify failure**

```bash
dotnet build --no-restore src/DocInt.slnx
```

Expected: FAILS — `SpreadsheetEngine` missing.

- [ ] **Step 4: Implement the engine**

Add to `src/DocInt.Api/DocInt.Api.csproj` in the PackageReference ItemGroup:

```xml
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
```

Create `src/DocInt.Api/Engines/SpreadsheetEngine.cs`:

```csharp
using System.Globalization;
using System.Text;
using DocInt.Api.Contracts;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace DocInt.Api.Engines;

/// <summary>
/// XLSX via OpenXML — the deliberate precision path: typed cell values straight from the
/// stored XML (never re-parsed display text), plus a Markdown rendering of each sheet.
/// </summary>
public sealed class SpreadsheetEngine : IExtractionEngine
{
    public IReadOnlyCollection<FileKind> Kinds { get; } = [FileKind.Xlsx];

    public Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct) =>
        Task.FromResult(ExtractCore(file));

    private static EngineOutcome ExtractCore(FileItem file)
    {
        try
        {
            using var ms = new MemoryStream(file.Bytes, writable: false);
            using var doc = SpreadsheetDocument.Open(ms, isEditable: false);
            var workbookPart = doc.WorkbookPart
                ?? throw new InvalidDataException("workbook part missing");

            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
                .Elements<S.SharedStringItem>().Select(i => i.InnerText).ToArray() ?? [];
            var dateStyles = DateStyleIndexes(workbookPart.WorkbookStylesPart?.Stylesheet);

            var warnings = new List<string>(file.Warnings);
            var tables = new List<TableResult>();
            var sheets = workbookPart.Workbook.Sheets?.Elements<S.Sheet>().ToArray() ?? [];
            foreach (var sheet in sheets)
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
                var rows = ReadRows(worksheetPart, sharedStrings, dateStyles, warnings);
                tables.Add(new TableResult(sheet.Name?.Value ?? "Sheet", RenderMarkdown(rows), rows));
            }
            if (sheets.Length == 0) warnings.Add("workbook has no sheets");

            var markdown = string.Join("\n\n", tables.Select(t => $"## {t.Name}\n\n{t.Markdown}"));
            return new EngineOutcome(
                new FileResult(file.Name, file.Kind, markdown, tables, null, warnings, null),
                tables.Count);
        }
        catch (Exception ex) when (ex is DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
            or FileFormatException or InvalidDataException)
        {
            return Errors.For(file, ErrorCodes.Corrupt, $"file is not a readable XLSX workbook: {ex.Message}");
        }
    }

    private static HashSet<uint> DateStyleIndexes(S.Stylesheet? stylesheet)
    {
        var result = new HashSet<uint>();
        var formats = stylesheet?.CellFormats?.Elements<S.CellFormat>().ToArray() ?? [];
        for (var i = 0u; i < formats.Length; i++)
        {
            var id = formats[i].NumberFormatId?.Value ?? 0;
            if (id is >= 14 and <= 22 or >= 45 and <= 47) result.Add(i);
        }
        return result;
    }

    private static List<IReadOnlyList<object?>> ReadRows(
        WorksheetPart worksheetPart, string[] sharedStrings, HashSet<uint> dateStyles, List<string> warnings)
    {
        var grid = new List<IReadOnlyList<object?>>();
        var sheetData = worksheetPart.Worksheet.GetFirstChild<S.SheetData>();
        if (sheetData is null) return grid;

        var maxColumns = 0;
        var rawRows = new List<(uint Index, Dictionary<int, object?> Cells)>();
        foreach (var row in sheetData.Elements<S.Row>())
        {
            var cells = new Dictionary<int, object?>();
            var fallbackColumn = 0;
            foreach (var cell in row.Elements<S.Cell>())
            {
                var column = cell.CellReference?.Value is { } reference
                    ? ColumnIndex(reference) : fallbackColumn;
                fallbackColumn = column + 1;
                cells[column] = CellValue(cell, sharedStrings, dateStyles, warnings);
                maxColumns = Math.Max(maxColumns, column + 1);
            }
            rawRows.Add((row.RowIndex?.Value ?? (uint)(rawRows.Count + 1), cells));
        }

        foreach (var (_, cells) in rawRows)
        {
            var materialized = new object?[maxColumns];
            foreach (var (column, value) in cells) materialized[column] = value;
            grid.Add(materialized);
        }
        // Trim trailing all-null rows.
        while (grid.Count > 0 && grid[^1].All(v => v is null)) grid.RemoveAt(grid.Count - 1);
        return grid;
    }

    private static object? CellValue(
        S.Cell cell, string[] sharedStrings, HashSet<uint> dateStyles, List<string> warnings)
    {
        var raw = cell.DataType?.Value == S.CellValues.InlineString
            ? cell.InlineString?.InnerText
            : cell.CellValue?.InnerText;

        if (raw is null)
        {
            if (cell.CellFormula is not null)
                warnings.Add($"formula cell {cell.CellReference?.Value} has no cached value");
            return null;
        }

        var dataType = cell.DataType?.Value;
        if (dataType == S.CellValues.SharedString)
            return int.TryParse(raw, out var i) && i < sharedStrings.Length ? sharedStrings[i] : raw;
        if (dataType == S.CellValues.Boolean)
            return raw == "1";
        if (dataType == S.CellValues.String || dataType == S.CellValues.InlineString)
            return raw;
        if (dataType == S.CellValues.Error)
        {
            warnings.Add($"cell {cell.CellReference?.Value} contains error '{raw}'");
            return null;
        }
        if (dataType == S.CellValues.Date)
            return Iso(DateTime.Parse(raw, CultureInfo.InvariantCulture));

        // Numeric (no DataType): date-styled → ISO string, otherwise decimal (double on overflow).
        if (cell.StyleIndex?.Value is { } style && dateStyles.Contains(style))
            return Iso(DateTime.FromOADate(double.Parse(raw, CultureInfo.InvariantCulture)));
        try
        {
            return decimal.Parse(raw, NumberStyles.Float, CultureInfo.InvariantCulture);
        }
        catch (OverflowException)
        {
            return double.Parse(raw, CultureInfo.InvariantCulture);
        }
    }

    private static string Iso(DateTime value) =>
        value.TimeOfDay == TimeSpan.Zero
            ? value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : value.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);

    private static int ColumnIndex(string cellReference)
    {
        var index = 0;
        foreach (var c in cellReference)
        {
            if (!char.IsAsciiLetterUpper(c) && !char.IsAsciiLetterLower(c)) break;
            index = index * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }
        return index - 1;
    }

    private static string RenderMarkdown(List<IReadOnlyList<object?>> rows)
    {
        if (rows.Count == 0) return "";
        var sb = new StringBuilder();
        AppendRow(sb, rows[0]);
        sb.Append('|');
        for (var i = 0; i < rows[0].Count; i++) sb.Append(" --- |");
        sb.Append('\n');
        foreach (var row in rows.Skip(1)) AppendRow(sb, row);
        return sb.ToString().TrimEnd('\n');

        static void AppendRow(StringBuilder sb, IReadOnlyList<object?> row)
        {
            sb.Append('|');
            foreach (var cell in row)
                sb.Append(' ').Append(Render(cell)).Append(" |");
            sb.Append('\n');
        }

        static string Render(object? value) => value switch
        {
            null => "",
            bool b => b ? "true" : "false",
            decimal d => d.ToString(CultureInfo.InvariantCulture),
            double d => d.ToString(CultureInfo.InvariantCulture),
            string s => s.Replace("|", "\\|"),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""
        };
    }
}
```

In `src/DocInt.Api/Program.cs`, insert after `builder.Services.AddSingleton<ExtractionService>();`:

```csharp
    builder.Services.AddSingleton<IExtractionEngine, SpreadsheetEngine>();
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (4 new; all previous still green — the xlsx contract path now runs the real engine).

- [ ] **Step 6: Gate, commit, merge**

```bash
git add -A && git commit -m "Step 6: SpreadsheetEngine — OpenXML typed cells with numeric fidelity + Markdown"
git checkout main && git merge step-06-spreadsheet && git branch -d step-06-spreadsheet && git push origin main
```

---

### Task 7: LayoutEngine + Document Intelligence adapter

**Files:**
- Create: `src/DocInt.Api/Engines/ILayoutAnalysisClient.cs`, `src/DocInt.Api/Engines/AzureLayoutAnalysisClient.cs`, `src/DocInt.Api/Engines/LayoutEngine.cs`
- Modify: `src/DocInt.Api/DocInt.Api.csproj` (Azure packages), `src/DocInt.Api/Program.cs`, `tests/DocInt.Tests/ContractTestFactory.cs` (fake adapter instead of fake engine)
- Test: `tests/DocInt.Tests/LayoutEngineTests.cs`, update `tests/DocInt.Tests/ExtractContractTests.cs`

**Interfaces:**
- Consumes: seam types from Task 4, options from Task 2.
- Produces: `LayoutAnalysis(string Markdown, int PageCount, IReadOnlyList<string> Warnings)`; `ILayoutAnalysisClient { bool IsConfigured; Task<LayoutAnalysis> AnalyzeAsync(BinaryData content, CancellationToken ct) }`; `AzureLayoutAnalysisClient` (key or `DefaultAzureCredential`; throws `EngineUnconfiguredException` if used unconfigured); `LayoutEngine : IExtractionEngine` (Kinds = [Pdf, Docx, Pptx, Html]; unconfigured → `engine_unconfigured`; `RequestFailedException` with status 400 → `corrupt`). Test-side: `FakeLayoutClient` in `ContractTestFactory.cs`.

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-07-layout
```

- [ ] **Step 2: Write the failing tests**

Create `tests/DocInt.Tests/LayoutEngineTests.cs`:

```csharp
using Azure;
using DocInt.Api.Contracts;
using DocInt.Api.Engines;

namespace DocInt.Tests;

public class LayoutEngineTests
{
    private static FileItem Pdf() => new()
    { Index = 0, Name = "manual.pdf", Kind = FileKind.Pdf, Bytes = TestBytes.Pdf };

    [Fact]
    public async Task Returns_markdown_pages_and_merged_warnings()
    {
        var engine = new LayoutEngine(new FakeLayoutClient(
            new LayoutAnalysis("# extracted", 5, ["di: low confidence region"])));
        var item = Pdf();
        item.Warnings.Add("claimed content type mismatch");
        var outcome = await engine.ExtractAsync(item, CancellationToken.None);

        Assert.Equal("# extracted", outcome.Result.Markdown);
        Assert.Equal(5, outcome.PagesProcessed);
        Assert.Contains("claimed content type mismatch", outcome.Result.Warnings);
        Assert.Contains("di: low confidence region", outcome.Result.Warnings);
        Assert.Null(outcome.Result.Error);
    }

    [Fact]
    public async Task Unconfigured_client_yields_engine_unconfigured()
    {
        var engine = new LayoutEngine(new FakeLayoutClient(configured: false));
        var outcome = await engine.ExtractAsync(Pdf(), CancellationToken.None);
        Assert.Equal(ErrorCodes.EngineUnconfigured, outcome.Result.Error!.Code);
    }

    [Fact]
    public async Task Service_content_rejection_maps_to_corrupt()
    {
        var engine = new LayoutEngine(new FakeLayoutClient(
            thrower: () => throw new RequestFailedException(400, "InvalidContent: file is damaged")));
        var outcome = await engine.ExtractAsync(Pdf(), CancellationToken.None);
        Assert.Equal(ErrorCodes.Corrupt, outcome.Result.Error!.Code);
    }
}
```

Update `tests/DocInt.Tests/ContractTestFactory.cs` — replace the whole file (the layout fake moves from engine level to adapter level; `FakeLayoutClient` is shared with `LayoutEngineTests`):

```csharp
using DocInt.Api.Contracts;
using DocInt.Api.Engines;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace DocInt.Tests;

public sealed class FakeEngine(IReadOnlyCollection<FileKind> kinds, Func<FileItem, EngineOutcome> produce)
    : IExtractionEngine
{
    public IReadOnlyCollection<FileKind> Kinds => kinds;

    public Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct) =>
        Task.FromResult(produce(file));

    public static EngineOutcome Markdown(FileItem f, string markdown, int pages) =>
        new(new FileResult(f.Name, f.Kind, markdown, null, null, f.Warnings.ToArray(), null), pages);

    public static EngineOutcome Image(FileItem f, string description) =>
        new(new FileResult(f.Name, f.Kind, null, null, description, f.Warnings.ToArray(), null), 1);
}

public sealed class FakeLayoutClient(
    LayoutAnalysis? result = null, bool configured = true, Action? thrower = null) : ILayoutAnalysisClient
{
    public bool IsConfigured => configured;

    public Task<LayoutAnalysis> AnalyzeAsync(BinaryData content, CancellationToken ct)
    {
        thrower?.Invoke();
        return Task.FromResult(result ?? new LayoutAnalysis("LAYOUT_MD_SENTINEL", 3, []));
    }
}

public class ContractTestFactory : DocIntAppFactory
{
    protected override void ConfigureFakes(IServiceCollection services)
    {
        services.RemoveAll<ILayoutAnalysisClient>();
        services.AddSingleton<ILayoutAnalysisClient>(new FakeLayoutClient());
        services.AddSingleton<IExtractionEngine>(new FakeEngine(
            [FileKind.Image],
            f => FakeEngine.Image(f, "VISION_DESCRIPTION_SENTINEL")));
    }
}
```

In `tests/DocInt.Tests/ExtractContractTests.cs`, replace the whole `Kind_without_registered_engine_gets_engine_error` test (all engines are registered from now on; the interesting unconfigured path replaces it):

```csharp
    [Fact]
    public async Task Unconfigured_layout_engine_yields_per_file_engine_unconfigured()
    {
        using var bare = new DocIntAppFactory();   // real adapters, no Azure config
        using var form = Multipart.Form(("manual.pdf", TestBytes.Pdf, "application/pdf"));
        var result = await ReadResponse(await bare.CreateClient().PostAsync("/v1/extract", form));
        Assert.Equal(ErrorCodes.EngineUnconfigured, result.Files[0].Error!.Code);
    }
```

and delete the now-unused `PdfOnlyEngineFactory` class.

- [ ] **Step 3: Run to verify compile failure**

```bash
dotnet build --no-restore src/DocInt.slnx
```

Expected: FAILS — `ILayoutAnalysisClient`, `LayoutAnalysis`, `LayoutEngine` missing.

- [ ] **Step 4: Implement adapter + engine**

Add to `src/DocInt.Api/DocInt.Api.csproj` PackageReference ItemGroup:

```xml
    <PackageReference Include="Azure.AI.DocumentIntelligence" Version="1.0.0" />
    <PackageReference Include="Azure.Identity" Version="1.14.0" />
```

Create `src/DocInt.Api/Engines/ILayoutAnalysisClient.cs`:

```csharp
namespace DocInt.Api.Engines;

public sealed record LayoutAnalysis(string Markdown, int PageCount, IReadOnlyList<string> Warnings);

/// <summary>Thin seam over Azure Document Intelligence — the test fake boundary and credential boundary.</summary>
public interface ILayoutAnalysisClient
{
    bool IsConfigured { get; }
    Task<LayoutAnalysis> AnalyzeAsync(BinaryData content, CancellationToken ct);
}
```

Create `src/DocInt.Api/Engines/AzureLayoutAnalysisClient.cs`:

```csharp
using Azure;
using Azure.AI.DocumentIntelligence;
using Azure.Identity;
using DocInt.Api.Configuration;
using Microsoft.Extensions.Options;

namespace DocInt.Api.Engines;

public sealed class AzureLayoutAnalysisClient : ILayoutAnalysisClient
{
    private readonly DocumentIntelligenceClient? _client;

    public AzureLayoutAnalysisClient(IOptions<DocumentIntelligenceOptions> options)
    {
        var o = options.Value;
        if (string.IsNullOrWhiteSpace(o.Endpoint)) return;
        var endpoint = new Uri(o.Endpoint);
        _client = string.IsNullOrWhiteSpace(o.ApiKey)
            ? new DocumentIntelligenceClient(endpoint, new DefaultAzureCredential())
            : new DocumentIntelligenceClient(endpoint, new AzureKeyCredential(o.ApiKey));
    }

    public bool IsConfigured => _client is not null;

    public async Task<LayoutAnalysis> AnalyzeAsync(BinaryData content, CancellationToken ct)
    {
        if (_client is null)
            throw new EngineUnconfiguredException("DocumentIntelligence:Endpoint is not configured");
        var options = new AnalyzeDocumentOptions("prebuilt-layout", content)
        {
            OutputContentFormat = DocumentContentFormat.Markdown
        };
        var operation = await _client.AnalyzeDocumentAsync(WaitUntil.Completed, options, ct);
        var result = operation.Value;
        return new LayoutAnalysis(
            result.Content ?? "",
            result.Pages?.Count ?? 0,
            result.Warnings?.Select(w => $"{w.Code}: {w.Message}").ToArray() ?? []);
    }
}
```

Create `src/DocInt.Api/Engines/LayoutEngine.cs`:

```csharp
using Azure;
using DocInt.Api.Contracts;

namespace DocInt.Api.Engines;

/// <summary>PDF (incl. scanned), DOCX, PPTX, HTML → DI prebuilt-layout → Markdown.</summary>
public sealed class LayoutEngine(ILayoutAnalysisClient client) : IExtractionEngine
{
    public IReadOnlyCollection<FileKind> Kinds { get; } =
        [FileKind.Pdf, FileKind.Docx, FileKind.Pptx, FileKind.Html];

    public async Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct)
    {
        if (!client.IsConfigured)
            return Errors.For(file, ErrorCodes.EngineUnconfigured,
                "document layout engine is not configured: set DocumentIntelligence:Endpoint");
        try
        {
            var analysis = await client.AnalyzeAsync(BinaryData.FromBytes(file.Bytes), ct);
            var warnings = file.Warnings.Concat(analysis.Warnings).ToArray();
            return new EngineOutcome(
                new FileResult(file.Name, file.Kind, analysis.Markdown, null, null, warnings, null),
                analysis.PageCount);
        }
        catch (RequestFailedException ex) when (ex.Status == 400)
        {
            return Errors.For(file, ErrorCodes.Corrupt,
                $"document intelligence rejected the file: {ex.ErrorCode ?? ex.Message}");
        }
    }
}
```

In `src/DocInt.Api/Program.cs`, insert after `builder.Services.AddSingleton<IExtractionEngine, SpreadsheetEngine>();`:

```csharp
    builder.Services.AddSingleton<ILayoutAnalysisClient, AzureLayoutAnalysisClient>();
    builder.Services.AddSingleton<IExtractionEngine, LayoutEngine>();
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (3 new + updated contract tests; pdf/docx/pptx/html contract paths now run the real `LayoutEngine` over `FakeLayoutClient`).

- [ ] **Step 6: Gate, commit, merge**

```bash
git add -A && git commit -m "Step 7: LayoutEngine over Azure Document Intelligence adapter (prebuilt-layout, Markdown)"
git checkout main && git merge step-07-layout && git branch -d step-07-layout && git push origin main
```

---

### Task 8: VisionEngine + Azure OpenAI adapter (factual-observations prompt)

**Files:**
- Create: `src/DocInt.Api/Engines/IVisionChatClient.cs`, `src/DocInt.Api/Engines/AzureVisionChatClient.cs`, `src/DocInt.Api/Engines/VisionPrompt.cs`, `src/DocInt.Api/Engines/VisionEngine.cs`
- Modify: `src/DocInt.Api/DocInt.Api.csproj` (Azure.AI.OpenAI), `src/DocInt.Api/Program.cs`, `tests/DocInt.Tests/ContractTestFactory.cs` (fake vision adapter replaces the last fake engine)
- Test: `tests/DocInt.Tests/VisionEngineTests.cs`

**Interfaces:**
- Consumes: seam types, `AzureOpenAIOptions`, `FileItem.ImageMediaType`.
- Produces: `IVisionChatClient { bool IsConfigured; Task<string> DescribeImageAsync(string systemPrompt, BinaryData image, string mediaType, CancellationToken ct) }`; `AzureVisionChatClient`; `VisionPrompt.System` (const string — the guardrail, snapshot-tested); `VisionEngine : IExtractionEngine` (Kinds = [Image], PagesProcessed = 1). Test-side: `FakeVisionClient` recording `(systemPrompt, mediaType)`.

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-08-vision
```

- [ ] **Step 2: Write the failing tests**

Create `tests/DocInt.Tests/VisionEngineTests.cs`:

```csharp
using DocInt.Api.Contracts;
using DocInt.Api.Engines;

namespace DocInt.Tests;

public class VisionEngineTests
{
    private static FileItem Png() => new()
    { Index = 0, Name = "photo.png", Kind = FileKind.Image, ImageMediaType = "image/png", Bytes = TestBytes.Png };

    [Fact]
    public async Task Returns_description_with_pinned_prompt_and_media_type()
    {
        var fake = new FakeVisionClient("tinted lenses, UV400 sticker");
        var outcome = await new VisionEngine(fake).ExtractAsync(Png(), CancellationToken.None);

        Assert.Equal("tinted lenses, UV400 sticker", outcome.Result.ImageDescription);
        Assert.Equal(1, outcome.PagesProcessed);
        Assert.Equal(VisionPrompt.System, fake.LastSystemPrompt);
        Assert.Equal("image/png", fake.LastMediaType);
        Assert.Null(outcome.Result.Error);
    }

    [Fact]
    public async Task Unconfigured_client_yields_engine_unconfigured()
    {
        var outcome = await new VisionEngine(new FakeVisionClient(configured: false))
            .ExtractAsync(Png(), CancellationToken.None);
        Assert.Equal(ErrorCodes.EngineUnconfigured, outcome.Result.Error!.Code);
    }

    [Fact]
    public void Prompt_snapshot_pins_the_guardrail_wording()
    {
        // Deliberate double-entry bookkeeping: the guardrail cannot drift silently.
        Assert.Equal(
            "You describe product photographs for a document record. List only what is directly visible: "
            + "objects, text, markings, labels, symbols, materials, colors, quantities. Transcribe visible "
            + "text and codes exactly as printed. Do not identify product categories, do not infer purpose, "
            + "compliance status, quality, or anything not visible in the image. Output a plain-text description.",
            VisionPrompt.System);
    }
}
```

Update `tests/DocInt.Tests/ContractTestFactory.cs` — replace the `ContractTestFactory` class and add `FakeVisionClient` (file otherwise unchanged; `FakeEngine` stays for `ExtractionServiceTests`):

```csharp
public sealed class FakeVisionClient(
    string description = "VISION_DESCRIPTION_SENTINEL", bool configured = true) : IVisionChatClient
{
    public string? LastSystemPrompt { get; private set; }
    public string? LastMediaType { get; private set; }

    public bool IsConfigured => configured;

    public Task<string> DescribeImageAsync(string systemPrompt, BinaryData image, string mediaType, CancellationToken ct)
    {
        LastSystemPrompt = systemPrompt;
        LastMediaType = mediaType;
        return Task.FromResult(description);
    }
}

public class ContractTestFactory : DocIntAppFactory
{
    protected override void ConfigureFakes(IServiceCollection services)
    {
        services.RemoveAll<ILayoutAnalysisClient>();
        services.AddSingleton<ILayoutAnalysisClient>(new FakeLayoutClient());
        services.RemoveAll<IVisionChatClient>();
        services.AddSingleton<IVisionChatClient>(new FakeVisionClient());
    }
}
```

- [ ] **Step 3: Run to verify compile failure**

```bash
dotnet build --no-restore src/DocInt.slnx
```

Expected: FAILS — `IVisionChatClient`, `VisionEngine`, `VisionPrompt` missing.

- [ ] **Step 4: Implement prompt, adapter, engine**

Add to `src/DocInt.Api/DocInt.Api.csproj` PackageReference ItemGroup:

```xml
    <PackageReference Include="Azure.AI.OpenAI" Version="2.3.0" />
```

Create `src/DocInt.Api/Engines/VisionPrompt.cs`:

```csharp
namespace DocInt.Api.Engines;

/// <summary>
/// The factual-observations-only guardrail (Decision 12 open item 3). Docint states what
/// images SHOW, never what the product IS — classification lives upstream in EuGo-Web.
/// Changing this wording is a spec change: update the snapshot test AND the design doc.
/// </summary>
public static class VisionPrompt
{
    public const string System =
        "You describe product photographs for a document record. List only what is directly visible: "
        + "objects, text, markings, labels, symbols, materials, colors, quantities. Transcribe visible "
        + "text and codes exactly as printed. Do not identify product categories, do not infer purpose, "
        + "compliance status, quality, or anything not visible in the image. Output a plain-text description.";
}
```

Create `src/DocInt.Api/Engines/IVisionChatClient.cs`:

```csharp
namespace DocInt.Api.Engines;

/// <summary>Thin seam over the Azure OpenAI vision chat call.</summary>
public interface IVisionChatClient
{
    bool IsConfigured { get; }
    Task<string> DescribeImageAsync(string systemPrompt, BinaryData image, string mediaType, CancellationToken ct);
}
```

Create `src/DocInt.Api/Engines/AzureVisionChatClient.cs`:

```csharp
using Azure;
using Azure.AI.OpenAI;
using Azure.Identity;
using DocInt.Api.Configuration;
using Microsoft.Extensions.Options;
using OpenAI.Chat;

namespace DocInt.Api.Engines;

public sealed class AzureVisionChatClient : IVisionChatClient
{
    private readonly ChatClient? _chat;

    public AzureVisionChatClient(IOptions<AzureOpenAIOptions> options)
    {
        var o = options.Value;
        if (string.IsNullOrWhiteSpace(o.Endpoint)) return;
        var endpoint = new Uri(o.Endpoint);
        var azureClient = string.IsNullOrWhiteSpace(o.ApiKey)
            ? new AzureOpenAIClient(endpoint, new DefaultAzureCredential())
            : new AzureOpenAIClient(endpoint, new AzureKeyCredential(o.ApiKey));
        _chat = azureClient.GetChatClient(o.DeploymentNameVision);
    }

    public bool IsConfigured => _chat is not null;

    public async Task<string> DescribeImageAsync(
        string systemPrompt, BinaryData image, string mediaType, CancellationToken ct)
    {
        if (_chat is null)
            throw new EngineUnconfiguredException("AzureOpenAI:Endpoint is not configured");
        var completion = await _chat.CompleteChatAsync(
            [
                new SystemChatMessage(systemPrompt),
                new UserChatMessage(ChatMessageContentPart.CreateImagePart(image, mediaType))
            ],
            cancellationToken: ct);
        return completion.Value.Content[0].Text;
    }
}
```

Create `src/DocInt.Api/Engines/VisionEngine.cs`:

```csharp
using DocInt.Api.Contracts;

namespace DocInt.Api.Engines;

/// <summary>JPG/PNG → one chat completion with the factual-observations guardrail.</summary>
public sealed class VisionEngine(IVisionChatClient client) : IExtractionEngine
{
    public IReadOnlyCollection<FileKind> Kinds { get; } = [FileKind.Image];

    public async Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct)
    {
        if (!client.IsConfigured)
            return Errors.For(file, ErrorCodes.EngineUnconfigured,
                "vision engine is not configured: set AzureOpenAI:Endpoint");
        var description = await client.DescribeImageAsync(
            VisionPrompt.System,
            BinaryData.FromBytes(file.Bytes),
            file.ImageMediaType ?? "image/png",
            ct);
        return new EngineOutcome(
            new FileResult(file.Name, file.Kind, null, null, description, file.Warnings.ToArray(), null), 1);
    }
}
```

In `src/DocInt.Api/Program.cs`, insert after the LayoutEngine registrations:

```csharp
    builder.Services.AddSingleton<IVisionChatClient, AzureVisionChatClient>();
    builder.Services.AddSingleton<IExtractionEngine, VisionEngine>();
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (3 new; image contract path now runs real `VisionEngine` over `FakeVisionClient`). All three engines are now real; test fakes exist only at the two adapter seams.

- [ ] **Step 6: Gate, commit, merge**

```bash
git add -A && git commit -m "Step 8: VisionEngine over Azure OpenAI adapter with pinned factual-observations prompt"
git checkout main && git merge step-08-vision && git branch -d step-08-vision && git push origin main
```

---

### Task 9: Hardening — telemetry, pages metric, log redaction, request-size cap

**Files:**
- Create: `src/DocInt.Api/Telemetry/DocIntTelemetry.cs`
- Modify: `src/DocInt.Api/Engines/ExtractionService.cs`, `src/DocInt.Api/Program.cs`, `tests/DocInt.Tests/DocInt.Tests.csproj` (metrics-testing package)
- Test: `tests/DocInt.Tests/TelemetryTests.cs`, `tests/DocInt.Tests/RedactionTests.cs`, append one test to `tests/DocInt.Tests/MultipartReaderTests.cs`

**Interfaces:**
- Consumes: everything prior; golden `text.pdf` (whose body text `EuGo text PDF golden - Part M3 screw UV400` is the redaction probe).
- Produces: `DocIntTelemetry { const SourceName = "EuGo.DocInt"; const MeterName = "EuGo.DocInt"; const PagesProcessedInstrument = "docint.pages_processed"; ActivitySource ActivitySource; Counter<long> PagesProcessed }`; per-file activity `docint.extract_file` with tags `docint.kind`, `docint.size_bytes`, `docint.outcome` (no filenames in trace tags); Kestrel `MaxRequestBodySize` = `DocIntOptions.MaxRequestBytes`; reader rejects declared `Content-Length` above the cap with 400.

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-09-hardening
```

- [ ] **Step 2: Write the failing tests**

Add to `tests/DocInt.Tests/DocInt.Tests.csproj` PackageReference ItemGroup:

```xml
    <PackageReference Include="Microsoft.Extensions.Diagnostics.Testing" Version="10.7.0" />
```

Create `tests/DocInt.Tests/TelemetryTests.cs`:

```csharp
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
```

Create `tests/DocInt.Tests/RedactionTests.cs`:

```csharp
using System.Collections.Concurrent;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace DocInt.Tests;

/// <summary>Captures every formatted log line written through ILogger across all providers.</summary>
public sealed class CapturingLoggerProvider : ILoggerProvider
{
    public ConcurrentQueue<string> Lines { get; } = [];

    public ILogger CreateLogger(string categoryName) => new CapturingLogger(Lines);

    public void Dispose() { }

    private sealed class CapturingLogger(ConcurrentQueue<string> lines) : ILogger
    {
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
        public bool IsEnabled(LogLevel logLevel) => true;
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
            Exception? exception, Func<TState, Exception?, string> formatter) =>
            lines.Enqueue(formatter(state, exception));
    }
}

public class RedactionTests
{
    private sealed class CapturingFactory : ContractTestFactory
    {
        public CapturingLoggerProvider Capture { get; } = new();

        protected override void ConfigureFakes(IServiceCollection services)
        {
            base.ConfigureFakes(services);
            services.AddSingleton<ILoggerProvider>(Capture);
        }
    }

    [Fact]
    public async Task Logs_carry_filenames_but_never_document_content_or_extraction_output()
    {
        using var factory = new CapturingFactory();
        using var form = Multipart.Form(("text.pdf", Golden.Bytes("text.pdf"), "application/pdf"));
        var response = await factory.CreateClient().PostAsync("/v1/extract", form);
        response.EnsureSuccessStatusCode();

        var lines = factory.Capture.Lines.ToArray();
        Assert.Contains(lines, l => l.Contains("text.pdf"));
        Assert.DoesNotContain(lines, l => l.Contains("EuGo text PDF golden"));      // document content
        Assert.DoesNotContain(lines, l => l.Contains("LAYOUT_MD_SENTINEL"));        // extraction output
    }
}
```

The `Declared_content_length_over_request_cap_throws_bad_request` reader test originally scheduled here landed early, in Task 3b, alongside the `hints`-part size cap fix — not duplicated in this task.

- [ ] **Step 3: Run to verify failure**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx
```

Expected: FAILS — `DocInt.Api.Telemetry` missing. (The reader cap test would already pass — the reader has carried the Content-Length check since Task 3; it is asserted here for completeness.)

- [ ] **Step 4: Implement telemetry + wiring**

Create `src/DocInt.Api/Telemetry/DocIntTelemetry.cs`:

```csharp
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
```

Replace `src/DocInt.Api/Engines/ExtractionService.cs` with:

```csharp
using System.Diagnostics;
using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using DocInt.Api.Telemetry;
using Microsoft.Extensions.Options;

namespace DocInt.Api.Engines;

/// <summary>Runs a batch: bounded parallelism, request-order results, per-file span + metric + log line (no content).</summary>
public sealed class ExtractionService(
    EngineRouter router,
    IOptions<DocIntOptions> options,
    DocIntTelemetry telemetry,
    ILogger<ExtractionService> logger)
{
    public async Task<ExtractResponse> ExtractAsync(IReadOnlyList<FileItem> files, CancellationToken ct)
    {
        var results = new FileResult[files.Count];
        await Parallel.ForEachAsync(files,
            new ParallelOptions { MaxDegreeOfParallelism = options.Value.MaxParallelism, CancellationToken = ct },
            async (file, token) =>
            {
                var kindName = file.Kind?.Name() ?? "unknown";
                using var activity = telemetry.ActivitySource.StartActivity("docint.extract_file");
                activity?.SetTag("docint.kind", kindName);
                activity?.SetTag("docint.size_bytes", file.Bytes.Length);

                var started = Stopwatch.GetTimestamp();
                var outcome = file.Error is not null
                    ? new EngineOutcome(new FileResult(file.Name, file.Kind, null, null, null,
                        file.Warnings.ToArray(), file.Error), 0)
                    : await router.RouteAsync(file, token);
                results[file.Index] = outcome.Result;

                var outcomeCode = outcome.Result.Error?.Code ?? "ok";
                activity?.SetTag("docint.outcome", outcomeCode);
                if (outcome.PagesProcessed > 0)
                    telemetry.PagesProcessed.Add(outcome.PagesProcessed,
                        new KeyValuePair<string, object?>("kind", kindName));
                logger.LogInformation(
                    "Processed {FileName}: kind={Kind} sizeBytes={SizeBytes} outcome={Outcome} durationMs={DurationMs:0}",
                    file.Name, kindName, file.Bytes.Length, outcomeCode,
                    Stopwatch.GetElapsedTime(started).TotalMilliseconds);
            });
        return new ExtractResponse(results);
    }
}
```

In `src/DocInt.Api/Program.cs`:
1. Add usings: `using DocInt.Api.Telemetry;`, `using OpenTelemetry.Metrics;`, `using OpenTelemetry.Trace;`
2. Insert after `builder.AddDocIntOptions();`:

```csharp
    builder.Services.AddSingleton<DocIntTelemetry>();
    builder.Services.ConfigureOpenTelemetryTracerProvider(t => t.AddSource(DocIntTelemetry.SourceName));
    builder.Services.ConfigureOpenTelemetryMeterProvider(m => m.AddMeter(DocIntTelemetry.MeterName));
    builder.WebHost.ConfigureKestrel((context, kestrel) =>
    {
        var docint = new DocIntOptions();
        context.Configuration.GetSection(DocIntOptions.SectionName).Bind(docint);
        kestrel.Limits.MaxRequestBodySize = docint.MaxRequestBytes;
    });
```

(The reader's declared-Content-Length check from Task 3 is the in-memory-testable guard; the Kestrel limit is the belt for the real server, which `WebApplicationFactory`'s TestServer does not exercise.)

The new `DocIntTelemetry` constructor parameter breaks `ExtractionServiceTests` (Task 4), which constructs `ExtractionService` directly. In `tests/DocInt.Tests/ExtractionServiceTests.cs`, add a helper and use it in both tests:

```csharp
    private static DocInt.Api.Telemetry.DocIntTelemetry TestTelemetry()
    {
        var provider = new ServiceCollection().AddMetrics().BuildServiceProvider();
        return new DocInt.Api.Telemetry.DocIntTelemetry(
            provider.GetRequiredService<System.Diagnostics.Metrics.IMeterFactory>());
    }
```

with `using Microsoft.Extensions.DependencyInjection;` added to the file, and change both `new ExtractionService(new EngineRouter(...), options, NullLogger<ExtractionService>.Instance)` calls to `new ExtractionService(new EngineRouter(...), options, TestTelemetry(), NullLogger<ExtractionService>.Instance)`.

- [ ] **Step 5: Run tests to verify they pass**

```bash
dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
```

Expected: PASS (3 new).

- [ ] **Step 6: Gate, commit, merge**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
git add -A && git commit -m "Step 9: per-file spans, pages_processed metric, log-redaction test, request-size caps"
git checkout main && git merge step-09-hardening && git branch -d step-09-hardening && git push origin main
```

---

### Task 10: Env-gated live smoke suite

**Files:**
- Modify: `tests/DocInt.Tests/DocInt.Tests.csproj` (SkippableFact package), `CLAUDE.md` (live-smoke section)
- Test: `tests/DocInt.Tests/LiveSmokeTests.cs`

**Interfaces:**
- Consumes: all golden fixtures; the real `Program` with real Azure adapters (a plain `DocIntAppFactory` — no fakes — reading config from process env vars `DocumentIntelligence__Endpoint`, `DocumentIntelligence__ApiKey`, `AzureOpenAI__Endpoint`, `AzureOpenAI__ApiKey`).
- Produces: live tests that run only when `DOCINT_LIVE_TESTS=1` **and** the relevant endpoint env var is set; otherwise reported as SKIPPED (never failed, never silently green).

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-10-livesmoke
```

- [ ] **Step 2: Add the suite**

Add to `tests/DocInt.Tests/DocInt.Tests.csproj` PackageReference ItemGroup:

```xml
    <PackageReference Include="Xunit.SkippableFact" Version="1.5.23" />
```

Create `tests/DocInt.Tests/LiveSmokeTests.cs`:

```csharp
using System.Text.Json;
using DocInt.Api.Contracts;

namespace DocInt.Tests;

/// <summary>
/// Live smoke against real Azure. Gated: set DOCINT_LIVE_TESTS=1 plus
/// DocumentIntelligence__Endpoint (+ ApiKey unless using az login) and
/// AzureOpenAI__Endpoint (+ ApiKey) in the environment, then run
///   DOCINT_LIVE_TESTS=1 dotnet test --no-build src/DocInt.slnx --filter "FullyQualifiedName~LiveSmokeTests"
/// Without the gate every test here reports SKIPPED.
/// </summary>
public class LiveSmokeTests : IClassFixture<DocIntAppFactory>
{
    private static bool LiveEnabled =>
        Environment.GetEnvironmentVariable("DOCINT_LIVE_TESTS") == "1";
    private static bool HasDi =>
        !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("DocumentIntelligence__Endpoint"));
    private static bool HasVision =>
        !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("AzureOpenAI__Endpoint"));

    private readonly DocIntAppFactory _factory;

    public LiveSmokeTests(DocIntAppFactory factory) => _factory = factory;

    private async Task<FileResult> ExtractOne(string fixture, string contentType)
    {
        var client = _factory.CreateClient();
        client.Timeout = TimeSpan.FromMinutes(5);
        using var form = Multipart.Form((fixture, Golden.Bytes(fixture), contentType));
        var response = await client.PostAsync("/v1/extract", form);
        response.EnsureSuccessStatusCode();
        var envelope = JsonSerializer.Deserialize<ExtractResponse>(
            await response.Content.ReadAsStringAsync(), DocIntJson.Options)!;
        return envelope.Files[0];
    }

    [SkippableTheory]
    [InlineData("text.pdf", "application/pdf")]
    [InlineData("scanned.pdf", "application/pdf")]
    [InlineData("sample.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
    [InlineData("sample.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation")]
    [InlineData("sample.html", "text/html")]
    public async Task Layout_kinds_yield_nonempty_markdown(string fixture, string contentType)
    {
        Skip.IfNot(LiveEnabled && HasDi, "live tests disabled or DocumentIntelligence__Endpoint not set");
        var file = await ExtractOne(fixture, contentType);
        Assert.Null(file.Error);
        Assert.False(string.IsNullOrWhiteSpace(file.Markdown));
    }

    [SkippableFact]
    public async Task Scanned_pdf_proves_ocr()
    {
        Skip.IfNot(LiveEnabled && HasDi, "live tests disabled or DocumentIntelligence__Endpoint not set");
        var file = await ExtractOne("scanned.pdf", "application/pdf");
        Assert.Contains("OCR", file.Markdown, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public async Task Photo_description_mentions_lens_or_uv_cues()
    {
        Skip.IfNot(LiveEnabled && HasVision, "live tests disabled or AzureOpenAI__Endpoint not set");
        var file = await ExtractOne("photo.png", "image/png");
        Assert.Null(file.Error);
        Assert.False(string.IsNullOrWhiteSpace(file.ImageDescription));
        var description = file.ImageDescription!.ToLowerInvariant();
        Assert.True(description.Contains("uv") || description.Contains("lens") || description.Contains("sunglass"),
            $"description lacks lens/UV cues: {description}");
    }
}
```

- [ ] **Step 3: Verify skip behavior locally (no Azure configured)**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx
dotnet test --no-build src/DocInt.slnx --filter "FullyQualifiedName~LiveSmokeTests"
```

Expected: 7 tests reported **Skipped**, 0 failed.

- [ ] **Step 4: Document the gate in CLAUDE.md**

Append to `CLAUDE.md` under the Conventions section:

```markdown

## Live smoke tests

Azure-stubbed tests are the default; the live suite (`LiveSmokeTests`) self-skips unless gated on:

```bash
export DOCINT_LIVE_TESTS=1
export DocumentIntelligence__Endpoint=https://<di-resource>.cognitiveservices.azure.com/
export DocumentIntelligence__ApiKey=<key>        # omit to use DefaultAzureCredential (az login)
export AzureOpenAI__Endpoint=https://<aoai-resource>.openai.azure.com/
export AzureOpenAI__ApiKey=<key>                 # omit to use DefaultAzureCredential
dotnet test --no-build src/DocInt.slnx --filter "FullyQualifiedName~LiveSmokeTests"
```

Golden fixtures are committed binaries; regenerate only deliberately with `dotnet run --project tools/make-golden`.
```

- [ ] **Step 5: Gate, commit, merge**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
git add -A && git commit -m "Step 10: env-gated live smoke suite (DI kinds, OCR proof, vision cues)"
git checkout main && git merge step-10-livesmoke && git branch -d step-10-livesmoke && git push origin main
```

---

### Task 11: Multi-arch Dockerfile, CI docker job, README

**Files:**
- Create: `Dockerfile`, `README.md`
- Modify: `.github/workflows/ci.yml` (add docker job)

**Interfaces:**
- Consumes: the finished service; `.dockerignore` (already present).
- Produces: image building for `linux/amd64` + `linux/arm64` (user directive 2026-07-19); CI validates both platforms with `push: false` (ACR push is EuGo-infra's CD job).

- [ ] **Step 1: Branch**

```bash
git checkout main && git pull && git checkout -b step-11-delivery
```

- [ ] **Step 2: Dockerfile (cross-publish pattern — no QEMU-emulated compilation)**

Create `Dockerfile` at the repo root (build context = repo root):

```dockerfile
FROM mcr.microsoft.com/dotnet/aspnet:10.0-noble-chiseled-extra AS base
USER $APP_UID
WORKDIR /app
EXPOSE 8090

# SDK stage runs on the BUILD host's architecture and cross-publishes for TARGETARCH;
# only the (multi-arch) runtime base resolves per target platform.
FROM --platform=$BUILDPLATFORM mcr.microsoft.com/dotnet/sdk:10.0 AS build
ARG TARGETARCH
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["src/DocInt.Api/DocInt.Api.csproj", "DocInt.Api/"]
COPY ["src/ServiceDefaults/ServiceDefaults.csproj", "ServiceDefaults/"]
RUN dotnet restore "DocInt.Api/DocInt.Api.csproj" -a $TARGETARCH
COPY src/DocInt.Api/ DocInt.Api/
COPY src/ServiceDefaults/ ServiceDefaults/
RUN dotnet publish "DocInt.Api/DocInt.Api.csproj" -c $BUILD_CONFIGURATION -a $TARGETARCH --no-restore \
    -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "DocInt.Api.dll"]
```

- [ ] **Step 3: Verify the container builds and answers**

```bash
docker build -t eugo-docint:local .
docker run -d --rm -p 18090:8090 --name docint-smoke eugo-docint:local
sleep 5 && curl -s http://localhost:18090/healthz && curl -s http://localhost:18090/info
docker stop docint-smoke
```

Expected: `Healthy` and the `/info` JSON. If buildx + binfmt are available locally, optionally verify arm64 too: `docker buildx build --platform linux/arm64 -t eugo-docint:arm64-check .` — otherwise CI is the arm64 authority.

- [ ] **Step 4: Add the multi-arch docker job to CI**

In `.github/workflows/ci.yml`, append after the `build-test` job (same `jobs:` map):

```yaml
  docker:
    runs-on: ubuntu-latest
    needs: build-test
    steps:
      - uses: actions/checkout@v4
      - uses: docker/setup-qemu-action@v3
      - uses: docker/setup-buildx-action@v3
      - uses: docker/build-push-action@v6
        with:
          context: .
          platforms: linux/amd64,linux/arm64
          push: false
          tags: eugo-docint:ci
```

- [ ] **Step 5: README**

Create `README.md`:

```markdown
# EuGo-docint

Stateless, cluster-internal document-understanding service for the EuGo platform:
files in → Markdown, tables with typed numeric cells, image descriptions, per-file
warnings/errors out. Document-level language only — no compliance semantics, no storage,
no document content in logs.

- Design (authoritative): [docs/superpowers/specs/2026-07-19-eugo-docint-design.md](docs/superpowers/specs/2026-07-19-eugo-docint-design.md)
- Conventions & commands: [CLAUDE.md](CLAUDE.md)

## API

`POST /v1/extract` — `multipart/form-data`: N file parts named `files` (pdf/docx/pptx/html/xlsx/jpg/png),
optional `hints` part `{"<filename>":{"purpose":"bom|photo"}}`. Well-formed requests always return
`200` with per-file success or error. Also: `GET /healthz`, `GET /info`, OpenAPI JSON in Development.

## Run

```bash
dotnet run --project src/AppHost      # Aspire dashboard + telemetry (dev)
dotnet run --project src/DocInt.Api   # plain service on http://localhost:8090
```

Azure engines activate when configured (user-secrets or env; endpoint without key → DefaultAzureCredential):
`DocumentIntelligence:Endpoint`, `AzureOpenAI:Endpoint`, `AzureOpenAI:DeploymentNameVision` (default `gpt-4.1-mini`).
Unconfigured engines answer with per-file `engine_unconfigured` — the service always boots.

## Test

```bash
dotnet restore src/DocInt.slnx
dotnet build --no-restore src/DocInt.slnx
dotnet test --no-build src/DocInt.slnx
```

Live smoke against real Azure is env-gated — see CLAUDE.md. Container: `docker build -t eugo-docint .`
(CI builds linux/amd64 + linux/arm64). Kubernetes deployment lives in the EuGo-infra repo.
```

- [ ] **Step 6: Gate, commit, merge, verify CI**

```bash
dotnet restore src/DocInt.slnx && dotnet build --no-restore src/DocInt.slnx && dotnet test --no-build src/DocInt.slnx
git add -A && git commit -m "Step 11: multi-arch Dockerfile (amd64/arm64), CI docker job, README"
git checkout main && git merge step-11-delivery && git branch -d step-11-delivery && git push origin main
gh run watch --exit-status || gh run list --limit 3
```

Expected: both CI jobs green (`build-test`, then `docker` building both platforms). If the docker job fails on an action version, bump the failing action to its latest major and re-push.

---

## Completion checklist (maps back to the spec)

- [ ] `/v1/extract` frozen: multipart + hints, per-file success/failure inside 200, request-level 400s only for malformed requests — Tasks 3–4.
- [ ] Engines: Layout (DI prebuilt-layout → Markdown), Spreadsheet (typed cells, numeric fidelity), Vision (pinned factual-observations prompt) — Tasks 6–8.
- [ ] Stub-first: service boots with zero Azure config; `engine_unconfigured` per file; live smoke lights up via env when EuGo-infra delivers endpoints — Tasks 7, 8, 10.
- [ ] Hardening: bounded parallelism, per-file timeout, `docint.pages_processed`, per-file spans, log-redaction test, size caps — Tasks 4, 9.
- [ ] Delivery: golden fixtures committed, CI build-test + multi-arch docker (amd64/arm64, push: false), README — Tasks 5, 11.
- [ ] Out of scope (deliberate): `deploy/` K8s manifests (EuGo-infra), EuGo-Web client (T8, lives in EuGo-web).
