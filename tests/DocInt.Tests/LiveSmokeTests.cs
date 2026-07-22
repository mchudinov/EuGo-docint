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
