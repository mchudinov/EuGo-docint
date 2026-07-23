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
