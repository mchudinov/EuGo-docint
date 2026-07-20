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
