using DocInt.Api.Contracts;

namespace DocInt.Api.Engines;

/// <summary>Result plus the pages/sheets/images processed (feeds the docint.pages_processed metric).</summary>
public sealed record EngineOutcome(FileResult Result, int PagesProcessed);

public interface IExtractionEngine
{
    IReadOnlyCollection<FileKind> Kinds { get; }
    Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct);
}
