namespace DocInt.Api.Engines;

public sealed record LayoutAnalysis(string Markdown, int PageCount, IReadOnlyList<string> Warnings);

/// <summary>Thin seam over Azure Document Intelligence — the test fake boundary and credential boundary.</summary>
public interface ILayoutAnalysisClient
{
    bool IsConfigured { get; }
    Task<LayoutAnalysis> AnalyzeAsync(BinaryData content, CancellationToken ct);
}
