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
