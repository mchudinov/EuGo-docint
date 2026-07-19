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
