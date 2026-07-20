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
