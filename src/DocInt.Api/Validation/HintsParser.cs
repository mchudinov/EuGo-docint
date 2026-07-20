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
