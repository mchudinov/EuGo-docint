using DocInt.Api.Contracts;

namespace DocInt.Api.Engines;

public sealed class EngineUnconfiguredException(string message) : Exception(message);

public static class Errors
{
    public static EngineOutcome For(FileItem file, string code, string message) =>
        new(new FileResult(file.Name, file.Kind, null, null, null, file.Warnings.ToArray(),
            new FileError(code, message)), 0);
}
