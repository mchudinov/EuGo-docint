using System.Diagnostics;
using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using Microsoft.Extensions.Options;

namespace DocInt.Api.Engines;

/// <summary>Runs a batch: bounded parallelism, request-order results, per-file log line (no content).</summary>
public sealed class ExtractionService(
    EngineRouter router,
    IOptions<DocIntOptions> options,
    ILogger<ExtractionService> logger)
{
    public async Task<ExtractResponse> ExtractAsync(IReadOnlyList<FileItem> files, CancellationToken ct)
    {
        var results = new FileResult[files.Count];
        await Parallel.ForEachAsync(files,
            new ParallelOptions { MaxDegreeOfParallelism = options.Value.MaxParallelism, CancellationToken = ct },
            async (file, token) =>
            {
                var started = Stopwatch.GetTimestamp();
                var outcome = file.Error is not null
                    ? new EngineOutcome(new FileResult(file.Name, file.Kind, null, null, null,
                        file.Warnings.ToArray(), file.Error), 0)
                    : await router.RouteAsync(file, token);
                results[file.Index] = outcome.Result;
                logger.LogInformation(
                    "Processed {FileName}: kind={Kind} sizeBytes={SizeBytes} outcome={Outcome} durationMs={DurationMs:0}",
                    file.Name, file.Kind?.Name() ?? "unknown", file.Bytes.Length,
                    outcome.Result.Error?.Code ?? "ok",
                    Stopwatch.GetElapsedTime(started).TotalMilliseconds);
            });
        return new ExtractResponse(results);
    }
}
