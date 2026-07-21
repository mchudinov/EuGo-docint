using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using Microsoft.Extensions.Options;

namespace DocInt.Api.Engines;

/// <summary>Picks the engine by detected kind; owns the per-file timeout and the catch-all.</summary>
public sealed class EngineRouter
{
    private readonly Dictionary<FileKind, IExtractionEngine> _byKind = [];
    private readonly TimeSpan _timeout;

    public EngineRouter(IEnumerable<IExtractionEngine> engines, IOptions<DocIntOptions> options)
    {
        foreach (var engine in engines)
            foreach (var kind in engine.Kinds)
                _byKind[kind] = engine;
        _timeout = TimeSpan.FromSeconds(options.Value.PerFileTimeoutSeconds);
    }

    public async Task<EngineOutcome> RouteAsync(FileItem file, CancellationToken requestCt)
    {
        var kind = file.Kind!.Value;
        if (!_byKind.TryGetValue(kind, out var engine))
            return Errors.For(file, ErrorCodes.EngineError, $"no engine registered for kind '{kind.Name()}'");

        using var cts = CancellationTokenSource.CreateLinkedTokenSource(requestCt);
        cts.CancelAfter(_timeout);
        try
        {
            return await engine.ExtractAsync(file, cts.Token);
        }
        catch (OperationCanceledException) when (cts.IsCancellationRequested && !requestCt.IsCancellationRequested)
        {
            return Errors.For(file, ErrorCodes.Timeout,
                $"extraction exceeded the per-file timeout of {_timeout.TotalSeconds:0}s");
        }
        catch (OperationCanceledException) when (requestCt.IsCancellationRequested)
        {
            // The request itself was cancelled — genuine client/host abandonment. There is
            // no client left to answer, so propagate instead of manufacturing a per-file error.
            throw;
        }
        catch (EngineUnconfiguredException ex)
        {
            return Errors.For(file, ErrorCodes.EngineUnconfigured, ex.Message);
        }
        catch (Exception ex)
        {
            // Any other exception, including a stray OperationCanceledException from an
            // engine's own mechanism (e.g. an SDK client's default HttpClient.Timeout) that
            // is tied to neither our per-file timeout nor the request token, lands here.
            // Exception message only — never document content.
            return Errors.For(file, ErrorCodes.EngineError, ex.Message);
        }
    }
}
