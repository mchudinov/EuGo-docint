using System.Diagnostics;
using DocInt.Api.Configuration;
using DocInt.Api.Contracts;
using DocInt.Api.Engines;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DocInt.Tests;

public class ExtractionServiceTests
{
    private sealed class CountingEngine : IExtractionEngine
    {
        private int _inFlight;
        public int MaxObserved;
        public IReadOnlyCollection<FileKind> Kinds { get; } = [FileKind.Pdf];

        public async Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct)
        {
            var now = Interlocked.Increment(ref _inFlight);
            InterlockedExtensions.Max(ref MaxObserved, now);
            await Task.Delay(50, ct);
            Interlocked.Decrement(ref _inFlight);
            return FakeEngine.Markdown(file, $"md-{file.Index}", 1);
        }
    }

    private static class InterlockedExtensions
    {
        public static void Max(ref int location, int value)
        {
            int current;
            while (value > (current = Volatile.Read(ref location)))
                Interlocked.CompareExchange(ref location, value, current);
        }
    }

    [Fact]
    public async Task Respects_parallelism_cap_and_preserves_request_order()
    {
        var options = Microsoft.Extensions.Options.Options.Create(
            new DocIntOptions { MaxParallelism = 2 });
        var engine = new CountingEngine();
        var service = new ExtractionService(
            new EngineRouter([engine], options), options, NullLogger<ExtractionService>.Instance);

        var files = Enumerable.Range(0, 8).Select(i => new FileItem
        {
            Index = i, Name = $"f{i}.pdf", Kind = FileKind.Pdf, Bytes = TestBytes.Pdf
        }).ToArray();

        var response = await service.ExtractAsync(files, CancellationToken.None);

        Assert.Equal(8, response.Files.Count);
        Assert.All(Enumerable.Range(0, 8), i => Assert.Equal($"md-{i}", response.Files[i].Markdown));
        Assert.True(engine.MaxObserved <= 2, $"observed {engine.MaxObserved} concurrent extractions");
    }

    [Fact]
    public async Task Timeout_produces_per_file_timeout_error()
    {
        var options = Microsoft.Extensions.Options.Options.Create(
            new DocIntOptions { PerFileTimeoutSeconds = 1 });
        var hanging = new FakeHangingEngine();
        var service = new ExtractionService(
            new EngineRouter([hanging], options), options, NullLogger<ExtractionService>.Instance);

        var files = new[] { new FileItem { Index = 0, Name = "slow.pdf", Kind = FileKind.Pdf, Bytes = TestBytes.Pdf } };
        var response = await service.ExtractAsync(files, CancellationToken.None);

        Assert.Equal(ErrorCodes.Timeout, response.Files[0].Error!.Code);
    }

    private sealed class FakeHangingEngine : IExtractionEngine
    {
        public IReadOnlyCollection<FileKind> Kinds { get; } = [FileKind.Pdf];
        public async Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct)
        {
            await Task.Delay(TimeSpan.FromMinutes(5), ct);
            throw new UnreachableException();
        }
    }
}
