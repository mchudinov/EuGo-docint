namespace DocInt.Api.Engines;

/// <summary>Thin seam over the Azure OpenAI vision chat call.</summary>
public interface IVisionChatClient
{
    bool IsConfigured { get; }
    Task<string> DescribeImageAsync(string systemPrompt, BinaryData image, string mediaType, CancellationToken ct);
}
