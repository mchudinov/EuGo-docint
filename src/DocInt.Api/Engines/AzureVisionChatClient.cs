using Azure;
using Azure.AI.OpenAI;
using Azure.Identity;
using DocInt.Api.Configuration;
using Microsoft.Extensions.Options;
using OpenAI.Chat;

namespace DocInt.Api.Engines;

public sealed class AzureVisionChatClient : IVisionChatClient
{
    private readonly ChatClient? _chat;

    public AzureVisionChatClient(IOptions<AzureOpenAIOptions> options)
    {
        var o = options.Value;
        if (string.IsNullOrWhiteSpace(o.Endpoint)) return;
        var endpoint = new Uri(o.Endpoint);
        var azureClient = string.IsNullOrWhiteSpace(o.ApiKey)
            ? new AzureOpenAIClient(endpoint, new DefaultAzureCredential())
            : new AzureOpenAIClient(endpoint, new AzureKeyCredential(o.ApiKey));
        _chat = azureClient.GetChatClient(o.DeploymentNameVision);
    }

    public bool IsConfigured => _chat is not null;

    public async Task<string> DescribeImageAsync(
        string systemPrompt, BinaryData image, string mediaType, CancellationToken ct)
    {
        if (_chat is null)
            throw new EngineUnconfiguredException("AzureOpenAI:Endpoint is not configured");
        var completion = await _chat.CompleteChatAsync(
            [
                new SystemChatMessage(systemPrompt),
                new UserChatMessage(ChatMessageContentPart.CreateImagePart(image, mediaType))
            ],
            cancellationToken: ct);
        return completion.Value.Content[0].Text;
    }
}
