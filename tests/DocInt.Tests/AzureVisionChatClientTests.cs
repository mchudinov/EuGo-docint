using DocInt.Api.Engines;
using OpenAI.Chat;

namespace DocInt.Tests;

public class AzureVisionChatClientTests
{
    // Regression: completion.Value.Content[0].Text used to throw ArgumentOutOfRangeException
    // straight out of the adapter when the model returns an empty/content-filtered response
    // (zero content parts). Exercised against a real ChatCompletion built via the SDK's own
    // OpenAIChatModelFactory — no network/transport mocking needed, and no reimplementation of
    // the guard under test: AzureVisionChatClient.ExtractText is the exact code DescribeImageAsync
    // calls.

    [Fact]
    public void ExtractText_throws_clean_error_when_content_is_empty()
    {
        var completion = OpenAIChatModelFactory.ChatCompletion(content: new ChatMessageContent());

        var ex = Assert.Throws<InvalidOperationException>(() => AzureVisionChatClient.ExtractText(completion));
        Assert.Equal("vision model returned no text content", ex.Message);
    }

    [Fact]
    public void ExtractText_returns_first_content_part_text_when_present()
    {
        var completion = OpenAIChatModelFactory.ChatCompletion(content: new ChatMessageContent("tinted lenses, UV400 sticker"));

        Assert.Equal("tinted lenses, UV400 sticker", AzureVisionChatClient.ExtractText(completion));
    }
}
