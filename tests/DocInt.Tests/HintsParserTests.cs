using DocInt.Api.Validation;

namespace DocInt.Tests;

public class HintsParserTests
{
    [Fact]
    public void Parses_filename_to_purpose_map()
    {
        var hints = HintsParser.Parse("""{"bom.xlsx":{"purpose":"bom"},"p.png":{"purpose":"photo"}}""");
        Assert.Equal("bom", hints["bom.xlsx"]);
        Assert.Equal("photo", hints["p.png"]);
    }

    [Fact]
    public void Entry_without_purpose_is_ignored()
    {
        var hints = HintsParser.Parse("""{"a.pdf":{}}""");
        Assert.Empty(hints);
    }

    [Fact]
    public void Invalid_json_throws_bad_request()
    {
        Assert.Throws<BadExtractRequestException>(() => HintsParser.Parse("not json"));
        Assert.Throws<BadExtractRequestException>(() => HintsParser.Parse("[1,2]"));
    }
}
