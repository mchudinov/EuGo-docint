using DocInt.Api.Contracts;
using DocInt.Api.Validation;

namespace DocInt.Tests;

public class GoldenFixturesTests
{
    [Theory]
    [InlineData("text.pdf", FileKind.Pdf)]
    [InlineData("scanned.pdf", FileKind.Pdf)]
    [InlineData("sample.docx", FileKind.Docx)]
    [InlineData("sample.pptx", FileKind.Pptx)]
    [InlineData("sample.html", FileKind.Html)]
    [InlineData("bom.xlsx", FileKind.Xlsx)]
    [InlineData("photo.png", FileKind.Image)]
    [InlineData("corrupt.xlsx", FileKind.Xlsx)]
    public void Fixture_exists_and_detects_as_expected_kind(string name, FileKind expected)
    {
        var bytes = Golden.Bytes(name);
        Assert.NotEmpty(bytes);
        Assert.Equal(expected, FileKindDetector.Detect(name, null, bytes).Kind);
    }

    [Fact]
    public void Unknown_bin_is_undetectable()
    {
        Assert.Null(FileKindDetector.Detect("unknown.bin", null, Golden.Bytes("unknown.bin")).Kind);
    }
}
