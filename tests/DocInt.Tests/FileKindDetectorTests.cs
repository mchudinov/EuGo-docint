using DocInt.Api.Contracts;
using DocInt.Api.Validation;

namespace DocInt.Tests;

public class FileKindDetectorTests
{
    [Fact]
    public void Pdf_magic_wins_regardless_of_claims()
    {
        var r = FileKindDetector.Detect("doc.bin", "application/octet-stream", TestBytes.Pdf);
        Assert.Equal(FileKind.Pdf, r.Kind);
        Assert.Null(r.ImageMediaType);
    }

    [Fact]
    public void Png_and_jpeg_magic_detect_image_with_media_type()
    {
        Assert.Equal(("image/png", FileKind.Image),
            (FileKindDetector.Detect("a.png", null, TestBytes.Png).ImageMediaType,
             FileKindDetector.Detect("a.png", null, TestBytes.Png).Kind));
        Assert.Equal("image/jpeg", FileKindDetector.Detect("b.jpg", null, TestBytes.Jpeg).ImageMediaType);
    }

    [Fact]
    public void Zip_disambiguates_by_content_type_then_extension()
    {
        Assert.Equal(FileKind.Xlsx, FileKindDetector.Detect("f.bin",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", TestBytes.Zip).Kind);
        Assert.Equal(FileKind.Docx, FileKindDetector.Detect("f.docx", null, TestBytes.Zip).Kind);
        Assert.Equal(FileKind.Pptx, FileKindDetector.Detect("f.pptx", "application/octet-stream", TestBytes.Zip).Kind);
        Assert.Null(FileKindDetector.Detect("f.zip", null, TestBytes.Zip).Kind);
    }

    [Fact]
    public void Html_detected_by_content_type_extension_or_sniff()
    {
        Assert.Equal(FileKind.Html, FileKindDetector.Detect("p.bin", "text/html; charset=utf-8", "<p>x</p>"u8).Kind);
        Assert.Equal(FileKind.Html, FileKindDetector.Detect("p.html", null, "<p>x</p>"u8).Kind);
        Assert.Equal(FileKind.Html, FileKindDetector.Detect("p.bin", null, TestBytes.Html).Kind);
    }

    [Fact]
    public void Claimed_type_contradicting_bytes_trusts_bytes_and_warns()
    {
        var r = FileKindDetector.Detect("fake.pdf", "application/pdf", TestBytes.Png);
        Assert.Equal(FileKind.Image, r.Kind);
        Assert.NotNull(r.Warning);
        Assert.Contains("application/pdf", r.Warning);
    }

    [Fact]
    public void Matching_claim_produces_no_warning()
    {
        Assert.Null(FileKindDetector.Detect("a.pdf", "application/pdf", TestBytes.Pdf).Warning);
    }

    [Fact]
    public void Undetectable_returns_null_kind()
    {
        Assert.Null(FileKindDetector.Detect("x.bin", null, TestBytes.Garbage).Kind);
        var claimed = FileKindDetector.Detect("x.pdf", "application/pdf", TestBytes.Garbage);
        Assert.Null(claimed.Kind);
        Assert.NotNull(claimed.Warning);
    }
}
