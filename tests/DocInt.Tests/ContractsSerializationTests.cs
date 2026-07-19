using System.Text.Json;
using DocInt.Api.Contracts;

namespace DocInt.Tests;

public class ContractsSerializationTests
{
    [Fact]
    public void Kind_serializes_lowercase_and_nulls_are_omitted()
    {
        var result = new FileResult("manual.pdf", FileKind.Pdf, "# md", null, null, ["w1"], null);
        var json = JsonSerializer.Serialize(new ExtractResponse([result]), DocIntJson.Options);
        Assert.Contains("\"kind\":\"pdf\"", json);
        Assert.Contains("\"warnings\":[\"w1\"]", json);
        Assert.DoesNotContain("tables", json);
        Assert.DoesNotContain("imageDescription", json);
        Assert.DoesNotContain("error", json);
    }

    [Fact]
    public void Typed_cells_serialize_as_native_json_scalars()
    {
        var table = new TableResult("Sheet1", "| a |",
            [["Part", 40m, 19.99m, true, null]]);
        var result = new FileResult("bom.xlsx", FileKind.Xlsx, "# md", [table], null, [], null);
        var json = JsonSerializer.Serialize(result, DocIntJson.Options);
        Assert.Contains("[\"Part\",40,19.99,true,null]", json);
    }

    [Fact]
    public void Error_result_serializes_code_and_message()
    {
        var result = new FileResult("x.bin", null, null, null, null, [],
            new FileError(ErrorCodes.UnsupportedType, "could not detect a supported file type for 'x.bin'"));
        var json = JsonSerializer.Serialize(result, DocIntJson.Options);
        Assert.Contains("\"code\":\"unsupported_type\"", json);
        Assert.DoesNotContain("\"kind\"", json);
    }

    [Fact]
    public void Kind_names_are_lowercase()
    {
        Assert.Equal("pdf", FileKind.Pdf.Name());
        Assert.Equal("xlsx", FileKind.Xlsx.Name());
        Assert.Equal("image", FileKind.Image.Name());
    }
}
