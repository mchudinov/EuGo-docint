using DocInt.Api.Contracts;
using DocInt.Api.Engines;

namespace DocInt.Tests;

public class SpreadsheetEngineTests
{
    private static async Task<EngineOutcome> Run(string fixture)
    {
        var engine = new SpreadsheetEngine();
        var item = new FileItem { Index = 0, Name = fixture, Kind = FileKind.Xlsx, Bytes = Golden.Bytes(fixture) };
        return await engine.ExtractAsync(item, CancellationToken.None);
    }

    [Fact]
    public async Task Bom_xlsx_yields_exact_typed_cells()
    {
        var outcome = await Run("bom.xlsx");
        var bom = outcome.Result.Tables![0];

        Assert.Equal("BoM", bom.Name);
        Assert.Equal(["Part", "Qty", "UnitPrice", "Total", "InStock", "Updated"],
            bom.Rows[0].Cast<string>().ToArray());
        Assert.Equal("M3 screw", bom.Rows[1][0]);
        Assert.Equal(40m, bom.Rows[1][1]);
        Assert.Equal(19.99m, bom.Rows[1][2]);       // exact — never re-parsed display text
        Assert.Equal(799.6m, bom.Rows[1][3]);       // formula cell: cached value
        Assert.Equal(true, bom.Rows[1][4]);
        Assert.Equal("2026-07-19", bom.Rows[1][5]); // date-styled number → ISO-8601
        Assert.Equal(0.125m, bom.Rows[2][2]);
        Assert.Equal(false, bom.Rows[2][4]);
        Assert.Null(outcome.Result.Error);
    }

    [Fact]
    public async Task Markdown_renders_both_sheets()
    {
        var outcome = await Run("bom.xlsx");
        Assert.Equal(2, outcome.Result.Tables!.Count);
        Assert.Equal("Notes", outcome.Result.Tables[1].Name);
        Assert.Contains("## BoM", outcome.Result.Markdown);
        Assert.Contains("## Notes", outcome.Result.Markdown);
        Assert.Contains("| M3 screw | 40 | 19.99 |", outcome.Result.Markdown);
        Assert.Equal(2, outcome.PagesProcessed);
    }

    [Fact]
    public async Task Corrupt_xlsx_maps_to_corrupt_error()
    {
        var outcome = await Run("corrupt.xlsx");
        Assert.Equal(ErrorCodes.Corrupt, outcome.Result.Error!.Code);
    }

    [Fact]
    public async Task Chartsheet_sheet_is_skipped_with_warning_not_thrown()
    {
        // Regression: a <sheet> pointing at a chartsheet/dialogsheet resolves via GetPartById to a
        // non-WorksheetPart; the engine must skip that tab (with a warning) rather than throw, and
        // must still extract the real worksheet(s) present in the same workbook.
        var outcome = await Run("chartsheet.xlsx");

        Assert.Null(outcome.Result.Error);
        var data = Assert.Single(outcome.Result.Tables!);
        Assert.Equal("Data", data.Name);
        Assert.Equal(["Item", "Count"], data.Rows[0].Cast<string>().ToArray());
        Assert.Equal("Widget", data.Rows[1][0]);
        Assert.Equal(7m, data.Rows[1][1]);
        Assert.Contains(outcome.Result.Warnings, w => w.Contains("Chart1") && w.Contains("skipped"));
        Assert.Equal(1, outcome.PagesProcessed);
    }

    [Fact]
    public async Task Http_contract_returns_typed_json_numbers()
    {
        using var factory = new ContractTestFactory();
        using var form = Multipart.Form(("bom.xlsx", Golden.Bytes("bom.xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        var response = await factory.CreateClient().PostAsync("/v1/extract", form);
        var json = await response.Content.ReadAsStringAsync();
        Assert.Contains("\"kind\":\"xlsx\"", json);
        Assert.Contains("19.99", json);
        Assert.DoesNotContain("\"19.99\"", json);   // number, not string
        Assert.Contains("\"2026-07-19\"", json);
    }
}
