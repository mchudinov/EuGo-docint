using System.Net;
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

    // --- Regression: crafted "1e400" numeric cell must never 500 the request. ---
    // decimal.Parse overflows (as before); the OverflowException fallback previously
    // returned double.Parse's result unconditionally, and double.Parse("1e400") yields
    // double.PositiveInfinity without throwing. That non-finite double used to survive
    // into the response and crash System.Text.Json serialization outside any try/catch.

    [Fact]
    public async Task Overflow_cell_is_kept_as_text_with_warning_not_infinity()
    {
        var outcome = await Run("overflow.xlsx");

        Assert.Null(outcome.Result.Error);
        var sheet = Assert.Single(outcome.Result.Tables!);
        Assert.Equal("normal", sheet.Rows[1][0]);
        Assert.Equal(42m, sheet.Rows[1][1]);            // sibling row unaffected
        Assert.Equal("huge", sheet.Rows[2][0]);
        Assert.Equal("1e400", sheet.Rows[2][1]);         // kept as text, never a non-finite double
        Assert.Contains(outcome.Result.Warnings,
            w => w.Contains("B3") && w.Contains("out of numeric range"));
    }

    [Fact]
    public async Task Http_contract_returns_200_for_overflow_cell()
    {
        using var factory = new ContractTestFactory();
        using var form = Multipart.Form(("overflow.xlsx", Golden.Bytes("overflow.xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

        var response = await factory.CreateClient().PostAsync("/v1/extract", form);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var json = await response.Content.ReadAsStringAsync();
        Assert.DoesNotContain("\"error\":{", json);
        Assert.Contains("\"1e400\"", json);
    }

    // --- Regression: a malformed-but-openable date/numeric cell must degrade to text +
    // warning on that one cell, not fail the whole file (and must never surface the raw
    // parse-exception text, which can embed cell content, as the file-level error message). ---

    [Fact]
    public async Task Malformed_cells_degrade_to_warnings_not_engine_error()
    {
        var outcome = await Run("malformed-cells.xlsx");

        Assert.Null(outcome.Result.Error);
        var sheet = Assert.Single(outcome.Result.Tables!);

        Assert.Equal("good", sheet.Rows[1][0]);
        Assert.Equal(42m, sheet.Rows[1][1]);
        Assert.Equal(1m, sheet.Rows[1][2]);

        Assert.Equal("bad number", sheet.Rows[2][0]);
        Assert.Equal("not-a-number", sheet.Rows[2][1]);  // FormatException on decimal/double.Parse
        Assert.Equal(2m, sheet.Rows[2][2]);              // sibling cell in the same row intact

        Assert.Equal("bad date", sheet.Rows[3][0]);
        Assert.Equal("not-a-date", sheet.Rows[3][1]);    // FormatException on DateTime.Parse
        Assert.Equal(3m, sheet.Rows[3][2]);

        Assert.Equal("bad serial", sheet.Rows[4][0]);
        Assert.Equal("1e30", sheet.Rows[4][1]);          // ArgumentException on DateTime.FromOADate
        Assert.Equal(4m, sheet.Rows[4][2]);

        Assert.Contains(outcome.Result.Warnings, w => w.Contains("B3") && w.Contains("number"));
        Assert.Contains(outcome.Result.Warnings, w => w.Contains("B4") && w.Contains("date"));
        Assert.Contains(outcome.Result.Warnings, w => w.Contains("B5") && w.Contains("date"));
    }
}
