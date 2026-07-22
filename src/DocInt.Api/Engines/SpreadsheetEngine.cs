using System.Globalization;
using System.Text;
using DocInt.Api.Contracts;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace DocInt.Api.Engines;

/// <summary>
/// XLSX via OpenXML — the deliberate precision path: typed cell values straight from the
/// stored XML (never re-parsed display text), plus a Markdown rendering of each sheet.
/// </summary>
public sealed class SpreadsheetEngine : IExtractionEngine
{
    public IReadOnlyCollection<FileKind> Kinds { get; } = [FileKind.Xlsx];

    public Task<EngineOutcome> ExtractAsync(FileItem file, CancellationToken ct) =>
        Task.FromResult(ExtractCore(file));

    private static EngineOutcome ExtractCore(FileItem file)
    {
        try
        {
            using var ms = new MemoryStream(file.Bytes, writable: false);
            using var doc = SpreadsheetDocument.Open(ms, isEditable: false);
            var workbookPart = doc.WorkbookPart
                ?? throw new InvalidDataException("workbook part missing");

            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
                .Elements<S.SharedStringItem>().Select(i => i.InnerText).ToArray() ?? [];
            var dateStyles = DateStyleIndexes(workbookPart.WorkbookStylesPart?.Stylesheet);

            var warnings = new List<string>(file.Warnings);
            var tables = new List<TableResult>();
            var sheets = workbookPart.Workbook.Sheets?.Elements<S.Sheet>().ToArray() ?? [];
            foreach (var sheet in sheets)
            {
                var sheetName = sheet.Name?.Value ?? "Sheet";
                var relationshipId = sheet.Id?.Value;
                // A <sheet> can point at a chartsheet/dialogsheet (a chart or dialog on its own tab —
                // a normal Excel feature) instead of a worksheet, and can in principle carry no r:id at
                // all. Neither is corruption: skip that one tab and keep extracting the rest.
                if (string.IsNullOrWhiteSpace(relationshipId)
                    || !workbookPart.TryGetPartById(relationshipId, out var part)
                    || part is not WorksheetPart worksheetPart)
                {
                    warnings.Add($"sheet '{sheetName}' skipped: not a worksheet");
                    continue;
                }
                var rows = ReadRows(worksheetPart, sharedStrings, dateStyles, warnings);
                tables.Add(new TableResult(sheetName, RenderMarkdown(rows), rows));
            }
            if (sheets.Length == 0) warnings.Add("workbook has no sheets");

            var markdown = string.Join("\n\n", tables.Select(t => $"## {t.Name}\n\n{t.Markdown}"));
            return new EngineOutcome(
                new FileResult(file.Name, file.Kind, markdown, tables, null, warnings, null),
                tables.Count);
        }
        catch (Exception ex) when (ex is DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
            or FileFormatException or InvalidDataException)
        {
            return Errors.For(file, ErrorCodes.Corrupt, $"file is not a readable XLSX workbook: {ex.Message}");
        }
    }

    private static HashSet<uint> DateStyleIndexes(S.Stylesheet? stylesheet)
    {
        var result = new HashSet<uint>();
        var formats = stylesheet?.CellFormats?.Elements<S.CellFormat>().ToArray() ?? [];
        for (var i = 0u; i < formats.Length; i++)
        {
            var id = formats[i].NumberFormatId?.Value ?? 0;
            if (id is >= 14 and <= 22 or >= 45 and <= 47) result.Add(i);
        }
        return result;
    }

    private static List<IReadOnlyList<object?>> ReadRows(
        WorksheetPart worksheetPart, string[] sharedStrings, HashSet<uint> dateStyles, List<string> warnings)
    {
        var grid = new List<IReadOnlyList<object?>>();
        var sheetData = worksheetPart.Worksheet.GetFirstChild<S.SheetData>();
        if (sheetData is null) return grid;

        var maxColumns = 0;
        var rawRows = new List<(uint Index, Dictionary<int, object?> Cells)>();
        foreach (var row in sheetData.Elements<S.Row>())
        {
            var cells = new Dictionary<int, object?>();
            var fallbackColumn = 0;
            foreach (var cell in row.Elements<S.Cell>())
            {
                var column = cell.CellReference?.Value is { } reference
                    ? ColumnIndex(reference) : fallbackColumn;
                fallbackColumn = column + 1;
                cells[column] = CellValue(cell, sharedStrings, dateStyles, warnings);
                maxColumns = Math.Max(maxColumns, column + 1);
            }
            rawRows.Add((row.RowIndex?.Value ?? (uint)(rawRows.Count + 1), cells));
        }

        foreach (var (_, cells) in rawRows)
        {
            var materialized = new object?[maxColumns];
            foreach (var (column, value) in cells) materialized[column] = value;
            grid.Add(materialized);
        }
        // Trim trailing all-null rows.
        while (grid.Count > 0 && grid[^1].All(v => v is null)) grid.RemoveAt(grid.Count - 1);
        return grid;
    }

    private static object? CellValue(
        S.Cell cell, string[] sharedStrings, HashSet<uint> dateStyles, List<string> warnings)
    {
        var raw = cell.DataType?.Value == S.CellValues.InlineString
            ? cell.InlineString?.InnerText
            : cell.CellValue?.InnerText;

        if (raw is null)
        {
            if (cell.CellFormula is not null)
                warnings.Add($"formula cell {cell.CellReference?.Value} has no cached value");
            return null;
        }

        var dataType = cell.DataType?.Value;
        if (dataType == S.CellValues.SharedString)
            return int.TryParse(raw, out var i) && i < sharedStrings.Length ? sharedStrings[i] : raw;
        if (dataType == S.CellValues.Boolean)
            return raw == "1";
        if (dataType == S.CellValues.String || dataType == S.CellValues.InlineString)
            return raw;
        if (dataType == S.CellValues.Error)
        {
            warnings.Add($"cell {cell.CellReference?.Value} contains error '{raw}'");
            return null;
        }
        if (dataType == S.CellValues.Date)
            return Iso(DateTime.Parse(raw, CultureInfo.InvariantCulture));

        // Numeric (no DataType): date-styled → ISO string, otherwise decimal (double on overflow).
        if (cell.StyleIndex?.Value is { } style && dateStyles.Contains(style))
            return Iso(DateTime.FromOADate(double.Parse(raw, CultureInfo.InvariantCulture)));
        try
        {
            return decimal.Parse(raw, NumberStyles.Float, CultureInfo.InvariantCulture);
        }
        catch (OverflowException)
        {
            return double.Parse(raw, CultureInfo.InvariantCulture);
        }
    }

    private static string Iso(DateTime value) =>
        value.TimeOfDay == TimeSpan.Zero
            ? value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : value.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);

    private static int ColumnIndex(string cellReference)
    {
        var index = 0;
        foreach (var c in cellReference)
        {
            if (!char.IsAsciiLetterUpper(c) && !char.IsAsciiLetterLower(c)) break;
            index = index * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }
        return index - 1;
    }

    private static string RenderMarkdown(List<IReadOnlyList<object?>> rows)
    {
        if (rows.Count == 0) return "";
        var sb = new StringBuilder();
        AppendRow(sb, rows[0]);
        sb.Append('|');
        for (var i = 0; i < rows[0].Count; i++) sb.Append(" --- |");
        sb.Append('\n');
        foreach (var row in rows.Skip(1)) AppendRow(sb, row);
        return sb.ToString().TrimEnd('\n');

        static void AppendRow(StringBuilder sb, IReadOnlyList<object?> row)
        {
            sb.Append('|');
            foreach (var cell in row)
                sb.Append(' ').Append(Render(cell)).Append(" |");
            sb.Append('\n');
        }

        static string Render(object? value) => value switch
        {
            null => "",
            bool b => b ? "true" : "false",
            decimal d => d.ToString(CultureInfo.InvariantCulture),
            double d => d.ToString(CultureInfo.InvariantCulture),
            string s => s.Replace("|", "\\|"),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""
        };
    }
}
