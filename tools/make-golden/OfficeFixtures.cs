using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace MakeGolden;

public static class OfficeFixtures
{
    public static byte[] Docx()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("EuGo DOCX golden document"))),
                new W.Paragraph(new W.Run(new W.Text("Bill of Materials evidence text for extraction")))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    public static byte[] BomXlsx()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new S.Workbook();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new S.Stylesheet(
                new S.Fonts(new S.Font()),
                new S.Fills(new S.Fill()),
                new S.Borders(new S.Border()),
                new S.CellFormats(
                    new S.CellFormat(),
                    new S.CellFormat { NumberFormatId = 14, ApplyNumberFormat = true }));

            var shared = workbookPart.AddNewPart<SharedStringTablePart>();
            shared.SharedStringTable = new S.SharedStringTable();
            int Str(string s)
            {
                var i = 0;
                foreach (var item in shared.SharedStringTable.Elements<S.SharedStringItem>())
                {
                    if (item.InnerText == s) return i;
                    i++;
                }
                shared.SharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(s)));
                return i;
            }

            var bomSheet = workbookPart.AddNewPart<WorksheetPart>();
            bomSheet.Worksheet = new S.Worksheet(new S.SheetData(
                Row(1, SharedCell("A1", Str("Part")), SharedCell("B1", Str("Qty")),
                       SharedCell("C1", Str("UnitPrice")), SharedCell("D1", Str("Total")),
                       SharedCell("E1", Str("InStock")), SharedCell("F1", Str("Updated"))),
                Row(2, SharedCell("A2", Str("M3 screw")), NumberCell("B2", "40"), NumberCell("C2", "19.99"),
                       FormulaCell("D2", "B2*C2", "799.6"), BoolCell("E2", true),
                       DateCell("F2", new DateTime(2026, 7, 19))),
                Row(3, SharedCell("A3", Str("Washer")), NumberCell("B3", "100"), NumberCell("C3", "0.125"),
                       FormulaCell("D3", "B3*C3", "12.5"), BoolCell("E3", false),
                       DateCell("F3", new DateTime(2026, 1, 2)))));

            var notesSheet = workbookPart.AddNewPart<WorksheetPart>();
            notesSheet.Worksheet = new S.Worksheet(new S.SheetData(
                Row(1, SharedCell("A1", Str("EuGo BoM golden notes")))));

            workbookPart.Workbook.AppendChild(new S.Sheets(
                new S.Sheet { Id = workbookPart.GetIdOfPart(bomSheet), SheetId = 1U, Name = "BoM" },
                new S.Sheet { Id = workbookPart.GetIdOfPart(notesSheet), SheetId = 2U, Name = "Notes" }));
            workbookPart.Workbook.Save();
        }
        return ms.ToArray();

        static S.Row Row(uint index, params S.Cell[] cells)
        {
            var row = new S.Row { RowIndex = index };
            row.Append(cells);
            return row;
        }
        static S.Cell SharedCell(string r, int sharedIndex) => new()
        { CellReference = r, DataType = S.CellValues.SharedString, CellValue = new S.CellValue(sharedIndex.ToString()) };
        static S.Cell NumberCell(string r, string number) => new()
        { CellReference = r, CellValue = new S.CellValue(number) };
        static S.Cell FormulaCell(string r, string formula, string cached) => new()
        { CellReference = r, CellFormula = new S.CellFormula(formula), CellValue = new S.CellValue(cached) };
        static S.Cell BoolCell(string r, bool value) => new()
        { CellReference = r, DataType = S.CellValues.Boolean, CellValue = new S.CellValue(value ? "1" : "0") };
        static S.Cell DateCell(string r, DateTime date) => new()
        { CellReference = r, StyleIndex = 1U, CellValue = new S.CellValue(date.ToOADate().ToString(CultureInfo.InvariantCulture)) };
    }

    /// <summary>
    /// A workbook with one normal worksheet plus a chartsheet tab. Chartsheets/dialogsheets (a chart or
    /// dialog placed on its own tab) are a normal Excel feature: their &lt;sheet&gt; entry resolves via
    /// GetPartById to a ChartsheetPart/DialogsheetPart, not a WorksheetPart. This fixture proves the
    /// engine skips that tab with a warning instead of throwing.
    /// </summary>
    public static byte[] ChartsheetXlsx()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new S.Workbook();

            var shared = workbookPart.AddNewPart<SharedStringTablePart>();
            shared.SharedStringTable = new S.SharedStringTable();
            int Str(string s)
            {
                var i = 0;
                foreach (var item in shared.SharedStringTable.Elements<S.SharedStringItem>())
                {
                    if (item.InnerText == s) return i;
                    i++;
                }
                shared.SharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(s)));
                return i;
            }

            var dataSheet = workbookPart.AddNewPart<WorksheetPart>();
            dataSheet.Worksheet = new S.Worksheet(new S.SheetData(
                Row(1, SharedCell("A1", Str("Item")), SharedCell("B1", Str("Count"))),
                Row(2, SharedCell("A2", Str("Widget")), NumberCell("B2", "7"))));

            var chartsheetPart = workbookPart.AddNewPart<ChartsheetPart>();
            chartsheetPart.Chartsheet = new S.Chartsheet(
                new S.ChartSheetProperties(),
                new S.ChartSheetViews(new S.ChartSheetView { WorkbookViewId = 0 }));

            workbookPart.Workbook.AppendChild(new S.Sheets(
                new S.Sheet { Id = workbookPart.GetIdOfPart(dataSheet), SheetId = 1U, Name = "Data" },
                new S.Sheet { Id = workbookPart.GetIdOfPart(chartsheetPart), SheetId = 2U, Name = "Chart1" }));
            workbookPart.Workbook.Save();
        }
        return ms.ToArray();

        static S.Row Row(uint index, params S.Cell[] cells)
        {
            var row = new S.Row { RowIndex = index };
            row.Append(cells);
            return row;
        }
        static S.Cell SharedCell(string r, int sharedIndex) => new()
        { CellReference = r, DataType = S.CellValues.SharedString, CellValue = new S.CellValue(sharedIndex.ToString()) };
        static S.Cell NumberCell(string r, string number) => new()
        { CellReference = r, CellValue = new S.CellValue(number) };
    }

    public static byte[] Pptx(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
        {
            var presentationPart = doc.AddPresentationPart();
            presentationPart.Presentation = new P.Presentation();

            var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rIdMaster");
            masterPart.SlideMaster = new P.SlideMaster(
                new P.CommonSlideData(EmptyShapeTree()),
                new P.ColorMap
                {
                    Background1 = D.ColorSchemeIndexValues.Light1,
                    Text1 = D.ColorSchemeIndexValues.Dark1,
                    Background2 = D.ColorSchemeIndexValues.Light2,
                    Text2 = D.ColorSchemeIndexValues.Dark2,
                    Accent1 = D.ColorSchemeIndexValues.Accent1,
                    Accent2 = D.ColorSchemeIndexValues.Accent2,
                    Accent3 = D.ColorSchemeIndexValues.Accent3,
                    Accent4 = D.ColorSchemeIndexValues.Accent4,
                    Accent5 = D.ColorSchemeIndexValues.Accent5,
                    Accent6 = D.ColorSchemeIndexValues.Accent6,
                    Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                });

            var themePart = masterPart.AddNewPart<ThemePart>("rIdTheme");
            themePart.Theme = MinimalTheme();

            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rIdLayout");
            layoutPart.SlideLayout = new P.SlideLayout(
                new P.CommonSlideData(EmptyShapeTree()),
                new P.ColorMapOverride(new D.MasterColorMapping()));
            masterPart.SlideMaster.AppendChild(new P.SlideLayoutIdList(
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = "rIdLayout" }));

            var slidePart = presentationPart.AddNewPart<SlidePart>("rIdSlide");
            slidePart.AddPart(layoutPart);
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 2U, Name = "Title" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle(),
                            new D.Paragraph(new D.Run(new D.Text(text))))))),
                new P.ColorMapOverride(new D.MasterColorMapping()));

            presentationPart.Presentation.Append(
                new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = "rIdMaster" }),
                new P.SlideIdList(new P.SlideId { Id = 256U, RelationshipId = "rIdSlide" }),
                new P.SlideSize { Cx = 9144000, Cy = 6858000 },
                new P.NotesSize { Cx = 6858000, Cy = 9144000 });
            presentationPart.Presentation.Save();
        }
        return ms.ToArray();

        static P.ShapeTree EmptyShapeTree() => new(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties());
    }

    private static D.Theme MinimalTheme() => new(
        new D.ThemeElements(
            new D.ColorScheme(
                new D.Dark1Color(new D.SystemColor { Val = D.SystemColorValues.WindowText }),
                new D.Light1Color(new D.SystemColor { Val = D.SystemColorValues.Window }),
                new D.Dark2Color(new D.RgbColorModelHex { Val = "44546A" }),
                new D.Light2Color(new D.RgbColorModelHex { Val = "E7E6E6" }),
                new D.Accent1Color(new D.RgbColorModelHex { Val = "4472C4" }),
                new D.Accent2Color(new D.RgbColorModelHex { Val = "ED7D31" }),
                new D.Accent3Color(new D.RgbColorModelHex { Val = "A5A5A5" }),
                new D.Accent4Color(new D.RgbColorModelHex { Val = "FFC000" }),
                new D.Accent5Color(new D.RgbColorModelHex { Val = "5B9BD5" }),
                new D.Accent6Color(new D.RgbColorModelHex { Val = "70AD47" }),
                new D.Hyperlink(new D.RgbColorModelHex { Val = "0563C1" }),
                new D.FollowedHyperlinkColor(new D.RgbColorModelHex { Val = "954F72" }))
            { Name = "Office" },
            new D.FontScheme(
                new D.MajorFont(
                    new D.LatinFont { Typeface = "Calibri Light" },
                    new D.EastAsianFont { Typeface = "" },
                    new D.ComplexScriptFont { Typeface = "" }),
                new D.MinorFont(
                    new D.LatinFont { Typeface = "Calibri" },
                    new D.EastAsianFont { Typeface = "" },
                    new D.ComplexScriptFont { Typeface = "" }))
            { Name = "Office" },
            new D.FormatScheme(
                new D.FillStyleList(SolidPh(), SolidPh(), SolidPh()),
                new D.LineStyleList(
                    new D.Outline(SolidPh()),
                    new D.Outline(SolidPh()),
                    new D.Outline(SolidPh())),
                new D.EffectStyleList(
                    new D.EffectStyle(new D.EffectList()),
                    new D.EffectStyle(new D.EffectList()),
                    new D.EffectStyle(new D.EffectList())),
                new D.BackgroundFillStyleList(SolidPh(), SolidPh(), SolidPh()))
            { Name = "Office" }))
    { Name = "MinimalTheme" };

    private static D.SolidFill SolidPh() =>
        new(new D.SchemeColor { Val = D.SchemeColorValues.PhColor });
}
