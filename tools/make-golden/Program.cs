using System.Text;
using MakeGolden;

var outputDir = args.Length > 0 ? args[0] : "tests/DocInt.Tests/golden";
Directory.CreateDirectory(outputDir);

void Save(string name, byte[] bytes)
{
    File.WriteAllBytes(Path.Combine(outputDir, name), bytes);
    Console.WriteLine($"{name}: {bytes.Length} bytes");
}

Save("text.pdf", MiniPdf.TextPdf("EuGo text PDF golden - Part M3 screw UV400"));

var (jpeg, width, height) = ImageFixtures.ScannedPageJpeg();
Save("scanned.pdf", MiniPdf.ImagePdf(jpeg, width, height));

Save("sample.docx", OfficeFixtures.Docx());
Save("sample.pptx", OfficeFixtures.Pptx("EuGo PPTX golden slide"));
Save("sample.html", Encoding.UTF8.GetBytes(
    "<!DOCTYPE html><html><body><h1>EuGo HTML golden</h1>"
    + "<table><tr><th>Part</th><th>Qty</th></tr><tr><td>M3 screw</td><td>40</td></tr></table>"
    + "</body></html>"));
Save("bom.xlsx", OfficeFixtures.BomXlsx());
Save("chartsheet.xlsx", OfficeFixtures.ChartsheetXlsx());
Save("overflow.xlsx", OfficeFixtures.OverflowXlsx());
Save("malformed-cells.xlsx", OfficeFixtures.MalformedCellsXlsx());
Save("photo.png", ImageFixtures.SunglassesPng());
Save("corrupt.xlsx", [0x50, 0x4B, 0x03, 0x04, .. Enumerable.Repeat((byte)0xDE, 64)]);
Save("unknown.bin", [.. Enumerable.Range(1, 64).Select(i => (byte)i)]);

Console.WriteLine($"Golden fixtures written to {Path.GetFullPath(outputDir)}");
