using System.Text;

namespace MakeGolden;

/// <summary>Deterministic minimal PDF writer — enough structure for Azure Document Intelligence.</summary>
public static class MiniPdf
{
    public static byte[] TextPdf(string text) => Build(
        Ascii("<< /Type /Catalog /Pages 2 0 R >>"),
        Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"),
        Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>"),
        StreamObject(Encoding.ASCII.GetBytes($"BT /F1 24 Tf 72 700 Td ({Escape(text)}) Tj ET")),
        Ascii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"));

    public static byte[] ImagePdf(byte[] jpeg, int width, int height) => Build(
        Ascii("<< /Type /Catalog /Pages 2 0 R >>"),
        Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"),
        Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>"),
        StreamObject(Encoding.ASCII.GetBytes("q 612 0 0 792 0 0 cm /Im0 Do Q")),
        StreamObject(jpeg,
            $"/Type /XObject /Subtype /Image /Width {width} /Height {height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode "));

    private static byte[] Ascii(string s) => Encoding.ASCII.GetBytes(s);

    private static string Escape(string s) =>
        s.Replace("\\", @"\\").Replace("(", @"\(").Replace(")", @"\)");

    private static byte[] StreamObject(byte[] data, string extraDict = "")
    {
        var head = Encoding.ASCII.GetBytes($"<< {extraDict}/Length {data.Length} >>\nstream\n");
        var tail = Encoding.ASCII.GetBytes("\nendstream");
        return [.. head, .. data, .. tail];
    }

    private static byte[] Build(params byte[][] objects)
    {
        using var ms = new MemoryStream();
        void WriteAscii(string s)
        {
            var bytes = Encoding.ASCII.GetBytes(s);
            ms.Write(bytes, 0, bytes.Length);
        }

        WriteAscii("%PDF-1.4\n");
        var offsets = new long[objects.Length];
        for (var i = 0; i < objects.Length; i++)
        {
            offsets[i] = ms.Position;
            WriteAscii($"{i + 1} 0 obj\n");
            ms.Write(objects[i], 0, objects[i].Length);
            WriteAscii("\nendobj\n");
        }
        var xrefPos = ms.Position;
        WriteAscii($"xref\n0 {objects.Length + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets)
            WriteAscii($"{offset:0000000000} 00000 n \n");
        WriteAscii($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xrefPos}\n%%EOF");
        return ms.ToArray();
    }
}
