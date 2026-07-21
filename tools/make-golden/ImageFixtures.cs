using SkiaSharp;

namespace MakeGolden;

public static class ImageFixtures
{
    /// <summary>Synthetic sunglasses product photo with visible "UV400" text.</summary>
    public static byte[] SunglassesPng()
    {
        using var surface = SKSurface.Create(new SKImageInfo(640, 400));
        var canvas = surface.Canvas;
        canvas.Clear(new SKColor(230, 230, 235));

        using var frame = new SKPaint { Color = new SKColor(25, 25, 25), IsAntialias = true };
        using var lens = new SKPaint { Color = new SKColor(60, 45, 95), IsAntialias = true };
        canvas.DrawRect(new SKRect(130, 150, 510, 165), frame);
        canvas.DrawCircle(210, 200, 75, frame);
        canvas.DrawCircle(430, 200, 75, frame);
        canvas.DrawCircle(210, 200, 65, lens);
        canvas.DrawCircle(430, 200, 65, lens);

        using var ink = new SKPaint { Color = new SKColor(15, 15, 15), IsAntialias = true };
        using var big = new SKFont(SKTypeface.Default, 42);
        using var small = new SKFont(SKTypeface.Default, 24);
        canvas.DrawText("UV400", 265, 330, SKTextAlign.Left, big, ink);
        canvas.DrawText("Sunglasses model SG-1", 175, 370, SKTextAlign.Left, small, ink);

        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }

    /// <summary>A "scanned page": white JPEG with rendered text only — proves OCR in live smoke.</summary>
    public static (byte[] Jpeg, int Width, int Height) ScannedPageJpeg()
    {
        const int width = 1240;
        const int height = 1650;
        using var surface = SKSurface.Create(new SKImageInfo(width, height));
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.White);

        using var ink = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        using var font = new SKFont(SKTypeface.Default, 48);
        canvas.DrawText("EuGo scanned golden page", 100, 220, SKTextAlign.Left, font, ink);
        canvas.DrawText("OCR TARGET UV400", 100, 320, SKTextAlign.Left, font, ink);
        canvas.DrawText("This page exists to prove OCR works", 100, 420, SKTextAlign.Left, font, ink);

        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Jpeg, 85);
        return (data.ToArray(), width, height);
    }
}
