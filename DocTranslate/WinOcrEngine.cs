using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using WinOcr = Windows.Media.Ocr;

namespace DocTranslate;

/// <summary>
/// Wraps Windows.Media.Ocr — CPU-native, no install required on Windows 10/11.
/// Returns OCR lines in PDF-point coordinates (origin top-left, y down).
/// </summary>
public record OcrLine(string Text, double X, double Y, double W, double H);

public static class WinOcrEngine
{
    private static WinOcr.OcrEngine? GetEngine(string langCode)
    {
        foreach (var tag in new[] { MapLang(langCode), "en" })
        {
            try
            {
                var e = WinOcr.OcrEngine.TryCreateFromLanguage(new Language(tag));
                if (e is not null) return e;
            }
            catch { }
        }
        return null;
    }

    private static string MapLang(string code) => code switch
    {
        "zh" => "zh-Hans",
        "nb" => "nb",
        _    => code,
    };

    /// <summary>
    /// OCR a PNG/JPEG/BMP file produced by rasterizing a PDF page.
    /// <paramref name="dpi"/> is the DPI used when rasterizing (default 150).
    /// Returned coordinates are in PDF points (72 pt = 1 inch), y measured from top.
    /// </summary>
    public static async Task<List<OcrLine>> RecognizeFileAsync(
        string imagePath, string langCode, double dpi = 150)
    {
        try
        {
            var bytes = await File.ReadAllBytesAsync(imagePath);
            return await RecognizeCoreAsync(bytes, langCode, dpi);
        }
        catch { return new(); }
    }

    /// <summary>
    /// OCR raw image bytes (e.g. an image embedded in a DOCX/PPTX).
    /// Assumes 96 DPI screen resolution; returned coordinates are in points.
    /// </summary>
    public static Task<List<OcrLine>> RecognizeBytesAsync(
        byte[] imageBytes, string langCode, double dpi = 96)
        => RecognizeCoreAsync(imageBytes, langCode, dpi);

    private static async Task<List<OcrLine>> RecognizeCoreAsync(
        byte[] imageBytes, string langCode, double dpi)
    {
        var engine = GetEngine(langCode);
        if (engine is null) return new();

        try
        {
            SoftwareBitmap bitmap;
            double effectiveDpi = dpi;

            using (var ms  = new MemoryStream(imageBytes))
            using (var ras = ms.AsRandomAccessStream())
            {
                var decoder = await BitmapDecoder.CreateAsync(ras);
                uint maxDim = WinOcr.OcrEngine.MaxImageDimension;
                uint srcW   = decoder.PixelWidth;
                uint srcH   = decoder.PixelHeight;
                double scale = Math.Min(1.0, (double)maxDim / Math.Max(srcW, srcH));

                if (scale < 1.0)
                {
                    effectiveDpi *= scale;
                    var t = new BitmapTransform
                    {
                        InterpolationMode = BitmapInterpolationMode.Fant,
                        ScaledWidth       = (uint)(srcW * scale),
                        ScaledHeight      = (uint)(srcH * scale),
                    };
                    bitmap = await decoder.GetSoftwareBitmapAsync(
                        BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied,
                        t, ExifOrientationMode.IgnoreExifOrientation,
                        ColorManagementMode.DoNotColorManage);
                }
                else
                {
                    bitmap = await decoder.GetSoftwareBitmapAsync(
                        BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
                }
            }

            var result     = await engine.RecognizeAsync(bitmap);
            double toPoint = 72.0 / effectiveDpi;

            return result.Lines
                .Select(line =>
                {
                    double x = double.MaxValue, y = double.MaxValue, x2 = 0, y2 = 0;
                    foreach (var w in line.Words)
                    {
                        x  = Math.Min(x,  w.BoundingRect.X);
                        y  = Math.Min(y,  w.BoundingRect.Y);
                        x2 = Math.Max(x2, w.BoundingRect.X + w.BoundingRect.Width);
                        y2 = Math.Max(y2, w.BoundingRect.Y + w.BoundingRect.Height);
                    }
                    return new OcrLine(line.Text,
                        x * toPoint, y * toPoint,
                        (x2 - x) * toPoint, (y2 - y) * toPoint);
                })
                .Where(l => !string.IsNullOrWhiteSpace(l.Text))
                .ToList();
        }
        catch { return new(); }
    }
}
