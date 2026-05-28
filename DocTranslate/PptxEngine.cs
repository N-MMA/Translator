using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;

namespace DocTranslate;

/// <summary>
/// Translates .pptx files.
/// Primary:  Microsoft PowerPoint COM Interop (requires PowerPoint).
/// Fallback: DocumentFormat.OpenXml (covers shapes, tables, text boxes).
/// </summary>
public static class PptxEngine
{
    public static bool IsPowerPointAvailable()
    {
        try { return Type.GetTypeFromProgID("PowerPoint.Application") is not null; }
        catch { return false; }
    }

    public static async Task<string> TranslateAsync(
        string srcPath, string fromCode, string toCode,
        TranslatorBridge bridge,
        IProgress<(int Done, int Total)>? progress = null)
    {
        return IsPowerPointAvailable()
            ? await TranslateViaPptComAsync(srcPath, fromCode, toCode, bridge, progress)
            : await TranslateViaOpenXmlAsync(srcPath, fromCode, toCode, bridge, progress);
    }

    // ── PowerPoint COM ────────────────────────────────────────────────────────
    private static async Task<string> TranslateViaPptComAsync(
        string srcPath, string fromCode, string toCode,
        TranslatorBridge bridge,
        IProgress<(int Done, int Total)>? progress)
    {
        var tmp = Path.Combine(Path.GetTempPath(), $"doctranslate_{Guid.NewGuid():N}.pptx");
        File.Copy(srcPath, tmp, overwrite: true);

        await Task.Run(() =>
        {
            dynamic? pptApp = null;
            dynamic? pres   = null;
            try
            {
                var pptType = Type.GetTypeFromProgID("PowerPoint.Application")!;
                pptApp = Activator.CreateInstance(pptType)!;
                // WithWindow: 0 = msoFalse — open without a visible window
                pres = pptApp.Presentations.Open(tmp, 0, 0, 0);

                var textRanges = new List<(dynamic Range, string Text)>();

                foreach (dynamic slide in pres.Slides)
                {
                    foreach (dynamic shape in slide.Shapes)
                    {
                        try
                        {
                            if ((bool)shape.HasTextFrame)
                                CollectPptRanges(shape.TextFrame.TextRange, textRanges);
                        }
                        catch { }

                        try
                        {
                            if ((bool)shape.HasTable)
                            {
                                dynamic tbl = shape.Table;
                                for (int r = 1; r <= (int)tbl.Rows.Count; r++)
                                    for (int c = 1; c <= (int)tbl.Columns.Count; c++)
                                        CollectPptRanges(
                                            tbl.Cell(r, c).Shape.TextFrame.TextRange,
                                            textRanges);
                            }
                        }
                        catch { }
                    }
                }

                int total = textRanges.Count, done = 0;
                foreach (var (range, original) in textRanges)
                {
                    if (!string.IsNullOrWhiteSpace(original))
                    {
                        var translated = bridge.TranslateAsync(original, fromCode, toCode)
                            .GetAwaiter().GetResult();
                        if (!string.IsNullOrEmpty(translated) && translated != original)
                            range.Text = translated;
                    }
                    done++;
                    progress?.Report((done, total));
                }

                pres.Save();
            }
            finally
            {
                if (pres   is not null) { try { pres.Close();   } catch { } Marshal.ReleaseComObject(pres);   }
                if (pptApp is not null) { try { pptApp.Quit();  } catch { } Marshal.ReleaseComObject(pptApp); }
            }
        });

        return tmp;
    }

    private static void CollectPptRanges(dynamic textRange, List<(dynamic, string)> list)
    {
        string text = ((string)(textRange.Text ?? ""))
            .TrimEnd('\r', '\n', '\x0D', '\x0A');
        if (!string.IsNullOrWhiteSpace(text))
            list.Add((textRange, text));
    }

    // ── OpenXML fallback ──────────────────────────────────────────────────────
    private static async Task<string> TranslateViaOpenXmlAsync(
        string srcPath, string fromCode, string toCode,
        TranslatorBridge bridge,
        IProgress<(int Done, int Total)>? progress)
    {
        var tmp = Path.Combine(Path.GetTempPath(), $"doctranslate_{Guid.NewGuid():N}.pptx");
        File.Copy(srcPath, tmp, overwrite: true);

        using var pres = DocumentFormat.OpenXml.Packaging.PresentationDocument
            .Open(tmp, isEditable: true);
        var presPart = pres.PresentationPart
            ?? throw new Exception("Invalid PPTX: no presentation part.");

        // Collect paragraphs from all slides (tables are nested inside shapes so
        // Descendants covers them automatically).
        var slideParts = presPart.SlideParts.ToList();
        var allParas   = slideParts
            .SelectMany(sp => sp.Slide
                .Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            .ToList();

        int total = allParas.Count, done = 0;
        foreach (var para in allParas)
        {
            // Only walk <a:r> regular runs — skip <a:fld> (slide numbers, dates).
            var runs = para
                .Descendants<DocumentFormat.OpenXml.Drawing.Run>()
                .ToList();

            if (runs.Any())
            {
                var fullText = string.Concat(runs.SelectMany(r =>
                    r.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                     .Select(t => t.Text)));

                if (!string.IsNullOrWhiteSpace(fullText))
                {
                    var translated = await bridge.TranslateAsync(fullText, fromCode, toCode);
                    if (!string.IsNullOrEmpty(translated) && translated != fullText)
                    {
                        bool first = true;
                        foreach (var run in runs)
                            foreach (var t in run.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList())
                            { if (first) { t.Text = translated; first = false; } else t.Text = ""; }
                    }
                }
            }
            done++;
            progress?.Report((done, total));
        }

        // OCR every embedded image on each slide and add a translated text shape.
        foreach (var sp in slideParts)
            await OcrAndInsertPptxImagesAsync(sp, fromCode, toCode, bridge);

        foreach (var sp in slideParts)
            sp.Slide.Save();

        return tmp;
    }

    private static async Task OcrAndInsertPptxImagesAsync(
        DocumentFormat.OpenXml.Packaging.SlidePart slidePart,
        string fromCode, string toCode, TranslatorBridge bridge)
    {
        // Find all picture shapes (<p:pic>) in the slide.
        var pictures = slidePart.Slide
            .Descendants<DocumentFormat.OpenXml.Presentation.Picture>()
            .ToList();

        foreach (var pic in pictures)
        {
            var blip = pic.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            if (blip?.Embed?.Value is not { } relId) continue;

            if (slidePart.GetPartById(relId) is not DocumentFormat.OpenXml.Packaging.ImagePart imgPart)
                continue;

            byte[] imgBytes;
            using (var s = imgPart.GetStream())
            {
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                imgBytes = ms.ToArray();
            }

            var lines = await WinOcrEngine.RecognizeBytesAsync(imgBytes, fromCode);
            if (lines.Count == 0) continue;

            var ocrText    = string.Join(" ", lines.Select(l => l.Text));
            var translated = await bridge.TranslateAsync(ocrText, fromCode, toCode);
            if (string.IsNullOrWhiteSpace(translated)) continue;

            // Get the picture's position and size from its transform so we can
            // place the text box directly below it.
            var xfrm = pic.Descendants<DocumentFormat.OpenXml.Drawing.Transform2D>().FirstOrDefault();
            long offX = xfrm?.Offset?.X ?? 457200L;   // default ~0.5 inch in EMUs
            long offY = xfrm?.Offset?.Y ?? 457200L;
            long extX = xfrm?.Extents?.Cx ?? 2743200L; // default ~3 inch wide
            long extY = xfrm?.Extents?.Cy ?? 914400L;  // default ~1 inch tall

            // Add a text box shape immediately below the picture.
            long tbY    = offY + extY + 91440L; // 0.1 inch gap
            long tbH    = 457200L;              // 0.5 inch tall
            var  spTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

            var sp = BuildTextShape(translated, offX, tbY, extX, tbH);
            spTree.Append(sp);
        }
    }

    private static DocumentFormat.OpenXml.Presentation.Shape BuildTextShape(
        string text, long x, long y, long cx, long cy)
    {
        var shape = new DocumentFormat.OpenXml.Presentation.Shape();

        shape.NonVisualShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 9000, Name = "OcrCaption" },
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(
                new DocumentFormat.OpenXml.Drawing.ShapeLocks { NoGrouping = true }),
            new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties(
                new DocumentFormat.OpenXml.Presentation.PlaceholderShape()));

        shape.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties(
            new DocumentFormat.OpenXml.Drawing.Transform2D(
                new DocumentFormat.OpenXml.Drawing.Offset { X = x, Y = y },
                new DocumentFormat.OpenXml.Drawing.Extents { Cx = cx, Cy = cy }),
            new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle });

        shape.TextBody = new DocumentFormat.OpenXml.Presentation.TextBody(
            new DocumentFormat.OpenXml.Drawing.BodyProperties(),
            new DocumentFormat.OpenXml.Drawing.ListStyle(),
            new DocumentFormat.OpenXml.Drawing.Paragraph(
                new DocumentFormat.OpenXml.Drawing.Run(
                    new DocumentFormat.OpenXml.Drawing.RunProperties
                        { Language = "pt-PT", FontSize = 1000, Italic = true },
                    new DocumentFormat.OpenXml.Drawing.Text(text))));

        return shape;
    }
}
