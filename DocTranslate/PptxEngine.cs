using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

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

        foreach (var sp in slideParts)
            sp.Slide.Save();

        return tmp;
    }
}
