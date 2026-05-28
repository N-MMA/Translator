using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace DocTranslate;

/// <summary>
/// Translates .docx files.
/// Primary:  Microsoft Word COM Interop (100% fidelity — requires Word).
/// Fallback: DocumentFormat.OpenXml   (~80% fidelity — no Word needed).
/// </summary>
public static class DocxEngine
{
    public static bool IsWordAvailable()
    {
        try { return Type.GetTypeFromProgID("Word.Application") is not null; }
        catch { return false; }
    }

    public static async Task<string> TranslateAsync(
        string srcPath, string fromCode, string toCode,
        TranslatorBridge bridge,
        IProgress<(int Done, int Total)>? progress = null)
    {
        return IsWordAvailable()
            ? await TranslateViaWordComAsync(srcPath, fromCode, toCode, bridge, progress)
            : await TranslateViaOpenXmlAsync(srcPath, fromCode, toCode, bridge, progress);
    }

    // ── Word COM ──────────────────────────────────────────────────────────────
    private static async Task<string> TranslateViaWordComAsync(
        string srcPath, string fromCode, string toCode,
        TranslatorBridge bridge,
        IProgress<(int Done, int Total)>? progress)
    {
        var tmp = Path.Combine(Path.GetTempPath(), $"doctranslate_{Guid.NewGuid():N}.docx");
        File.Copy(srcPath, tmp, overwrite: true);

        await Task.Run(() =>
        {
            dynamic? wordApp = null;
            dynamic? doc     = null;
            try
            {
                var wordType = Type.GetTypeFromProgID("Word.Application")!;
                wordApp = Activator.CreateInstance(wordType)!;
                wordApp.Visible = false;
                wordApp.DisplayAlerts = 0;
                doc = wordApp.Documents.Open(tmp);

                var ranges = new List<dynamic>();
                foreach (dynamic para  in doc.Paragraphs) ranges.Add(para.Range);
                foreach (dynamic table in doc.Tables)
                    foreach (dynamic row in table.Rows)
                        foreach (dynamic cell in row.Cells)
                            ranges.Add(cell.Range);
                foreach (dynamic section in doc.Sections)
                {
                    foreach (dynamic hf in section.Headers)
                        foreach (dynamic para in hf.Range.Paragraphs) ranges.Add(para.Range);
                    foreach (dynamic hf in section.Footers)
                        foreach (dynamic para in hf.Range.Paragraphs) ranges.Add(para.Range);
                }

                int total = ranges.Count, done = 0;
                foreach (dynamic range in ranges)
                {
                    string original = (range.Text ?? "")
                        .ToString().TrimEnd('\r','\a','\x0D','\x07');
                    if (!string.IsNullOrWhiteSpace(original))
                    {
                        string translated = bridge
                            .TranslateAsync(original, fromCode, toCode)
                            .GetAwaiter().GetResult();
                        if (!string.IsNullOrEmpty(translated) && translated != original)
                            range.Text = translated;
                    }
                    done++;
                    progress?.Report((done, total));
                }
                doc.Save();
            }
            finally
            {
                if (doc     is not null) { try { doc.Close(false);     } catch { } Marshal.ReleaseComObject(doc);     }
                if (wordApp is not null) { try { wordApp.Quit(false);  } catch { } Marshal.ReleaseComObject(wordApp); }
            }
        });
        return tmp;
    }

    // ── OpenXML fallback ──────────────────────────────────────────────────────
    private static async Task<string> TranslateViaOpenXmlAsync(
        string srcPath, string fromCode, string toCode,
        TranslatorBridge bridge,
        IProgress<(int Done, int Total)>? progress)
    {
        var tmp = Path.Combine(Path.GetTempPath(), $"doctranslate_{Guid.NewGuid():N}.docx");
        File.Copy(srcPath, tmp, overwrite: true);

        using var pkg = DocumentFormat.OpenXml.Packaging.WordprocessingDocument
            .Open(tmp, isEditable: true);
        var body = pkg.MainDocumentPart?.Document.Body;
        if (body is null) return tmp;

        var allParas = new List<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
        allParas.AddRange(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>());
        foreach (var part in pkg.MainDocumentPart!.HeaderParts)
            allParas.AddRange(part.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>());
        foreach (var part in pkg.MainDocumentPart.FooterParts)
            allParas.AddRange(part.Footer.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>());

        int total = allParas.Count, done = 0;
        foreach (var para in allParas)
        {
            var runs = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().ToList();
            if (runs.Any())
            {
                var fullText = string.Concat(runs.SelectMany(r =>
                    r.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                     .Select(t => t.Text)));
                if (!string.IsNullOrWhiteSpace(fullText))
                {
                    var translated = await bridge.TranslateAsync(fullText, fromCode, toCode);
                    if (!string.IsNullOrEmpty(translated) && translated != fullText)
                    {
                        bool first = true;
                        foreach (var run in runs)
                            foreach (var t in run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList())
                            { if (first) { t.Text = translated; first = false; } else t.Text = ""; }
                    }
                }
            }
            done++;
            progress?.Report((done, total));
        }

        // OCR every embedded image and insert a translated caption paragraph after it.
        await OcrAndInsertDocxImagesAsync(pkg, fromCode, toCode, bridge);

        pkg.MainDocumentPart.Document.Save();
        return tmp;
    }

    private static async Task OcrAndInsertDocxImagesAsync(
        DocumentFormat.OpenXml.Packaging.WordprocessingDocument pkg,
        string fromCode, string toCode, TranslatorBridge bridge)
    {
        var mainPart = pkg.MainDocumentPart!;

        // Find all Drawing elements that reference an image part via blip r:embed.
        var drawings = mainPart.Document.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>()
            .ToList();

        foreach (var drawing in drawings)
        {
            // Get the relationship ID of the embedded image.
            var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            if (blip?.Embed?.Value is not { } relId) continue;

            if (mainPart.GetPartById(relId) is not DocumentFormat.OpenXml.Packaging.ImagePart imgPart)
                continue;

            byte[] imgBytes;
            using (var s = imgPart.GetStream()) imgBytes = ReadAllBytes(s);

            var lines = await WinOcrEngine.RecognizeBytesAsync(imgBytes, fromCode);
            if (lines.Count == 0) continue;

            var ocrText = string.Join(" ", lines.Select(l => l.Text));
            var translated = await bridge.TranslateAsync(ocrText, fromCode, toCode);
            if (string.IsNullOrWhiteSpace(translated)) continue;

            // Insert a new paragraph with the translated text immediately after
            // the paragraph containing this drawing.
            var parentPara = drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                                    .FirstOrDefault();
            if (parentPara is null) continue;

            var newPara = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                        new DocumentFormat.OpenXml.Wordprocessing.Italic(),
                        new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "555555" }),
                    new DocumentFormat.OpenXml.Wordprocessing.Text(translated)
                        { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }
                ));
            parentPara.InsertAfterSelf(newPara);
        }
    }

    private static byte[] ReadAllBytes(Stream s)
    {
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }
}
