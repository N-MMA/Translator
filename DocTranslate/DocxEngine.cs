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
        pkg.MainDocumentPart.Document.Save();
        return tmp;
    }
}
