using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Globalization;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocTranslate;

public class PdfTextBlock
{
    public int    PageIndex  { get; set; }
    public double X          { get; set; }
    public double Y          { get; set; }
    public double Width      { get; set; }
    public double Height     { get; set; }
    public string Original   { get; set; } = "";
    public string Translated { get; set; } = "";
    public double FontSize   { get; set; } = 10;
    public bool   Bold       { get; set; }
    public XColor Color      { get; set; } = XColors.Black;
}

/// <summary>
/// PDF engine: extract text blocks with position, translate (caller),
/// flatten pages to images and overlay translated text.
/// Table borders and images are captured in the flattened background.
/// </summary>
public static class PdfEngine
{
    public static List<PdfTextBlock> ExtractBlocks(string srcPath)
    {
        var blocks = new List<PdfTextBlock>();
        using var doc = PdfReader.Open(srcPath, PdfDocumentOpenMode.ReadOnly);
        for (int pi = 0; pi < doc.PageCount; pi++)
        {
            var page = doc.Pages[pi];
            blocks.AddRange(ParsePage(page, pi));
        }
        return blocks;
    }

    private static List<PdfTextBlock> ParsePage(PdfSharp.Pdf.PdfPage page, int pageIdx)
    {
        var blocks = new List<PdfTextBlock>();
        var ph = page.Height.Point;
        var pw = page.Width.Point;
        try
        {
            string raw = "";
            for (int i = 0; i < page.Contents.Elements.Count; i++)
            {
                var dict = page.Contents.Elements.GetDictionary(i);
                if (dict?.Stream is { } s)
                    raw += Encoding.Latin1.GetString(s.Value) + "\n";
            }
            blocks.AddRange(ParseStream(raw, pageIdx, pw, ph));
        }
        catch { }
        return blocks;
    }

    private static List<PdfTextBlock> ParseStream(
        string content, int pageIdx, double pw, double ph)
    {
        var blocks  = new List<PdfTextBlock>();
        var lines   = content.Split('\n');
        double cx = 72, cy = 72, fs = 10;
        bool inBT = false;
        var buf = new StringBuilder();

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (line == "BT") { inBT = true; buf.Clear(); continue; }
            if (line == "ET")
            {
                inBT = false;
                var txt = buf.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(txt))
                    blocks.Add(new PdfTextBlock {
                        PageIndex=pageIdx, X=cx, Y=ph-cy-fs,
                        Width=Math.Max(pw-cx-36,100), Height=fs*1.4,
                        Original=txt, FontSize=fs });
                buf.Clear(); continue;
            }
            if (!inBT) continue;

            if (line.EndsWith(" Tf"))
            {
                var p = line.Split(' ');
                if (p.Length >= 2 && double.TryParse(p[^2],
                    NumberStyles.Any, CultureInfo.InvariantCulture, out double f) && f > 0)
                    fs = f;
            }
            else if (line.EndsWith(" Td") || line.EndsWith(" TD"))
            {
                var p = line.Split(' ');
                if (p.Length >= 3)
                {
                    if (double.TryParse(p[^3], NumberStyles.Any, CultureInfo.InvariantCulture, out double dx)) cx += dx;
                    if (double.TryParse(p[^2], NumberStyles.Any, CultureInfo.InvariantCulture, out double dy)) cy -= dy;
                }
            }
            else if (line.EndsWith(" Tm"))
            {
                var p = line.Split(' ');
                if (p.Length >= 7)
                {
                    if (double.TryParse(p[^3], NumberStyles.Any, CultureInfo.InvariantCulture, out double tx)) cx = tx;
                    if (double.TryParse(p[^2], NumberStyles.Any, CultureInfo.InvariantCulture, out double ty)) cy = ty;
                }
            }
            else if (line.EndsWith(" Tj") || line.Contains("("))
            {
                int s = line.IndexOf('('), e = line.LastIndexOf(')');
                if (s >= 0 && e > s) buf.Append(line[(s+1)..e] + " ");
            }
            else if (line.Contains("TJ"))
            {
                bool ins = false;
                foreach (char c in line)
                {
                    if (c == '(' && !ins) { ins = true; continue; }
                    if (c == ')' && ins)  { ins = false; continue; }
                    if (ins) buf.Append(c);
                }
            }
        }
        return blocks;
    }

    public static string RenderOverlay(string srcPath, List<PdfTextBlock> blocks)
    {
        var outPath = Path.Combine(Path.GetTempPath(),
                                   $"doctranslate_pdf_{Guid.NewGuid():N}.pdf");

        // Group blocks by page
        var byPage = new Dictionary<int, List<PdfTextBlock>>();
        foreach (var b in blocks)
        {
            if (!string.IsNullOrWhiteSpace(b.Translated))
            {
                if (!byPage.ContainsKey(b.PageIndex)) byPage[b.PageIndex] = new();
                byPage[b.PageIndex].Add(b);
            }
        }

        using var srcDoc = PdfReader.Open(srcPath, PdfDocumentOpenMode.Import);
        using var outDoc = new PdfDocument();

        for (int pi = 0; pi < srcDoc.PageCount; pi++)
        {
            var srcPage = srcDoc.Pages[pi];
            double pw = srcPage.Width.Point, ph = srcPage.Height.Point;

            var outPage = outDoc.AddPage();
            outPage.Width  = srcPage.Width;
            outPage.Height = srcPage.Height;

            using var gfx = XGraphics.FromPdfPage(outPage);

            // Background: import source page as XPdfForm
            var xForm = XPdfForm.FromFile(srcPath);
            xForm.PageIndex = pi;
            gfx.DrawImage(xForm, new XRect(0, 0, pw, ph));

            // Overlay translated blocks
            if (byPage.TryGetValue(pi, out var pageBlocks))
            {
                foreach (var block in pageBlocks)
                {
                    if (string.IsNullOrWhiteSpace(block.Translated)) continue;
                    var rect = new XRect(block.X, block.Y, block.Width, block.Height);
                    gfx.DrawRectangle(XBrushes.White, rect);
                    var style = block.Bold ? XFontStyleEx.Bold : XFontStyleEx.Regular;
                    var font  = new XFont("Arial", block.FontSize, style);
                    gfx.DrawString(block.Translated, font,
                                   new XSolidBrush(block.Color),
                                   rect, XStringFormats.TopLeft);
                }
            }
        }

        outDoc.Save(outPath);
        return outPath;
    }
}
