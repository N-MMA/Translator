using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace DocTranslate;

public partial class MainWindow : Window
{
    private static readonly Dictionary<string, (string From, string To)> Languages = new()
    {
        ["🇵🇹 Portuguese"]           = ("en","pt"), ["🇪🇸 Spanish"]    = ("en","es"),
        ["🇫🇷 French"]               = ("en","fr"), ["🇩🇪 German"]     = ("en","de"),
        ["🇮🇹 Italian"]              = ("en","it"), ["🇳🇱 Dutch"]      = ("en","nl"),
        ["🇵🇱 Polish"]               = ("en","pl"), ["🇷🇺 Russian"]    = ("en","ru"),
        ["🇺🇦 Ukrainian"]            = ("en","uk"), ["🇸🇪 Swedish"]    = ("en","sv"),
        ["🇩🇰 Danish"]               = ("en","da"), ["🇳🇴 Norwegian"]  = ("en","nb"),
        ["🇫🇮 Finnish"]              = ("en","fi"), ["🇬🇷 Greek"]      = ("en","el"),
        ["🇹🇷 Turkish"]              = ("en","tr"), ["🇷🇴 Romanian"]   = ("en","ro"),
        ["🇨🇿 Czech"]                = ("en","cs"), ["🇸🇰 Slovak"]     = ("en","sk"),
        ["🇭🇺 Hungarian"]            = ("en","hu"), ["🇧🇬 Bulgarian"]  = ("en","bg"),
        ["🇮🇪 Irish"]                = ("en","ga"), ["🇸🇦 Arabic"]     = ("en","ar"),
        ["🇮🇱 Hebrew"]               = ("en","he"), ["🇮🇷 Persian"]    = ("en","fa"),
        ["🇮🇳 Hindi"]                = ("en","hi"), ["🇮🇳 Urdu"]       = ("en","ur"),
        ["🇯🇵 Japanese"]             = ("en","ja"), ["🇨🇳 Chinese"]    = ("en","zh"),
        ["🇰🇷 Korean"]               = ("en","ko"), ["🇹🇭 Thai"]       = ("en","th"),
        ["🇻🇳 Vietnamese"]           = ("en","vi"), ["🇮🇩 Indonesian"] = ("en","id"),
        ["🇲🇾 Malay"]                = ("en","ms"), ["🇵🇭 Tagalog"]    = ("en","tl"),
        ["🌍 Esperanto"]             = ("en","eo"),
    };

    private readonly TranslatorBridge _bridge = new();
    private string? _srcPath, _fileType, _outPath;
    private bool    _running;
    private List<PdfTextBlock>? _pdfBlocks;

    public MainWindow()
    {
        InitializeComponent();
        var srcLangs = new List<string> { "🇬🇧 English" };
        srcLangs.AddRange(Languages.Keys.Where(k => !k.Contains("English")));
        SrcLangBox.ItemsSource   = srcLangs;
        SrcLangBox.SelectedIndex = 0;
        TgtLangBox.ItemsSource   = Languages.Keys.ToList();
        TgtLangBox.SelectedIndex = 0;
        SrcLangBox.SelectionChanged += async (_, _) => await UpdateModelStatus();
        TgtLangBox.SelectionChanged += async (_, _) => await UpdateModelStatus();
        Loaded += OnLoaded;
        Closed += (_, _) => _bridge.Dispose();
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        SetBanner("Starting translation engine…", "#FF9F43");
        bool wordAvail = DocxEngine.IsWordAvailable();
        WordBadge.Text       = wordAvail ? "● Word COM" : "● Word: OpenXML";
        WordBadge.Foreground = new SolidColorBrush(wordAvail
            ? Color.FromRgb(0x43,0xD9,0xAD)
            : Color.FromRgb(0xFF,0x9F,0x43));

        bool pptAvail = PptxEngine.IsPowerPointAvailable();
        PptBadge.Text       = pptAvail ? "● PowerPoint COM" : "● PPT: OpenXML";
        PptBadge.Foreground = new SolidColorBrush(pptAvail
            ? Color.FromRgb(0x43,0xD9,0xAD)
            : Color.FromRgb(0xFF,0x9F,0x43));

        string pythonExe  = FindPython();
        string scriptPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "translator", "translator.py");

        bool ok = await _bridge.StartAsync(pythonExe, scriptPath);
        SetBanner(ok
            ? "✔  Translation engine ready — fully offline after model download"
            : "❌  Could not start Python. Ensure Python 3.10+ is installed and in PATH.",
            ok ? "#43D9AD" : "#FF6B6B");
        if (ok) await UpdateModelStatus();
    }

    private static string FindPython()
    {
        foreach (var c in new[]{"python","python3","py"})
        {
            try
            {
                var p = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName=c, Arguments="--version", UseShellExecute=false,
                  RedirectStandardOutput=true, CreateNoWindow=true });
                p?.WaitForExit(2000);
                if (p?.ExitCode == 0) return c;
            }
            catch { }
        }
        return "python";
    }

    private async Task UpdateModelStatus()
    {
        if (!_bridge.IsReady) return;
        var srcName = SrcLangBox.SelectedItem?.ToString() ?? "";
        var tgtName = TgtLangBox.SelectedItem?.ToString() ?? "";
        string fromCode = srcName.Contains("English") ? "en"
            : Languages.TryGetValue(srcName, out var sv) ? sv.To : "en";
        string toCode = Languages.TryGetValue(tgtName, out var tv) ? tv.To : "pt";

        if (fromCode == toCode)
        { ModelStatus.Text="  ⚠  Source and target language are the same"; ModelStatus.Foreground=new SolidColorBrush(Color.FromRgb(0xFF,0x9F,0x43)); return; }

        var pairs  = await _bridge.GetInstalledPairsAsync();
        bool direct = pairs.Any(p => p.From==fromCode && p.To==toCode);
        bool pivot  = !direct && fromCode!="en" && toCode!="en"
                      && pairs.Any(p=>p.From==fromCode&&p.To=="en")
                      && pairs.Any(p=>p.From=="en"&&p.To==toCode);
        if (direct)
        { ModelStatus.Text=$"  ✔  Direct model — {srcName} → {tgtName}"; ModelStatus.Foreground=new SolidColorBrush(Color.FromRgb(0x43,0xD9,0xAD)); }
        else if (pivot)
        { ModelStatus.Text=$"  ✔  Pivot via English — {srcName} → {tgtName}"; ModelStatus.Foreground=new SolidColorBrush(Color.FromRgb(0x43,0xD9,0xAD)); }
        else
        { ModelStatus.Text=$"  ⚠  Model not installed for {tgtName} — click Manage Models"; ModelStatus.Foreground=new SolidColorBrush(Color.FromRgb(0xFF,0x9F,0x43)); }
    }

    private async void OnBrowse(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        { Filter="Supported Documents|*.docx;*.pdf;*.pptx|Word Documents|*.docx|PDF Files|*.pdf|PowerPoint Presentations|*.pptx" };
        if (dlg.ShowDialog() != true) return;
        _srcPath=dlg.FileName; _outPath=null; _pdfBlocks=null;
        SaveBtn.IsEnabled=false; ProgressBar.Value=0;
        string ext=Path.GetExtension(_srcPath).ToLower();
        _fileType = ext==".pdf" ? "pdf" : ext==".pptx" ? "pptx" : "docx";
        FileLabel.Text=Path.GetFileName(_srcPath);
        FileLabel.Foreground=new SolidColorBrush(Color.FromRgb(0xE8,0xEA,0xF6));
        TypeBadge.Visibility=Visibility.Visible;
        long kb=new FileInfo(_srcPath).Length/1024;
        if (_fileType=="pdf")
        {
            TypeBadge.Background=new SolidColorBrush(Color.FromRgb(0xFF,0x9F,0x43));
            TypeBadgeText.Text="  PDF  "; TypeBadgeText.Foreground=new SolidColorBrush(Color.FromRgb(0x1A,0x0A,0x00));
            FileInfo.Text=$"  ✔  {kb} KB — pages flattened to image, text overlaid";
        }
        else if (_fileType=="pptx")
        {
            TypeBadge.Background=new SolidColorBrush(Color.FromRgb(0xD0,0x41,0x27));
            TypeBadgeText.Text="  PPTX  "; TypeBadgeText.Foreground=new SolidColorBrush(Colors.White);
            FileInfo.Text=$"  ✔  {kb} KB — {(PptxEngine.IsPowerPointAvailable()?"PowerPoint COM — full fidelity":"OpenXML fallback")}";
        }
        else
        {
            TypeBadge.Background=new SolidColorBrush(Color.FromRgb(0x6C,0x63,0xFF));
            TypeBadgeText.Text="  DOCX  "; TypeBadgeText.Foreground=new SolidColorBrush(Colors.White);
            FileInfo.Text=$"  ✔  {kb} KB — {(DocxEngine.IsWordAvailable()?"Word COM — full fidelity":"OpenXML fallback")}";
        }
        StatusText.Text="File loaded. Choose languages and click Translate.";
        await UpdateModelStatus();
    }

    private async void OnTranslate(object sender, RoutedEventArgs e)
    {
        if (_srcPath is null){MessageBox.Show("Please select a file first.");return;}
        if (_running) return;
        var srcName=SrcLangBox.SelectedItem?.ToString()??"";
        var tgtName=TgtLangBox.SelectedItem?.ToString()??"";
        string fromCode=srcName.Contains("English")?"en":Languages.TryGetValue(srcName,out var sv)?sv.To:"en";
        string toCode=Languages.TryGetValue(tgtName,out var tv)?tv.To:"pt";
        if (fromCode==toCode){MessageBox.Show("Source and target are the same.");return;}
        _running=true; _outPath=null;
        SaveBtn.IsEnabled=false; TranslateBtn.IsEnabled=false; ProgressBar.Value=0;
        StatusText.Text=$"Translating to {tgtName}…";
        var progress=new Progress<(int Done,int Total)>(p=>{
            ProgressBar.Value=p.Total>0?(double)p.Done/p.Total*100:0;
            StatusText.Text=$"Translating… {p.Done}/{p.Total}";
        });
        try
        {
            if      (_fileType=="pdf")  await TranslatePdfAsync(fromCode,toCode,tgtName,progress);
            else if (_fileType=="pptx") await TranslatePptxAsync(fromCode,toCode,tgtName,progress);
            else                        await TranslateDocxAsync(fromCode,toCode,tgtName,progress);
        }
        catch(Exception ex)
        {
            MessageBox.Show($"Error:\n{ex.Message}","Error",MessageBoxButton.OK,MessageBoxImage.Error);
            StatusText.Text="❌  Error during translation.";
        }
        finally { _running=false; TranslateBtn.IsEnabled=true; }
    }

    private async Task TranslatePptxAsync(string from,string to,string tgtName,
        IProgress<(int,int)> progress)
    {
        _outPath=await PptxEngine.TranslateAsync(_srcPath!,from,to,_bridge,progress);
        ProgressBar.Value=100; SaveBtn.IsEnabled=true;
        StatusText.Text=$"✅  Translation to {tgtName} complete — click Save.";
    }

    private async Task TranslateDocxAsync(string from,string to,string tgtName,
        IProgress<(int,int)> progress)
    {
        _outPath=await DocxEngine.TranslateAsync(_srcPath!,from,to,_bridge,progress);
        ProgressBar.Value=100; SaveBtn.IsEnabled=true;
        StatusText.Text=$"✅  Translation to {tgtName} complete — click Save.";
    }

    private async Task TranslatePdfAsync(string from,string to,string tgtName,
        IProgress<(int,int)> progress)
    {
        StatusText.Text="Extracting text blocks from PDF…";
        _pdfBlocks=await Task.Run(()=>PdfEngine.ExtractBlocks(_srcPath!));

        // OCR fallback: pages with no extractable text (image-only / scanned tables)
        // are rasterised by Python/PyMuPDF, read by EasyOCR and pre-translated.
        StatusText.Text="Scanning for image-only pages (OCR)…";
        var ocrProgress=new Progress<(int Done,int Total)>(p=>
            StatusText.Text=$"OCR: page {p.Done}/{p.Total}…");
        await PdfEngine.FillWithOcrAsync(_srcPath!,_pdfBlocks,_bridge,from,to,ocrProgress);

        // Translate text-extracted blocks (OCR blocks already carry a translation).
        int total=_pdfBlocks.Count;
        for(int i=0;i<_pdfBlocks.Count;i++)
        {
            if(string.IsNullOrEmpty(_pdfBlocks[i].Translated))
                _pdfBlocks[i].Translated=await _bridge.TranslateAsync(
                    _pdfBlocks[i].Original,from,to);
            progress.Report((i+1,total));
        }
        ProgressBar.Value=100;
        var reviewWin=new PdfReviewWindow(_pdfBlocks,_srcPath!,tgtName){Owner=this};
        if(reviewWin.ShowDialog()==true && reviewWin.ConfirmedBlocks is {} confirmed)
        {
            StatusText.Text="Rendering final PDF…";
            _outPath=await Task.Run(()=>PdfEngine.RenderOverlay(_srcPath!,confirmed));
            SaveBtn.IsEnabled=true;
            StatusText.Text=$"✅  Translation to {tgtName} complete — click Save.";
        }
        else StatusText.Text="Review cancelled.";
    }

    private void OnSave(object sender,RoutedEventArgs e)
    {
        if(_outPath is null) return;
        string stem=Path.GetFileNameWithoutExtension(_srcPath??"document");
        string lang=(TgtLangBox.SelectedItem?.ToString()??"").Split(' ').Last().Trim();
        string ext=_fileType=="pdf"?".pdf":_fileType=="pptx"?".pptx":".docx";
        var dlg=new SaveFileDialog
        { FileName=$"{stem}_{lang}{ext}", DefaultExt=ext,
          Filter=_fileType=="pdf"?"PDF Files|*.pdf":_fileType=="pptx"?"PowerPoint Presentations|*.pptx":"Word Documents|*.docx" };
        if(dlg.ShowDialog()!=true) return;
        try { File.Copy(_outPath,dlg.FileName,true); MessageBox.Show($"Saved:\n{dlg.FileName}","Saved"); }
        catch(Exception ex){ MessageBox.Show($"Save failed:\n{ex.Message}","Error",MessageBoxButton.OK,MessageBoxImage.Error); }
    }

    private void OnManageModels(object sender,RoutedEventArgs e)
    {
        new ModelManagerWindow(_bridge){Owner=this}.ShowDialog();
        _=UpdateModelStatus();
    }

    private void SetBanner(string text,string hex)
    {
        BannerText.Text=text;
        BannerText.Foreground=new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
    }
}
