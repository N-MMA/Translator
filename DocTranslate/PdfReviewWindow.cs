using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DocTranslate;

public class PdfReviewWindow : Window
{
    private readonly List<PdfTextBlock> _blocks;
    private readonly List<TextBox>      _textBoxes = new();
    public  List<PdfTextBlock>?         ConfirmedBlocks { get; private set; }

    public PdfReviewWindow(List<PdfTextBlock> blocks, string srcPath, string langName)
    {
        _blocks = blocks;
        Title   = "Review & Edit Translations — PDF";
        Width=800; Height=660; Left=220; Top=120;
        Background = new SolidColorBrush(Color.FromRgb(0x0F,0x11,0x17));
        FontFamily = new FontFamily("Segoe UI");
        WindowStartupLocation = WindowStartupLocation.Manual;
        BuildUI(langName);
    }

    private void BuildUI(string langName)
    {
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition{Height=new GridLength(56)});
        root.RowDefinitions.Add(new RowDefinition{Height=new GridLength(30)});
        root.RowDefinitions.Add(new RowDefinition{Height=new GridLength(1,GridUnitType.Star)});
        root.RowDefinitions.Add(new RowDefinition{Height=new GridLength(60)});
        Content = root;

        // Header
        var hdr = new Border{Background=new SolidColorBrush(Color.FromRgb(0x1A,0x1D,0x27)),Padding=new Thickness(20,0,20,0)};
        Grid.SetRow(hdr,0); root.Children.Add(hdr);
        var hs = new StackPanel{Orientation=Orientation.Horizontal,VerticalAlignment=VerticalAlignment.Center};
        hdr.Child = hs;
        hs.Children.Add(new TextBlock{Text="✏️  Review & Edit Translations",FontSize=16,FontWeight=FontWeights.Bold,Foreground=new SolidColorBrush(Color.FromRgb(0xE8,0xEA,0xF6)),VerticalAlignment=VerticalAlignment.Center});
        hs.Children.Add(new TextBlock{Text="   Edit before generating PDF.",FontSize=11,Foreground=new SolidColorBrush(Color.FromRgb(0x7B,0x80,0xA0)),VerticalAlignment=VerticalAlignment.Center});

        // Info bar
        var info = new Border{Background=new SolidColorBrush(Color.FromRgb(0x22,0x26,0x3A)),Padding=new Thickness(16,0,0,0)};
        Grid.SetRow(info,1); root.Children.Add(info);
        info.Child = new TextBlock{Text=$"  {_blocks.Count} blocks · Target: {langName}   Tip: clear a block to skip it",FontSize=11,Foreground=new SolidColorBrush(Color.FromRgb(0x7B,0x80,0xA0)),VerticalAlignment=VerticalAlignment.Center};

        // Scroll
        var scroll = new ScrollViewer{VerticalScrollBarVisibility=ScrollBarVisibility.Auto,HorizontalScrollBarVisibility=ScrollBarVisibility.Disabled,Background=new SolidColorBrush(Color.FromRgb(0x0F,0x11,0x17))};
        Grid.SetRow(scroll,2); root.Children.Add(scroll);
        var stack = new StackPanel();
        scroll.Content = stack;

        int curPage = -1;
        foreach(var block in _blocks)
        {
            if(block.PageIndex != curPage)
            {
                curPage = block.PageIndex;
                var div = new Border{Background=new SolidColorBrush(Color.FromRgb(0x6C,0x63,0xFF)),Height=28,Margin=new Thickness(0,12,0,0),Padding=new Thickness(16,0,0,0)};
                div.Child = new TextBlock{Text=$"Page {curPage+1}",FontSize=11,FontWeight=FontWeights.Bold,Foreground=Brushes.White,VerticalAlignment=VerticalAlignment.Center};
                stack.Children.Add(div);
            }
            var card = new Border{Background=new SolidColorBrush(Color.FromRgb(0x22,0x26,0x3A)),CornerRadius=new CornerRadius(8),Margin=new Thickness(12,4,12,0),Padding=new Thickness(10,8,10,10),BorderBrush=new SolidColorBrush(Color.FromRgb(0x2E,0x32,0x50)),BorderThickness=new Thickness(1)};
            stack.Children.Add(card);
            var cs = new StackPanel(); card.Child = cs;
            cs.Children.Add(new TextBlock{Text="Original:",FontSize=10,FontWeight=FontWeights.Bold,Foreground=new SolidColorBrush(Color.FromRgb(0x7B,0x80,0xA0)),Margin=new Thickness(0,0,0,2)});
            cs.Children.Add(new TextBlock{Text=block.Original.Length>200?block.Original[..200]+"…":block.Original,FontSize=10,Foreground=new SolidColorBrush(Color.FromRgb(0x7B,0x80,0xA0)),TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,6)});
            cs.Children.Add(new TextBlock{Text="Translation:",FontSize=10,FontWeight=FontWeights.Bold,Foreground=new SolidColorBrush(Color.FromRgb(0xE8,0xEA,0xF6)),Margin=new Thickness(0,0,0,2)});
            int lines=Math.Max(2,Math.Min(6,block.Translated.Length/60+1));
            var tb = new TextBox{Text=block.Translated,MinHeight=lines*22,MaxHeight=140,FontSize=12,Background=new SolidColorBrush(Color.FromRgb(0x1A,0x1D,0x27)),Foreground=new SolidColorBrush(Color.FromRgb(0xE8,0xEA,0xF6)),BorderBrush=new SolidColorBrush(Color.FromRgb(0x2E,0x32,0x50)),BorderThickness=new Thickness(1),Padding=new Thickness(8,6,8,6),TextWrapping=TextWrapping.Wrap,AcceptsReturn=true,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,CaretBrush=Brushes.White};
            cs.Children.Add(tb);
            _textBoxes.Add(tb);
        }

        // Footer
        var footer = new Border{Background=new SolidColorBrush(Color.FromRgb(0x1A,0x1D,0x27)),Padding=new Thickness(16,0,16,0)};
        Grid.SetRow(footer,3); root.Children.Add(footer);
        var fg = new Grid();
        fg.ColumnDefinitions.Add(new ColumnDefinition{Width=new GridLength(1,GridUnitType.Star)});
        fg.ColumnDefinitions.Add(new ColumnDefinition{Width=GridLength.Auto});
        fg.ColumnDefinitions.Add(new ColumnDefinition{Width=GridLength.Auto});
        footer.Child = fg;
        var tip = new TextBlock{Text="Clear a block to skip it.",FontSize=10,Foreground=new SolidColorBrush(Color.FromRgb(0x7B,0x80,0xA0)),VerticalAlignment=VerticalAlignment.Center};
        Grid.SetColumn(tip,0); fg.Children.Add(tip);
        var cancelBtn = MakeBtn("Cancel","#22263A","#7B80A0"); Grid.SetColumn(cancelBtn,1); cancelBtn.Margin=new Thickness(0,10,8,10); cancelBtn.Click+=(_,_)=>{DialogResult=false;Close();}; fg.Children.Add(cancelBtn);
        var genBtn    = MakeBtn("✦  Generate PDF","#43D9AD","#0A1A14"); Grid.SetColumn(genBtn,2); genBtn.Margin=new Thickness(0,10,0,10); genBtn.FontWeight=FontWeights.Bold; genBtn.Click+=OnConfirm; fg.Children.Add(genBtn);
    }

    private void OnConfirm(object sender, RoutedEventArgs e)
    {
        for(int i=0;i<_blocks.Count;i++) _blocks[i].Translated=_textBoxes[i].Text.Trim();
        ConfirmedBlocks=_blocks; DialogResult=true; Close();
    }

    private static Button MakeBtn(string text,string bg,string fg) =>
        new(){Content=text,Height=40,Padding=new Thickness(20,0,20,0),FontSize=13,
              Background=new SolidColorBrush((Color)ColorConverter.ConvertFromString(bg)),
              Foreground=new SolidColorBrush((Color)ColorConverter.ConvertFromString(fg)),
              BorderThickness=new Thickness(0),Cursor=System.Windows.Input.Cursors.Hand};
}
