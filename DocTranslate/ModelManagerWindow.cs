using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DocTranslate;

/// <summary>
/// Lets the user browse all available Argos language models,
/// see which are installed, and download new ones.
/// Shows pivot availability automatically.
/// </summary>
public class ModelManagerWindow : Window
{
    private readonly TranslatorBridge _bridge;
    private StackPanel?               _listPanel;
    private TextBlock?                _statusLbl;

    public ModelManagerWindow(TranslatorBridge bridge)
    {
        _bridge = bridge;
        Title   = "Language Model Manager";
        Width   = 540;
        Height  = 560;
        Left    = 340;
        Top     = 180;
        Background        = new SolidColorBrush(Color.FromRgb(0x0F, 0x11, 0x17));
        FontFamily        = new FontFamily("Segoe UI");
        WindowStartupLocation = WindowStartupLocation.Manual;
        Loaded += async (_, _) => await LoadModelsAsync();
        BuildShell();
    }

    private void BuildShell()
    {
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(64) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(36) });
        Content = root;

        // Header
        var hdr = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(0x1A, 0x1D, 0x27)),
            Padding    = new Thickness(20, 0, 20, 0)
        };
        Grid.SetRow(hdr, 0);
        root.Children.Add(hdr);

        var hdrStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
        hdr.Child = hdrStack;
        hdrStack.Children.Add(new TextBlock
        {
            Text       = "Language Model Manager",
            FontSize   = 16, FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0xE8, 0xEA, 0xF6))
        });
        hdrStack.Children.Add(new TextBlock
        {
            Text       = "Models download once (~100MB) and work offline forever.",
            FontSize   = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(0x7B, 0x80, 0xA0))
        });

        // Scrollable list
        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Background = new SolidColorBrush(Color.FromRgb(0x22, 0x26, 0x3A)),
            Margin     = new Thickness(12, 8, 12, 0)
        };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        _listPanel = new StackPanel { Margin = new Thickness(0) };
        scroll.Content = _listPanel;

        // Status bar
        var statusBar = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(0x1A, 0x1D, 0x27)),
            Padding    = new Thickness(16, 0, 0, 0)
        };
        Grid.SetRow(statusBar, 2);
        root.Children.Add(statusBar);

        _statusLbl = new TextBlock
        {
            Text       = "Loading available packages…",
            FontSize   = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(0x7B, 0x80, 0xA0)),
            VerticalAlignment = VerticalAlignment.Center
        };
        statusBar.Child = _statusLbl;
    }

    private async Task LoadModelsAsync()
    {
        SetStatus("Fetching installed models…", "#7B80A0");
        _listPanel!.Children.Clear();

        List<(string From, string To)> installed;
        List<LangPackage> available;

        try
        {
            installed = await _bridge.GetInstalledPairsAsync();
        }
        catch
        {
            SetStatus("❌  Could not contact translation engine.", "#FF6B6B");
            return;
        }

        try
        {
            SetStatus("Fetching available packages from index…", "#FF9F43");
            available = await _bridge.GetAvailablePackagesAsync();
        }
        catch
        {
            available = new List<LangPackage>();
        }

        if (available.Count == 0)
        {
            // Show installed only
            foreach (var pair in installed)
            {
                AddRow($"en → {pair.To}", pair.From, pair.To, installed, true, null);
            }
            SetStatus($"  {installed.Count} model(s) installed. No internet to fetch available list.",
                      "#FF9F43");
            return;
        }

        // Sort: installed first, then alphabetical by to_name
        var sorted = available
            .OrderBy(p => !installed.Any(i => i.From == p.FromCode && i.To == p.ToCode))
            .ThenBy(p => p.ToName)
            .ToList();

        foreach (var pkg in sorted)
        {
            bool isInst = installed.Any(i => i.From == pkg.FromCode && i.To == pkg.ToCode);
            bool pivot  = !isInst && pkg.FromCode != "en" &&
                          installed.Any(i => i.From == pkg.FromCode && i.To == "en") &&
                          installed.Any(i => i.From == "en"          && i.To == pkg.ToCode);
            AddRow($"{pkg.FromName} → {pkg.ToName}", pkg.FromCode, pkg.ToCode,
                   installed, isInst, pivot ? "pivot" : null);
        }

        int usable = available.Count(p =>
        {
            bool direct = installed.Any(i => i.From == p.FromCode && i.To == p.ToCode);
            bool piv    = !direct && p.FromCode != "en" &&
                          installed.Any(i => i.From == p.FromCode && i.To == "en") &&
                          installed.Any(i => i.From == "en"         && i.To == p.ToCode);
            return direct || piv;
        });

        SetStatus($"  {installed.Count} installed · {usable}/{available.Count} usable (incl. pivoting)",
                  "#43D9AD");
    }

    private void AddRow(string label, string fromCode, string toCode,
                        List<(string From, string To)> installed,
                        bool isInstalled, string? tag)
    {
        var row = new Border
        {
            Padding         = new Thickness(12, 6, 12, 6),
            BorderBrush     = new SolidColorBrush(Color.FromRgb(0x2E, 0x32, 0x50)),
            BorderThickness = new Thickness(0, 0, 0, 1)
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        row.Child = grid;

        // Label
        var lbl = new StackPanel { Orientation = Orientation.Horizontal,
                                    VerticalAlignment = VerticalAlignment.Center };
        Grid.SetColumn(lbl, 0);
        grid.Children.Add(lbl);

        lbl.Children.Add(new TextBlock
        {
            Text       = label,
            FontSize   = 12,
            Foreground = new SolidColorBrush(Color.FromRgb(0xE8, 0xEA, 0xF6)),
            VerticalAlignment = VerticalAlignment.Center
        });

        if (tag == "pivot")
        {
            lbl.Children.Add(new Border
            {
                Background   = new SolidColorBrush(Color.FromRgb(0x1A, 0x2A, 0x1A)),
                CornerRadius = new CornerRadius(4),
                Margin       = new Thickness(8, 0, 0, 0),
                Padding      = new Thickness(6, 1, 6, 1),
                Child        = new TextBlock
                {
                    Text       = "pivot",
                    FontSize   = 9,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x43, 0xD9, 0xAD))
                }
            });
        }

        // Code badge
        var codeLbl = new TextBlock
        {
            Text       = $"{fromCode}→{toCode}",
            FontSize   = 10,
            Foreground = new SolidColorBrush(Color.FromRgb(0x7B, 0x80, 0xA0)),
            VerticalAlignment = VerticalAlignment.Center,
            Margin     = new Thickness(0, 0, 12, 0)
        };
        Grid.SetColumn(codeLbl, 1);
        grid.Children.Add(codeLbl);

        // Action button
        var btn = new Button
        {
            Height          = 28,
            FontSize        = 11,
            BorderThickness = new Thickness(0),
            Cursor          = System.Windows.Input.Cursors.Hand,
        };
        Grid.SetColumn(btn, 2);
        grid.Children.Add(btn);

        if (isInstalled)
        {
            btn.Content    = "✔  Installed";
            btn.Background = new SolidColorBrush(Color.FromRgb(0x43, 0xD9, 0xAD));
            btn.Foreground = new SolidColorBrush(Color.FromRgb(0x0A, 0x1A, 0x14));
            btn.IsEnabled  = false;
        }
        else
        {
            btn.Content    = "Download";
            btn.Background = new SolidColorBrush(Color.FromRgb(0x6C, 0x63, 0xFF));
            btn.Foreground = Brushes.White;
            btn.Click     += async (_, _) =>
            {
                btn.IsEnabled = false;
                btn.Content   = "Downloading…";
                SetStatus($"Downloading {label}…", "#FF9F43");

                bool ok = await _bridge.InstallPairAsync(fromCode, toCode);

                if (ok)
                {
                    btn.Content    = "✔  Installed";
                    btn.Background = new SolidColorBrush(Color.FromRgb(0x43, 0xD9, 0xAD));
                    btn.Foreground = new SolidColorBrush(Color.FromRgb(0x0A, 0x1A, 0x14));
                    SetStatus($"✅  {label} installed successfully!", "#43D9AD");
                    await LoadModelsAsync();   // refresh pivot labels
                }
                else
                {
                    btn.IsEnabled = true;
                    btn.Content   = "Retry";
                    btn.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0x6B, 0x6B));
                    SetStatus("❌  Download failed. Check your internet connection.", "#FF6B6B");
                }
            };
        }

        _listPanel!.Children.Add(row);
    }

    private void SetStatus(string text, string hexColor)
    {
        Dispatcher.Invoke(() =>
        {
            _statusLbl!.Text = text;
            _statusLbl.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString(hexColor));
        });
    }
}
