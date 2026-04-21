using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace PPEditer.Dialogs;

public partial class LicenseDialog : Window
{
    public LicenseDialog()
    {
        InitializeComponent();
        ContentArea.Content = BuildContent();
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();

    // ── Content ───────────────────────────────────────────────────────

    private static StackPanel BuildContent()
    {
        var p = new StackPanel();

        // ── 이 소프트웨어 ──────────────────────────────────────────────
        Add(p, H1("이 소프트웨어"));
        Add(p, LibCard(
            name:      "PPEditer",
            version:   "1.0.0",
            copyright: "Copyright © 2026 Noh JinMoon (HANDTECH — 상상공작소)",
            license:   "MIT License",
            links:     [
                ("GitHub", "https://github.com/elrang3843/PPEditer"),
            ]
        ));

        // ── 사용된 오픈소스 라이브러리 ────────────────────────────────
        Add(p, H1("사용된 오픈소스 라이브러리"));

        Add(p, LibCard(
            name:      ".NET 8 / WPF",
            version:   "8.0",
            copyright: "Copyright © Microsoft Corporation",
            license:   "MIT License",
            links:     [
                ("GitHub",  "https://github.com/dotnet/wpf"),
                ("License", "https://github.com/dotnet/wpf/blob/main/LICENSE.TXT"),
            ]
        ));

        Add(p, LibCard(
            name:      "DocumentFormat.OpenXml SDK",
            version:   "3.2.0",
            copyright: "Copyright © Microsoft Corporation",
            license:   "MIT License",
            links:     [
                ("NuGet",   "https://www.nuget.org/packages/DocumentFormat.OpenXml"),
                ("GitHub",  "https://github.com/dotnet/Open-XML-SDK"),
                ("License", "https://github.com/dotnet/Open-XML-SDK/blob/main/LICENSE"),
            ]
        ));

        Add(p, LibCard(
            name:      "PDFsharp",
            version:   "6.1.1",
            copyright: "Copyright © empira Software GmbH",
            license:   "MIT License",
            links:     [
                ("NuGet",   "https://www.nuget.org/packages/PDFsharp"),
                ("GitHub",  "https://github.com/empira/PDFsharp"),
                ("License", "https://github.com/empira/PDFsharp/blob/master/LICENSE"),
            ]
        ));

        // ── MIT License 전문 ──────────────────────────────────────────
        Add(p, H1("MIT License 전문 (PPEditer)"));
        Add(p, LicenseText(MitLicenseText));

        return p;
    }

    // ── UI builders ───────────────────────────────────────────────────

    private static void Add(StackPanel p, UIElement el) => p.Children.Add(el);

    private static UIElement H1(string t)
    {
        var sp = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };
        sp.Children.Add(new TextBlock
        {
            Text       = t,
            FontSize   = 15,
            FontWeight = FontWeights.Bold,
            Margin     = new Thickness(0, 0, 0, 5),
        });
        sp.Children.Add(new Separator());
        return sp;
    }

    private static Border LibCard(string name, string version, string copyright,
                                   string license, (string label, string url)[] links)
    {
        var card = new Border
        {
            BorderThickness  = new Thickness(1),
            BorderBrush      = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)),
            CornerRadius     = new CornerRadius(5),
            Padding          = new Thickness(14, 10, 14, 10),
            Margin           = new Thickness(0, 0, 0, 10),
        };

        var inner = new StackPanel();

        // 이름 + 버전
        var nameRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 3) };
        nameRow.Children.Add(new TextBlock
        {
            Text = name, FontWeight = FontWeights.Bold, FontSize = 14,
            VerticalAlignment = VerticalAlignment.Center,
        });
        nameRow.Children.Add(new Border
        {
            Background    = new SolidColorBrush(Color.FromRgb(0x27, 0x7C, 0xB0)),
            CornerRadius  = new CornerRadius(3),
            Padding       = new Thickness(6, 1, 6, 1),
            Margin        = new Thickness(8, 1, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child         = new TextBlock
            {
                Text       = "v" + version,
                FontSize   = 11,
                Foreground = Brushes.White,
            },
        });
        nameRow.Children.Add(new Border
        {
            Background    = new SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0x32)),
            CornerRadius  = new CornerRadius(3),
            Padding       = new Thickness(6, 1, 6, 1),
            Margin        = new Thickness(6, 1, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child         = new TextBlock
            {
                Text       = license,
                FontSize   = 11,
                Foreground = Brushes.White,
            },
        });
        inner.Children.Add(nameRow);

        // 저작권
        inner.Children.Add(new TextBlock
        {
            Text       = copyright,
            FontSize   = 12,
            Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)),
            Margin     = new Thickness(0, 2, 0, 6),
        });

        // 링크
        var linkRow = new TextBlock { Margin = new Thickness(0, 0, 0, 0) };
        bool first = true;
        foreach (var (label, url) in links)
        {
            if (!first)
                linkRow.Inlines.Add(new Run("  "));
            linkRow.Inlines.Add(LinkButton(label, url));
            first = false;
        }
        inner.Children.Add(linkRow);

        card.Child = inner;
        return card;
    }

    private static Inline LinkButton(string label, string url)
    {
        var hl = new Hyperlink(new Run($"🔗 {label}"))
        {
            NavigateUri = new Uri(url),
            Foreground  = new SolidColorBrush(Color.FromRgb(0x19, 0x76, 0xD2)),
        };
        hl.RequestNavigate += (_, e) =>
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        return hl;
    }

    private static Border LicenseText(string text)
    {
        return new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(0),
            Margin          = new Thickness(0, 0, 0, 20),
            Child           = new TextBox
            {
                Text             = text,
                IsReadOnly       = true,
                TextWrapping     = TextWrapping.Wrap,
                FontFamily       = new FontFamily("Consolas, Courier New"),
                FontSize         = 11,
                BorderThickness  = new Thickness(0),
                Padding          = new Thickness(12, 10, 12, 10),
                MaxHeight        = 220,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                AcceptsReturn    = true,
            },
        };
    }

    // ── License text ─────────────────────────────────────────────────

    private const string MitLicenseText =
        "MIT License\r\n\r\n" +
        "Copyright (c) 2026 Noh JinMoon (HANDTECH — 상상공작소)\r\n\r\n" +
        "Permission is hereby granted, free of charge, to any person obtaining a copy\r\n" +
        "of this software and associated documentation files (the \"Software\"), to deal\r\n" +
        "in the Software without restriction, including without limitation the rights\r\n" +
        "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell\r\n" +
        "copies of the Software, and to permit persons to whom the Software is\r\n" +
        "furnished to do so, subject to the following conditions:\r\n\r\n" +
        "The above copyright notice and this permission notice shall be included in all\r\n" +
        "copies or substantial portions of the Software.\r\n\r\n" +
        "THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR\r\n" +
        "IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,\r\n" +
        "FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE\r\n" +
        "AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER\r\n" +
        "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,\r\n" +
        "OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE\r\n" +
        "SOFTWARE.";
}
