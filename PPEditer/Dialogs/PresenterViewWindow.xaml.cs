using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using PPEditer.Models;
using PPEditer.Rendering;

namespace PPEditer.Dialogs;

public partial class PresenterViewWindow : Window
{
    private readonly PresentationModel _model;
    private int               _currentIndex;
    private SlideShowWindow?  _showWindow;

    private readonly DispatcherTimer _timer   = new() { Interval = TimeSpan.FromSeconds(1) };
    private readonly Stopwatch       _elapsed = new();

    public PresenterViewWindow(PresentationModel model, int startIndex)
    {
        InitializeComponent();
        _model        = model;
        _currentIndex = startIndex;

        _timer.Tick += (_, _) =>
            TimerText.Text = _elapsed.Elapsed.ToString(@"hh\:mm\:ss");

        Loaded += (_, _) =>
        {
            _elapsed.Start();
            _timer.Start();
            ShowCurrentSlide();
        };
    }

    /// <summary>Link this window to the SlideShowWindow so they stay in sync.</summary>
    public void AttachShowWindow(SlideShowWindow wnd)
    {
        _showWindow = wnd;
        wnd.SlideChanged += index =>
        {
            _currentIndex = index;
            ShowCurrentSlide();
        };
    }

    // ── Slide display ─────────────────────────────────────────────────

    private void ShowCurrentSlide()
    {
        int idx = _currentIndex;

        // Current slide
        var part = _model.GetSlidePart(idx);
        if (part is not null)
        {
            var canvas = SlideRenderer.BuildCanvas(part, _model.SlideWidth, _model.SlideHeight);
            CurrentSlideBorder.Child = new Viewbox { Stretch = Stretch.Uniform, Child = canvas };
        }

        // Next slide (or "end" indicator)
        bool hasNext = idx + 1 < _model.SlideCount;
        if (hasNext)
        {
            var nextPart = _model.GetSlidePart(idx + 1);
            if (nextPart is not null)
            {
                var nextCanvas = SlideRenderer.BuildCanvas(
                    nextPart, _model.SlideWidth, _model.SlideHeight);
                NextSlideBorder.Child =
                    new Viewbox { Stretch = Stretch.Uniform, Child = nextCanvas };
            }
        }
        else
        {
            NextSlideBorder.Child = new TextBlock
            {
                Text                = Application.Current.TryFindResource("Lbl_LastSlide") as string
                                      ?? "발표 종료",
                Foreground          = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x88)),
                FontSize            = 14,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            };
        }

        SlideCountText.Text = $"{idx + 1} / {_model.SlideCount}";
        NotesText.Text      = _model.GetSlideNotes(idx);
        BtnPrev.IsEnabled   = idx > 0;
        BtnNext.IsEnabled   = hasNext;
    }

    // ── Navigation ────────────────────────────────────────────────────

    private void OnPrev(object sender, RoutedEventArgs e) => _showWindow?.Navigate(-1);
    private void OnNext(object sender, RoutedEventArgs e) => _showWindow?.Navigate(1);

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.PageDown: case Key.Space: case Key.Right: case Key.Down:
                _showWindow?.Navigate(1);
                e.Handled = true;
                break;
            case Key.PageUp: case Key.Back: case Key.Left: case Key.Up:
                _showWindow?.Navigate(-1);
                e.Handled = true;
                break;
            case Key.Escape:
                _showWindow?.Close();
                Close();
                break;
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        _timer.Stop();
        _elapsed.Stop();
        base.OnClosed(e);
    }
}
