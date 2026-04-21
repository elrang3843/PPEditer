using System.Windows;
using System.Windows.Input;
using PPEditer.Models;
using PPEditer.Rendering;

namespace PPEditer.Dialogs;

public partial class SlideShowWindow : Window
{
    private readonly PresentationModel _model;
    private int _index;

    public SlideShowWindow(PresentationModel model, int startIndex)
    {
        InitializeComponent();
        _model = model;
        _index = Math.Max(0, Math.Min(startIndex, model.SlideCount - 1));
        ShowSlide();
    }

    private void ShowSlide()
    {
        var part = _model.GetSlidePart(_index);
        if (part is null) return;
        SlideViewbox.Child = SlideRenderer.BuildCanvas(part, _model.SlideWidth, _model.SlideHeight);
    }

    private void GoNext()
    {
        if (_index < _model.SlideCount - 1) { _index++; ShowSlide(); }
        else Close();
    }

    private void GoPrev()
    {
        if (_index > 0) { _index--; ShowSlide(); }
    }

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.PageDown:
            case Key.Space:
            case Key.Right:
            case Key.Down:
                GoNext(); break;
            case Key.PageUp:
            case Key.Back:
            case Key.Left:
            case Key.Up:
                GoPrev(); break;
            case Key.Escape:
                Close(); break;
        }
    }

    private void OnMouseLeft(object sender, MouseButtonEventArgs e) => GoNext();
    private void OnMouseRight(object sender, MouseButtonEventArgs e) => GoPrev();
}
