using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using PPEditer.Models;
using PPEditer.Rendering;

namespace PPEditer.Controls;

public sealed class SlideThumbnailItem
{
    public int          Index     { get; set; }
    public int          Number    => Index + 1;
    public BitmapSource? Thumbnail { get; set; }
}

public partial class SlideThumbnailPanel : UserControl
{
    public event Action<int>? SlideSelected;
    public event Action<int>? SlideDoubleClicked;

    private readonly ObservableCollection<SlideThumbnailItem> _items = new();
    private PresentationModel? _model;

    public SlideThumbnailPanel()
    {
        InitializeComponent();
        ThumbnailList.ItemsSource = _items;
    }

    // ── Public API ─────────────────────────────────────────────────────

    public void Refresh(PresentationModel model, int currentIndex)
    {
        _model = model;
        _items.Clear();
        if (!model.IsOpen) return;

        for (int i = 0; i < model.SlideCount; i++)
            _items.Add(new SlideThumbnailItem { Index = i, Thumbnail = RenderThumb(i) });

        SetCurrent(currentIndex);
    }

    public void RefreshSingle(int index)
    {
        if (_model is null || index < 0 || index >= _items.Count) return;
        _items[index].Thumbnail = RenderThumb(index);
        // Force ListView to re-evaluate the binding
        var item = _items[index];
        _items[index] = new SlideThumbnailItem { Index = index, Thumbnail = item.Thumbnail };
    }

    public void SetCurrent(int index)
    {
        if (index >= 0 && index < _items.Count)
        {
            ThumbnailList.SelectedIndex = index;
            ThumbnailList.ScrollIntoView(ThumbnailList.Items[index]);
        }
    }

    // ── Internal ───────────────────────────────────────────────────────

    private BitmapSource? RenderThumb(int index)
    {
        if (_model is null) return null;
        try
        {
            var part = _model.GetSlidePart(index);
            if (part is null) return null;
            return SlideRenderer.RenderToBitmap(part, _model.SlideWidth, _model.SlideHeight, 160, 90);
        }
        catch
        {
            return null;
        }
    }

    private void ThumbnailList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ThumbnailList.SelectedIndex >= 0)
            SlideSelected?.Invoke(ThumbnailList.SelectedIndex);
    }

    private void ThumbnailList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (ThumbnailList.SelectedIndex >= 0)
            SlideDoubleClicked?.Invoke(ThumbnailList.SelectedIndex);
    }
}
