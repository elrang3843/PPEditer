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
    public event Action<int>? SlideCopyRequested;
    public event Action<int>? SlideCutRequested;
    public event Action<int>? SlidePasteRequested;
    public event Action<int>? SlideDeleteRequested;

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

    private void ThumbnailList_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        var container = ItemsControl.ContainerFromElement(ThumbnailList, e.OriginalSource as DependencyObject) as ListViewItem;
        if (container is not null)
            container.IsSelected = true;
    }

    private void ThumbnailList_KeyDown(object sender, KeyEventArgs e)
    {
        int idx = ThumbnailList.SelectedIndex;
        if (idx < 0 || _model?.IsOpen != true) return;

        if (e.Key == Key.C && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
        {
            SlideCopyRequested?.Invoke(idx);
            e.Handled = true;
        }
        else if (e.Key == Key.X && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
        {
            SlideCutRequested?.Invoke(idx);
            e.Handled = true;
        }
        else if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
        {
            SlidePasteRequested?.Invoke(idx);
            e.Handled = true;
        }
        else if (e.Key == Key.Delete)
        {
            SlideDeleteRequested?.Invoke(idx);
            e.Handled = true;
        }
    }

    private void OnContextMenuOpened(object sender, RoutedEventArgs e)
    {
        bool hasItem = ThumbnailList.SelectedIndex >= 0 && _model?.IsOpen == true;
        CtxMenuCut.IsEnabled    = hasItem;
        CtxMenuCopy.IsEnabled   = hasItem;
        CtxMenuPaste.IsEnabled  = PresentationModel.HasSlideClipboard && _model?.IsOpen == true;
        CtxMenuDelete.IsEnabled = hasItem && _model?.SlideCount > 1;
    }

    private void OnCtxCopy(object sender, RoutedEventArgs e)
    {
        if (ThumbnailList.SelectedIndex >= 0)
            SlideCopyRequested?.Invoke(ThumbnailList.SelectedIndex);
    }

    private void OnCtxCut(object sender, RoutedEventArgs e)
    {
        if (ThumbnailList.SelectedIndex >= 0)
            SlideCutRequested?.Invoke(ThumbnailList.SelectedIndex);
    }

    private void OnCtxPaste(object sender, RoutedEventArgs e)
    {
        if (ThumbnailList.SelectedIndex >= 0)
            SlidePasteRequested?.Invoke(ThumbnailList.SelectedIndex);
    }

    private void OnCtxDelete(object sender, RoutedEventArgs e)
    {
        if (ThumbnailList.SelectedIndex >= 0)
            SlideDeleteRequested?.Invoke(ThumbnailList.SelectedIndex);
    }
}
