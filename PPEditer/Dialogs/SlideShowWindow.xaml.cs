using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using PPEditer.Models;
using PPEditer.Rendering;

namespace PPEditer.Dialogs;

public partial class SlideShowWindow : Window
{
    private readonly PresentationModel _model;
    private int    _index;
    private bool   _transitioning;

    private Viewbox? _activeViewbox;
    private Canvas?  _activeCanvas;
    private readonly Queue<ShapeAnimation> _pendingAnims = new();

    /// <summary>Fired whenever the displayed slide changes (index is 0-based).</summary>
    public event Action<int>? SlideChanged;

    public SlideShowWindow(PresentationModel model, int startIndex)
    {
        InitializeComponent();
        _model = model;
        _index = Math.Clamp(startIndex, 0, model.SlideCount - 1);
        Loaded += (_, _) => { Focus(); ShowSlide(new SlideTransition { Kind = TransitionKind.None }); };
    }

    // ── Navigation ────────────────────────────────────────────────────

    /// <summary>Navigate relative to the current slide (delta: +1 = next, -1 = prev).</summary>
    public void Navigate(int delta) { if (delta > 0) GoNext(); else GoPrev(); }

    private void GoNext()
    {
        if (_transitioning) return;
        if (_pendingAnims.Count > 0) { PlayNextAnimation(); return; }
        if (_index < _model.SlideCount - 1)
        {
            _index++;
            SlideChanged?.Invoke(_index);
            ShowSlide(_model.GetSlideTransition(_index));
        }
        else Close();
    }

    private void GoPrev()
    {
        if (_transitioning) return;
        _pendingAnims.Clear();
        if (_index > 0)
        {
            _index--;
            SlideChanged?.Invoke(_index);
            ShowSlide(new SlideTransition { Kind = TransitionKind.None });
        }
    }

    // ── Slide display ─────────────────────────────────────────────────

    private void ShowSlide(SlideTransition transition)
    {
        var part = _model.GetSlidePart(_index);
        if (part is null) return;

        var newCanvas = SlideRenderer.BuildCanvas(part, _model.SlideWidth, _model.SlideHeight);
        var newVb     = new Viewbox { Stretch = Stretch.Uniform, Child = newCanvas };

        _pendingAnims.Clear();
        foreach (var anim in _model.GetSlideAnimations(_index).OrderBy(a => a.TreeIndex))
        {
            HideShape(newCanvas, anim.TreeIndex);
            _pendingAnims.Enqueue(anim);
        }

        double dur = transition.DurationMs;

        if (_activeViewbox is null || transition.Kind == TransitionKind.None)
        {
            SlideContainer.Children.Clear();
            SlideContainer.Children.Add(newVb);
            AddWatermarkOverlay();
            _activeViewbox = newVb;
            _activeCanvas  = newCanvas;
            FlushAutoPlay();
            return;
        }

        SlideContainer.Children.Add(newVb);
        _transitioning = true;

        var old = _activeViewbox!;
        void Finish()
        {
            SlideContainer.Children.Remove(old);
            AddWatermarkOverlay();
            _activeViewbox = newVb;
            _activeCanvas  = newCanvas;
            _transitioning = false;
            FlushAutoPlay();
        }

        switch (transition.Kind)
        {
            case TransitionKind.Fade:    PlayFade   (old, newVb, Finish, dur); break;
            case TransitionKind.Push:    PlayPush   (old, newVb, Finish, dur); break;
            case TransitionKind.Wipe:    PlayWipe   (old, newVb, Finish, dur); break;
            case TransitionKind.Flip:    PlayFlip   (old, newVb, Finish, dur); break;
            case TransitionKind.Crumple: PlayCrumple(old, newVb, Finish, dur); break;
            case TransitionKind.Morph:   PlayMorph  (old, newVb, Finish, dur); break;
            default:
                SlideContainer.Children.Remove(old);
                _activeViewbox = newVb;
                _activeCanvas  = newCanvas;
                _transitioning = false;
                FlushAutoPlay();
                break;
        }
    }

    // ── Slide transitions (BeginAnimation) ───────────────────────────

    private void PlayFade(Viewbox old, Viewbox nw, Action done, double dur)
    {
        nw.Opacity = 0;
        old.BeginAnimation(UIElement.OpacityProperty, A(1, 0, dur));
        var fin = A(0, 1, dur); fin.Completed += (_, _) => done();
        nw.BeginAnimation(UIElement.OpacityProperty, fin);
    }

    private void PlayPush(Viewbox old, Viewbox nw, Action done, double dur)
    {
        double w  = ScreenW();
        var oldTt = TT(old, 0, 0);
        var newTt = TT(nw, w, 0);
        oldTt.BeginAnimation(TranslateTransform.XProperty, A(0, -w, dur));
        var fin = A(w, 0, dur); fin.Completed += (_, _) => done();
        newTt.BeginAnimation(TranslateTransform.XProperty, fin);
    }

    private void PlayWipe(Viewbox old, Viewbox nw, Action done, double dur)
    {
        double w  = ScreenW();
        var newTt = TT(nw, -w, 0);
        old.BeginAnimation(UIElement.OpacityProperty, A(1, 0, dur * 0.6));
        var fin = AE(-w, 0, dur, Out); fin.Completed += (_, _) => done();
        newTt.BeginAnimation(TranslateTransform.XProperty, fin);
    }

    private void PlayFlip(Viewbox old, Viewbox nw, Action done, double dur)
    {
        SlideContainer.Children.Remove(nw);
        SlideContainer.Children.Insert(0, nw);
        double h  = ScreenH();
        var oldTt = TT(old, 0, 0);
        old.BeginAnimation(UIElement.OpacityProperty, A(1, 0, dur * 0.6));
        var fin = AE(0, -h, dur, In); fin.Completed += (_, _) => done();
        oldTt.BeginAnimation(TranslateTransform.YProperty, fin);
    }

    private void PlayCrumple(Viewbox old, Viewbox nw, Action done, double dur)
    {
        SlideContainer.Children.Remove(nw);
        SlideContainer.Children.Insert(0, nw);
        nw.Opacity = 0;

        var sc = new ScaleTransform(1, 1);
        var rt = new RotateTransform(0);
        var tr = new TranslateTransform(0, 0);
        var tg = new TransformGroup();
        tg.Children.Add(sc); tg.Children.Add(rt); tg.Children.Add(tr);
        old.RenderTransformOrigin = new Point(0.5, 0.5);
        old.RenderTransform = tg;

        double h = ScreenH();
        sc.BeginAnimation(ScaleTransform.ScaleXProperty,   AE(1, 0.05, dur, In));
        sc.BeginAnimation(ScaleTransform.ScaleYProperty,   AE(1, 0.05, dur, In));
        rt.BeginAnimation(RotateTransform.AngleProperty,   A(0, 30, dur * 0.83));
        tr.BeginAnimation(TranslateTransform.YProperty,    AE(0, -h * 0.3, dur, In));
        old.BeginAnimation(UIElement.OpacityProperty,      A(1, 0, dur * 0.83));
        var fin = A(0, 1, dur); fin.Completed += (_, _) => done();
        nw.BeginAnimation(UIElement.OpacityProperty, fin);
    }

    private void PlayMorph(Viewbox old, Viewbox nw, Action done, double dur)
    {
        nw.Opacity = 0;
        var oldSc = new ScaleTransform(1, 1);
        old.RenderTransformOrigin = new Point(0.5, 0.5);
        old.RenderTransform = oldSc;
        var newSc = new ScaleTransform(0.92, 0.92);
        nw.RenderTransformOrigin = new Point(0.5, 0.5);
        nw.RenderTransform = newSc;

        old.BeginAnimation(UIElement.OpacityProperty, A(1, 0, dur));
        oldSc.BeginAnimation(ScaleTransform.ScaleXProperty, A(1, 1.06, dur));
        oldSc.BeginAnimation(ScaleTransform.ScaleYProperty, A(1, 1.06, dur));
        newSc.BeginAnimation(ScaleTransform.ScaleXProperty, AE(0.92, 1, dur, Out));
        newSc.BeginAnimation(ScaleTransform.ScaleYProperty, AE(0.92, 1, dur, Out));
        var fin = A(0, 1, dur); fin.Completed += (_, _) => done();
        nw.BeginAnimation(UIElement.OpacityProperty, fin);
    }

    // ── Shape animation playback ──────────────────────────────────────

    private void FlushAutoPlay()
    {
        if (_activeCanvas is null) return;
        var remaining = new Queue<ShapeAnimation>();
        while (_pendingAnims.Count > 0)
        {
            var a = _pendingAnims.Dequeue();
            if (a.AutoPlay)
            {
                var el = FindShape(_activeCanvas, a.TreeIndex);
                if (el is not null) { el.Visibility = Visibility.Visible; PlayShapeAnim(el, a); }
            }
            else
            {
                remaining.Enqueue(a);
            }
        }
        while (remaining.Count > 0)
            _pendingAnims.Enqueue(remaining.Dequeue());
    }

    private void PlayNextAnimation()
    {
        if (_pendingAnims.Count == 0 || _activeCanvas is null) return;
        var anim = _pendingAnims.Dequeue();
        var el   = FindShape(_activeCanvas, anim.TreeIndex);
        if (el is null) return;
        el.Visibility = Visibility.Visible;
        PlayShapeAnim(el, anim);
    }

    private static void PlayShapeAnim(FrameworkElement el, ShapeAnimation anim)
    {
        el.RenderTransformOrigin = new Point(0.5, 0.5);
        double dur = anim.DurationMs;
        var    rep = anim.RepeatCount == 0
                     ? RepeatBehavior.Forever
                     : new RepeatBehavior(anim.RepeatCount);

        switch (anim.Kind)
        {
            case AnimationKind.FadeIn:
            {
                el.Opacity = 0;
                var a = A(0, 1, dur); a.RepeatBehavior = rep;
                el.BeginAnimation(UIElement.OpacityProperty, a);
                break;
            }

            case AnimationKind.FlyIn:
            {
                el.Opacity = 0;
                var tt = new TranslateTransform(0, 60);
                el.RenderTransform = tt;
                var aOp = A(0, 1, dur * 0.6); aOp.RepeatBehavior = rep;
                var aTr = AE(60, 0, dur, Out); aTr.RepeatBehavior = rep;
                el.BeginAnimation(UIElement.OpacityProperty, aOp);
                tt.BeginAnimation(TranslateTransform.YProperty, aTr);
                break;
            }

            case AnimationKind.Pulse:
            {
                var sc = new ScaleTransform(1, 1);
                el.RenderTransform = sc;
                var kfx = PulseKF(dur); kfx.RepeatBehavior = rep;
                var kfy = PulseKF(dur); kfy.RepeatBehavior = rep;
                sc.BeginAnimation(ScaleTransform.ScaleXProperty, kfx);
                sc.BeginAnimation(ScaleTransform.ScaleYProperty, kfy);
                break;
            }

            case AnimationKind.Bounce:
            {
                var tt = new TranslateTransform(0, 0);
                el.RenderTransform = tt;
                var kf = new DoubleAnimationUsingKeyFrames { Duration = D(dur), RepeatBehavior = rep };
                kf.KeyFrames.Add(new EasingDoubleKeyFrame(0,   KeyTime.FromPercent(0.00)));
                kf.KeyFrames.Add(new EasingDoubleKeyFrame(-45, KeyTime.FromPercent(0.30),
                    new CubicEase { EasingMode = EasingMode.EaseOut }));
                kf.KeyFrames.Add(new EasingDoubleKeyFrame(0,   KeyTime.FromPercent(1.00),
                    new BounceEase { Bounces = 3, Bounciness = 5 }));
                tt.BeginAnimation(TranslateTransform.YProperty, kf);
                break;
            }

            case AnimationKind.WipeIn:
            {
                double w = double.IsNaN(el.Width) || el.Width <= 0 ? 200 : el.Width;
                el.Opacity = 0;
                var tt = new TranslateTransform(-w - 20, 0);
                el.RenderTransform = tt;
                var aOp = A(0, 1, dur * 0.4); aOp.RepeatBehavior = rep;
                var aTr = AE(-w - 20, 0, dur, Out); aTr.RepeatBehavior = rep;
                el.BeginAnimation(UIElement.OpacityProperty, aOp);
                tt.BeginAnimation(TranslateTransform.XProperty, aTr);
                break;
            }
        }
    }

    private static DoubleAnimationUsingKeyFrames PulseKF(double dur)
    {
        var kf = new DoubleAnimationUsingKeyFrames { Duration = D(dur) };
        kf.KeyFrames.Add(new SplineDoubleKeyFrame(1,    KeyTime.FromPercent(0)));
        kf.KeyFrames.Add(new SplineDoubleKeyFrame(1.18, KeyTime.FromPercent(0.4)));
        kf.KeyFrames.Add(new SplineDoubleKeyFrame(1,    KeyTime.FromPercent(1)));
        return kf;
    }

    // ── Helpers ───────────────────────────────────────────────────────

    private void AddWatermarkOverlay()
    {
        // Remove any previous watermark overlay (last child if non-Viewbox).
        if (SlideContainer.Children.Count > 0
            && SlideContainer.Children[^1] is not Viewbox)
            SlideContainer.Children.RemoveAt(SlideContainer.Children.Count - 1);

        var props = _model.GetDocProperties();
        if (!props.WatermarkShowOnSlide) return;
        var el = WatermarkRenderer.BuildOverlay(props, _model.SlideWidth, _model.SlideHeight);
        if (el is not null) SlideContainer.Children.Add(el);
    }

    private static FrameworkElement? FindShape(Canvas c, int idx) =>
        c.Children.OfType<FrameworkElement>()
                  .FirstOrDefault(fe => fe.Tag is int t && t == idx);

    private static void HideShape(Canvas c, int idx)
    {
        var el = FindShape(c, idx);
        if (el is not null) el.Visibility = Visibility.Hidden;
    }

    private static TranslateTransform TT(UIElement el, double x, double y)
    {
        var tt = new TranslateTransform(x, y);
        ((FrameworkElement)el).RenderTransform = tt;
        return tt;
    }

    private double ScreenW() => SlideContainer.ActualWidth  > 0 ? SlideContainer.ActualWidth  : 1280;
    private double ScreenH() => SlideContainer.ActualHeight > 0 ? SlideContainer.ActualHeight : 720;

    private static DoubleAnimation A(double from, double to, double ms) =>
        new DoubleAnimation(from, to, D(ms));

    private static DoubleAnimation AE(double from, double to, double ms, IEasingFunction ease) =>
        new DoubleAnimation(from, to, D(ms)) { EasingFunction = ease };

    private static Duration D(double ms) => new Duration(TimeSpan.FromMilliseconds(ms));

    private static readonly IEasingFunction Out = new CubicEase { EasingMode = EasingMode.EaseOut };
    private static readonly IEasingFunction In  = new CubicEase { EasingMode = EasingMode.EaseIn  };

    // ── Input ─────────────────────────────────────────────────────────

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.PageDown: case Key.Space: case Key.Right: case Key.Down: GoNext(); break;
            case Key.PageUp:   case Key.Back:  case Key.Left:  case Key.Up:   GoPrev(); break;
            case Key.Escape:                                                   Close();  break;
        }
    }

    private void OnMouseLeft(object sender, MouseButtonEventArgs e)  => GoNext();
    private void OnMouseRight(object sender, MouseButtonEventArgs e) => GoPrev();
}
