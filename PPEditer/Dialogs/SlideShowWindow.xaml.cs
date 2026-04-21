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

    public SlideShowWindow(PresentationModel model, int startIndex)
    {
        InitializeComponent();
        _model = model;
        _index = Math.Clamp(startIndex, 0, model.SlideCount - 1);
        Loaded += (_, _) => { Focus(); ShowSlide(TransitionKind.None); };
    }

    // ── Navigation ────────────────────────────────────────────────────

    private void GoNext()
    {
        if (_transitioning) return;
        if (_pendingAnims.Count > 0) { PlayNextAnimation(); return; }
        if (_index < _model.SlideCount - 1)
        {
            _index++;
            ShowSlide(_model.GetSlideTransition(_index).Kind);
        }
        else Close();
    }

    private void GoPrev()
    {
        if (_transitioning) return;
        _pendingAnims.Clear();
        if (_index > 0) { _index--; ShowSlide(TransitionKind.None); }
    }

    // ── Slide display ─────────────────────────────────────────────────

    private void ShowSlide(TransitionKind transition)
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

        if (_activeViewbox is null || transition == TransitionKind.None)
        {
            SlideContainer.Children.Clear();
            SlideContainer.Children.Add(newVb);
            _activeViewbox = newVb;
            _activeCanvas  = newCanvas;
            return;
        }

        SlideContainer.Children.Add(newVb);
        _transitioning = true;

        var old = _activeViewbox!;
        void Finish()
        {
            SlideContainer.Children.Remove(old);
            _activeViewbox = newVb;
            _activeCanvas  = newCanvas;
            _transitioning = false;
        }

        switch (transition)
        {
            case TransitionKind.Fade:    PlayFade   (old, newVb, Finish); break;
            case TransitionKind.Push:    PlayPush   (old, newVb, Finish); break;
            case TransitionKind.Wipe:    PlayWipe   (old, newVb, Finish); break;
            case TransitionKind.Flip:    PlayFlip   (old, newVb, Finish); break;
            case TransitionKind.Crumple: PlayCrumple(old, newVb, Finish); break;
            case TransitionKind.Morph:   PlayMorph  (old, newVb, Finish); break;
            default:
                SlideContainer.Children.Remove(old);
                _activeViewbox = newVb;
                _activeCanvas  = newCanvas;
                _transitioning = false;
                break;
        }
    }

    // ── Slide transitions (BeginAnimation) ───────────────────────────

    private void PlayFade(Viewbox old, Viewbox nw, Action done)
    {
        nw.Opacity = 0;
        old.BeginAnimation(UIElement.OpacityProperty, A(1, 0, 700));
        var fin = A(0, 1, 700); fin.Completed += (_, _) => done();
        nw.BeginAnimation(UIElement.OpacityProperty, fin);
    }

    private void PlayPush(Viewbox old, Viewbox nw, Action done)
    {
        double w  = ScreenW();
        var oldTt = TT(old, 0, 0);
        var newTt = TT(nw, w, 0);
        oldTt.BeginAnimation(TranslateTransform.XProperty, A(0, -w, 600));
        var fin = A(w, 0, 600); fin.Completed += (_, _) => done();
        newTt.BeginAnimation(TranslateTransform.XProperty, fin);
    }

    private void PlayWipe(Viewbox old, Viewbox nw, Action done)
    {
        double w  = ScreenW();
        var newTt = TT(nw, -w, 0);
        old.BeginAnimation(UIElement.OpacityProperty, A(1, 0, 400));
        var fin = AE(-w, 0, 650, Out); fin.Completed += (_, _) => done();
        newTt.BeginAnimation(TranslateTransform.XProperty, fin);
    }

    private void PlayFlip(Viewbox old, Viewbox nw, Action done)
    {
        SlideContainer.Children.Remove(nw);
        SlideContainer.Children.Insert(0, nw);
        double h  = ScreenH();
        var oldTt = TT(old, 0, 0);
        old.BeginAnimation(UIElement.OpacityProperty, A(1, 0, 300));
        var fin = AE(0, -h, 500, In); fin.Completed += (_, _) => done();
        oldTt.BeginAnimation(TranslateTransform.YProperty, fin);
    }

    private void PlayCrumple(Viewbox old, Viewbox nw, Action done)
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
        sc.BeginAnimation(ScaleTransform.ScaleXProperty,   AE(1, 0.05, 600, In));
        sc.BeginAnimation(ScaleTransform.ScaleYProperty,   AE(1, 0.05, 600, In));
        rt.BeginAnimation(RotateTransform.AngleProperty,   A(0, 30, 500));
        tr.BeginAnimation(TranslateTransform.YProperty,    AE(0, -h * 0.3, 600, In));
        old.BeginAnimation(UIElement.OpacityProperty,      A(1, 0, 500));
        var fin = A(0, 1, 600); fin.Completed += (_, _) => done();
        nw.BeginAnimation(UIElement.OpacityProperty, fin);
    }

    private void PlayMorph(Viewbox old, Viewbox nw, Action done)
    {
        nw.Opacity = 0;
        var oldSc = new ScaleTransform(1, 1);
        old.RenderTransformOrigin = new Point(0.5, 0.5);
        old.RenderTransform = oldSc;
        var newSc = new ScaleTransform(0.92, 0.92);
        nw.RenderTransformOrigin = new Point(0.5, 0.5);
        nw.RenderTransform = newSc;

        old.BeginAnimation(UIElement.OpacityProperty, A(1, 0, 600));
        oldSc.BeginAnimation(ScaleTransform.ScaleXProperty, A(1, 1.06, 600));
        oldSc.BeginAnimation(ScaleTransform.ScaleYProperty, A(1, 1.06, 600));
        newSc.BeginAnimation(ScaleTransform.ScaleXProperty, AE(0.92, 1, 600, Out));
        newSc.BeginAnimation(ScaleTransform.ScaleYProperty, AE(0.92, 1, 600, Out));
        var fin = A(0, 1, 600); fin.Completed += (_, _) => done();
        nw.BeginAnimation(UIElement.OpacityProperty, fin);
    }

    // ── Shape animation playback ──────────────────────────────────────

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

        switch (anim.Kind)
        {
            case AnimationKind.FadeIn:
                el.Opacity = 0;
                el.BeginAnimation(UIElement.OpacityProperty, A(0, 1, dur));
                break;

            case AnimationKind.FlyIn:
            {
                el.Opacity = 0;
                var tt = new TranslateTransform(0, 60);
                el.RenderTransform = tt;
                el.BeginAnimation(UIElement.OpacityProperty, A(0, 1, dur * 0.6));
                tt.BeginAnimation(TranslateTransform.YProperty, AE(60, 0, dur, Out));
                break;
            }

            case AnimationKind.Pulse:
            {
                var sc = new ScaleTransform(1, 1);
                el.RenderTransform = sc;
                sc.BeginAnimation(ScaleTransform.ScaleXProperty, PulseKF(dur));
                sc.BeginAnimation(ScaleTransform.ScaleYProperty, PulseKF(dur));
                break;
            }

            case AnimationKind.Bounce:
            {
                var tt = new TranslateTransform(0, 0);
                el.RenderTransform = tt;
                var kf = new DoubleAnimationUsingKeyFrames { Duration = D(dur) };
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
                el.BeginAnimation(UIElement.OpacityProperty, A(0, 1, dur * 0.4));
                tt.BeginAnimation(TranslateTransform.XProperty, AE(-w - 20, 0, dur, Out));
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
