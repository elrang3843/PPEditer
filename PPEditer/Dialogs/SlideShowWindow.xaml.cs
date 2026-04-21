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
    private int  _index;
    private bool _transitioning;

    private Viewbox?  _activeViewbox;
    private Canvas?   _activeCanvas;
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
        if (_pendingAnims.Count > 0)
        {
            PlayNextAnimation();
            return;
        }
        if (_index < _model.SlideCount - 1)
        {
            _index++;
            var kind = _model.GetSlideTransition(_index).Kind;
            ShowSlide(kind);
        }
        else
        {
            Close();
        }
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

        // Queue shape animations — hide animated shapes initially
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

        void Finish()
        {
            SlideContainer.Children.Remove(_activeViewbox!);
            _activeViewbox = newVb;
            _activeCanvas  = newCanvas;
            _transitioning = false;
        }

        switch (transition)
        {
            case TransitionKind.Fade:   PlayFade(_activeViewbox, newVb, Finish);   break;
            case TransitionKind.Push:   PlayPush(_activeViewbox, newVb, Finish);   break;
            case TransitionKind.Wipe:   PlayWipe(_activeViewbox, newVb, Finish);   break;
            case TransitionKind.Flip:   PlayFlip(_activeViewbox, newVb, Finish);   break;
            case TransitionKind.Crumple: PlayCrumple(_activeViewbox, newVb, Finish); break;
            case TransitionKind.Morph:  PlayMorph(_activeViewbox, newVb, Finish);  break;
            default:
                SlideContainer.Children.Remove(_activeViewbox);
                _activeViewbox = newVb;
                _activeCanvas  = newCanvas;
                _transitioning = false;
                break;
        }
    }

    // ── Transition animations ─────────────────────────────────────────

    private void PlayFade(Viewbox oldVb, Viewbox newVb, Action done)
    {
        newVb.Opacity = 0;
        var sb = new Storyboard();
        AddAnim(sb, oldVb, UIElement.OpacityProperty, 1, 0, 700);
        AddAnim(sb, newVb, UIElement.OpacityProperty, 0, 1, 700);
        sb.Completed += (_, _) => done();
        sb.Begin();
    }

    private void PlayPush(Viewbox oldVb, Viewbox newVb, Action done)
    {
        double w = SlideContainer.ActualWidth > 0 ? SlideContainer.ActualWidth : 1280;
        var oldTt = ApplyTranslate(oldVb, 0, 0);
        var newTt = ApplyTranslate(newVb, w, 0);
        var sb    = new Storyboard();
        AddAnimTt(sb, oldTt, TranslateTransform.XProperty, 0,  -w, 600);
        AddAnimTt(sb, newTt, TranslateTransform.XProperty, w,   0, 600);
        sb.Completed += (_, _) => done();
        sb.Begin();
    }

    private void PlayWipe(Viewbox oldVb, Viewbox newVb, Action done)
    {
        double w = SlideContainer.ActualWidth > 0 ? SlideContainer.ActualWidth : 1280;
        var newTt = ApplyTranslate(newVb, -w, 0);
        var sb    = new Storyboard();
        AddAnimTt(sb, newTt, TranslateTransform.XProperty, -w, 0, 650,
                  new CubicEase { EasingMode = EasingMode.EaseOut });
        AddAnim(sb, oldVb, UIElement.OpacityProperty, 1, 0, 400);
        sb.Completed += (_, _) => done();
        sb.Begin();
    }

    private void PlayFlip(Viewbox oldVb, Viewbox newVb, Action done)
    {
        // New slide is already visible underneath — insert at bottom of z-order
        SlideContainer.Children.Remove(newVb);
        SlideContainer.Children.Insert(0, newVb);
        double h = SlideContainer.ActualHeight > 0 ? SlideContainer.ActualHeight : 720;
        var oldTt = ApplyTranslate(oldVb, 0, 0);
        var sb    = new Storyboard();
        AddAnimTt(sb, oldTt, TranslateTransform.YProperty, 0, -h, 500,
                  new CubicEase { EasingMode = EasingMode.EaseIn });
        AddAnim(sb, oldVb, UIElement.OpacityProperty, 1, 0, 300);
        sb.Completed += (_, _) => done();
        sb.Begin();
    }

    private void PlayCrumple(Viewbox oldVb, Viewbox newVb, Action done)
    {
        // New slide fades in from behind while old slide shrinks/rotates away
        SlideContainer.Children.Remove(newVb);
        SlideContainer.Children.Insert(0, newVb);
        newVb.Opacity = 0;

        var tg      = new TransformGroup();
        var scale   = new ScaleTransform(1, 1);
        var rotate  = new RotateTransform(0);
        var trans   = new TranslateTransform(0, 0);
        tg.Children.Add(scale); tg.Children.Add(rotate); tg.Children.Add(trans);
        oldVb.RenderTransformOrigin = new Point(0.5, 0.5);
        oldVb.RenderTransform = tg;

        double h = SlideContainer.ActualHeight > 0 ? SlideContainer.ActualHeight : 720;
        var sb = new Storyboard();
        AddAnimTarget(sb, scale,  ScaleTransform.ScaleXProperty,   1, 0.05, 600);
        AddAnimTarget(sb, scale,  ScaleTransform.ScaleYProperty,   1, 0.05, 600);
        AddAnimTarget(sb, rotate, RotateTransform.AngleProperty,   0, 30,   500);
        AddAnimTarget(sb, trans,  TranslateTransform.YProperty,    0, -h * 0.3, 600,
                      new CubicEase { EasingMode = EasingMode.EaseIn });
        AddAnim(sb, oldVb,  UIElement.OpacityProperty, 1, 0, 500);
        AddAnim(sb, newVb,  UIElement.OpacityProperty, 0, 1, 600);
        sb.Completed += (_, _) => done();
        sb.Begin();
    }

    private void PlayMorph(Viewbox oldVb, Viewbox newVb, Action done)
    {
        // Smooth zoom cross-fade (simplified morph approximation)
        newVb.Opacity = 0;
        newVb.RenderTransformOrigin = new Point(0.5, 0.5);
        newVb.RenderTransform = new ScaleTransform(0.92, 0.92);
        oldVb.RenderTransformOrigin = new Point(0.5, 0.5);
        oldVb.RenderTransform = new ScaleTransform(1, 1);

        var sb = new Storyboard();
        AddAnim(sb, oldVb, UIElement.OpacityProperty, 1, 0, 600);
        AddAnimTarget(sb, (ScaleTransform)oldVb.RenderTransform, ScaleTransform.ScaleXProperty, 1, 1.06, 600);
        AddAnimTarget(sb, (ScaleTransform)oldVb.RenderTransform, ScaleTransform.ScaleYProperty, 1, 1.06, 600);
        AddAnim(sb, newVb, UIElement.OpacityProperty, 0, 1, 600);
        AddAnimTarget(sb, (ScaleTransform)newVb.RenderTransform, ScaleTransform.ScaleXProperty, 0.92, 1, 600,
                      new CubicEase { EasingMode = EasingMode.EaseOut });
        AddAnimTarget(sb, (ScaleTransform)newVb.RenderTransform, ScaleTransform.ScaleYProperty, 0.92, 1, 600,
                      new CubicEase { EasingMode = EasingMode.EaseOut });
        sb.Completed += (_, _) => done();
        sb.Begin();
    }

    // ── Storyboard helpers ────────────────────────────────────────────

    private static void AddAnim(Storyboard sb, DependencyObject target,
                                 DependencyProperty prop,
                                 double from, double to, double ms,
                                 IEasingFunction? ease = null)
    {
        var a = new DoubleAnimation(from, to, D(ms)) { EasingFunction = ease };
        Storyboard.SetTarget(a, target);
        Storyboard.SetTargetProperty(a, new PropertyPath(prop));
        sb.Children.Add(a);
    }

    private static void AddAnimTt(Storyboard sb, TranslateTransform tt,
                                   DependencyProperty prop,
                                   double from, double to, double ms,
                                   IEasingFunction? ease = null)
    {
        var a = new DoubleAnimation(from, to, D(ms)) { EasingFunction = ease };
        Storyboard.SetTarget(a, tt);
        Storyboard.SetTargetProperty(a, new PropertyPath(prop));
        sb.Children.Add(a);
    }

    private static void AddAnimTarget(Storyboard sb, Animatable target,
                                       DependencyProperty prop,
                                       double from, double to, double ms,
                                       IEasingFunction? ease = null)
    {
        var a = new DoubleAnimation(from, to, D(ms)) { EasingFunction = ease };
        Storyboard.SetTarget(a, target);
        Storyboard.SetTargetProperty(a, new PropertyPath(prop));
        sb.Children.Add(a);
    }

    private static TranslateTransform ApplyTranslate(UIElement el, double x, double y)
    {
        var tt = new TranslateTransform(x, y);
        ((FrameworkElement)el).RenderTransform = tt;
        return tt;
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
        double dur = anim.DurationMs;
        el.RenderTransformOrigin = new Point(0.5, 0.5);

        switch (anim.Kind)
        {
            case AnimationKind.FadeIn:
                el.Opacity = 0;
                el.BeginAnimation(UIElement.OpacityProperty,
                    new DoubleAnimation(0, 1, D(dur)));
                break;

            case AnimationKind.FlyIn:
            {
                el.Opacity = 0;
                var tt = new TranslateTransform(0, 60);
                el.RenderTransform = tt;
                el.BeginAnimation(UIElement.OpacityProperty,
                    new DoubleAnimation(0, 1, D(dur * 0.6)));
                tt.BeginAnimation(TranslateTransform.YProperty,
                    new DoubleAnimation(60, 0, D(dur))
                    { EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut } });
                break;
            }

            case AnimationKind.Pulse:
            {
                var st = new ScaleTransform(1, 1);
                el.RenderTransform = st;
                var kf = new DoubleAnimationUsingKeyFrames();
                kf.KeyFrames.Add(new SplineDoubleKeyFrame(1,   KeyTime.FromPercent(0)));
                kf.KeyFrames.Add(new SplineDoubleKeyFrame(1.18, KeyTime.FromPercent(0.4)));
                kf.KeyFrames.Add(new SplineDoubleKeyFrame(1,   KeyTime.FromPercent(1)));
                kf.Duration = D(dur);
                st.BeginAnimation(ScaleTransform.ScaleXProperty, kf);
                st.BeginAnimation(ScaleTransform.ScaleYProperty, (AnimationTimeline)kf.Clone());
                break;
            }

            case AnimationKind.Bounce:
            {
                var tt = new TranslateTransform(0, 0);
                el.RenderTransform = tt;
                var kf = new DoubleAnimationUsingKeyFrames();
                kf.KeyFrames.Add(new EasingDoubleKeyFrame(0,   KeyTime.FromPercent(0)));
                kf.KeyFrames.Add(new EasingDoubleKeyFrame(-45, KeyTime.FromPercent(0.3),
                    new CubicEase { EasingMode = EasingMode.EaseOut }));
                kf.KeyFrames.Add(new EasingDoubleKeyFrame(0,   KeyTime.FromPercent(1.0),
                    new BounceEase { Bounces = 3, Bounciness = 5 }));
                kf.Duration = D(dur);
                tt.BeginAnimation(TranslateTransform.YProperty, kf);
                break;
            }

            case AnimationKind.WipeIn:
            {
                el.Opacity = 0;
                var tt = new TranslateTransform(-el.Width - 20, 0);
                el.RenderTransform = tt;
                el.BeginAnimation(UIElement.OpacityProperty,
                    new DoubleAnimation(0, 1, D(dur * 0.3)));
                tt.BeginAnimation(TranslateTransform.XProperty,
                    new DoubleAnimation(-el.Width - 20, 0, D(dur))
                    { EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut } });
                break;
            }
        }
    }

    private static FrameworkElement? FindShape(Canvas canvas, int treeIndex) =>
        canvas.Children.OfType<FrameworkElement>()
                        .FirstOrDefault(fe => fe.Tag is int t && t == treeIndex);

    private static void HideShape(Canvas canvas, int treeIndex)
    {
        var el = FindShape(canvas, treeIndex);
        if (el is not null) el.Visibility = Visibility.Hidden;
    }

    // ── Input handlers ────────────────────────────────────────────────

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

    private static Duration D(double ms) => new Duration(TimeSpan.FromMilliseconds(ms));
}
