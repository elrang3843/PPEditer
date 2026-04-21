namespace PPEditer.Models;

public enum TransitionKind
{
    None,
    Fade,     // 밝기 변화
    Push,     // 밀어내기
    Wipe,     // 나타나기
    Flip,     // 집어내기
    Crumple,  // 구겨 던지기
    Morph,    // 모핑 (WPF slideshow only; stored as Fade in PPTX)
}

public sealed class SlideTransition
{
    public TransitionKind Kind       { get; set; } = TransitionKind.None;
    public double         DurationMs { get; set; } = 700;
}
