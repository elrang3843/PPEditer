namespace PPEditer.Models;

public enum AnimationKind
{
    None,
    FadeIn,   // 밝기 변화/나타나기
    FlyIn,    // 날아오기 (아래에서 위로)
    Pulse,    // 펄스/강조 (확대 후 복귀)
    Bounce,   // 팝콘/튕김
    WipeIn,   // 왼쪽에서 닦아내기
}

public sealed class ShapeAnimation
{
    public int           TreeIndex  { get; set; }
    public AnimationKind Kind       { get; set; }
    public double        DurationMs { get; set; } = 500;
}
