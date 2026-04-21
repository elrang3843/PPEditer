namespace PPEditer.Models;

public enum DrawTool
{
    Select,           // 선택 (기본)

    // 선
    Line,             // 직선 (드래그)
    SplineLine,       // 스프라인 선 (클릭으로 점 추가, 더블클릭 완료)
    PolyLine,         // 폴리곤 선 (클릭으로 점 추가, 더블클릭 완료)

    // 사각형
    Square,           // 정사각형 (드래그, 1:1 비율 고정)
    Rect,             // 직사각형
    Trapezoid,        // 사다리꼴
    Parallelogram,    // 평행사변형

    // 삼각형
    EqTriangle,       // 정삼각형 (드래그, 1:1 비율 고정)
    IsoTriangle,      // 이등변 삼각형
    RightTriangle,    // 직각 삼각형
    ScaleneTriangle,  // 비대칭 삼각형 (클릭 3점)

    // 원/타원
    Circle,           // 원 (드래그, 1:1 비율 고정)
    Ellipse,          // 타원
    Arc,              // 호

    // 면
    Polygon,          // 폴리곤 면 (클릭으로 점 추가, 더블클릭 완료)
    SplinePolygon,    // 스프라인 면 (클릭으로 점 추가, 더블클릭 완료)

    // 화살표
    Arrow,            // 오른쪽 화살표
}
