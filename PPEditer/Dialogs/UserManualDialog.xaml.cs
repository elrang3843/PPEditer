using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PPEditer.Dialogs;

public partial class UserManualDialog : Window
{
    private readonly Dictionary<string, StackPanel> _pages = new();

    private static readonly string[] Sections =
    [
        "시작하기", "파일 관리", "슬라이드 편집",
        "도형 및 텍스트", "서식 지정", "효과 넣기",
        "슬라이드 쇼", "단축키"
    ];

    public UserManualDialog()
    {
        InitializeComponent();
        foreach (var s in Sections)
        {
            NavList.Items.Add(new ListBoxItem
            {
                Content = s,
                Padding = new Thickness(14, 7, 14, 7),
            });
            _pages[s] = BuildPage(s);
        }
        NavList.SelectedIndex = 0;
    }

    private void OnNavChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NavList.SelectedItem is ListBoxItem li &&
            li.Content is string key &&
            _pages.TryGetValue(key, out var panel))
        {
            ContentArea.Content = panel;
            ContentArea.ScrollToTop();
        }
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();

    // ── Page router ───────────────────────────────────────────────────

    private static StackPanel BuildPage(string section) => section switch
    {
        "시작하기"      => PageGettingStarted(),
        "파일 관리"     => PageFileManagement(),
        "슬라이드 편집"  => PageSlideEditing(),
        "도형 및 텍스트" => PageShapesText(),
        "서식 지정"     => PageFormatting(),
        "효과 넣기"     => PageEffects(),
        "슬라이드 쇼"   => PageSlideShow(),
        "단축키"        => PageShortcuts(),
        _              => new StackPanel(),
    };

    // ── Pages ─────────────────────────────────────────────────────────

    private static StackPanel PageGettingStarted()
    {
        var p = P();
        Add(p, H1("PPEditer 시작하기"));
        Add(p, Para("PPEditer는 Microsoft PowerPoint PPTX 파일을 열고 편집하는 프레젠테이션 편집 프로그램입니다."));

        Add(p, H2("화면 구성"));
        Add(p, Bullet("상단 메뉴 및 도구 모음 — 주요 기능에 빠르게 접근"));
        Add(p, Bullet("왼쪽 슬라이드 패널 — 전체 슬라이드 목록 및 탐색"));
        Add(p, Bullet("중앙 편집 영역 — 슬라이드를 직접 편집"));
        Add(p, Bullet("하단 상태 표시줄 — 현재 슬라이드 번호, 확대/축소 비율"));

        Add(p, H2("빠른 시작"));
        Add(p, Bullet("새 프레젠테이션: 파일 > 새로 만들기 (Ctrl+N)"));
        Add(p, Bullet("기존 파일 열기: 파일 > 열기 (Ctrl+O)"));
        Add(p, Bullet("최근 파일 다시 열기: 파일 > 최근 파일"));

        Add(p, H2("파일 형식"));
        Add(p, Bullet("지원 형식: Microsoft PowerPoint (.pptx)"));
        Add(p, Bullet("내보내기: PDF 형식 지원"));
        return p;
    }

    private static StackPanel PageFileManagement()
    {
        var p = P();
        Add(p, H1("파일 관리"));

        Add(p, H2("새로 만들기  (Ctrl+N)"));
        Add(p, Para("빈 프레젠테이션을 새로 만듭니다. 저장하지 않은 변경사항이 있으면 저장 여부를 묻습니다."));

        Add(p, H2("열기  (Ctrl+O)"));
        Add(p, Para("기존 .pptx 파일을 엽니다. 파일 선택 대화상자에서 파일을 선택합니다."));

        Add(p, H2("최근 파일"));
        Add(p, Para("파일 > 최근 파일에서 최근에 열었던 파일을 빠르게 다시 열 수 있습니다."));

        Add(p, H2("저장  (Ctrl+S)"));
        Add(p, Para("현재 파일에 덮어쓰기 저장합니다. 새 파일인 경우 저장 위치를 지정합니다."));

        Add(p, H2("다른 이름으로 저장  (Ctrl+Shift+S)"));
        Add(p, Para("새 파일 이름 또는 다른 위치에 저장합니다."));

        Add(p, H2("PDF로 내보내기  (Ctrl+Shift+E)"));
        Add(p, Para("모든 슬라이드를 PDF 파일로 내보냅니다. 각 슬라이드가 한 페이지로 저장됩니다."));

        Add(p, H2("암호 설정"));
        Add(p, Para("파일 > 문서 정보 (Ctrl+Shift+I) > 암호 설정에서 PPTX 파일에 열기 암호를 설정하거나 제거합니다."));
        return p;
    }

    private static StackPanel PageSlideEditing()
    {
        var p = P();
        Add(p, H1("슬라이드 편집"));

        Add(p, H2("슬라이드 탐색"));
        Add(p, Bullet("왼쪽 슬라이드 패널에서 원하는 슬라이드를 클릭하여 선택합니다."));
        Add(p, Bullet("패널 아래 ◀ ▶ 버튼으로 이전/다음 슬라이드로 이동합니다."));

        Add(p, H2("슬라이드 추가  (Ctrl+M)"));
        Add(p, Bullet("메뉴 슬라이드 > 새 슬라이드 또는 슬라이드 패널 아래 [+] 버튼 클릭"));
        Add(p, Bullet("현재 슬라이드 다음에 새 슬라이드가 추가됩니다."));

        Add(p, H2("슬라이드 복제  (Ctrl+D)"));
        Add(p, Bullet("현재 슬라이드와 동일한 내용의 슬라이드를 바로 다음에 삽입합니다."));

        Add(p, H2("슬라이드 삭제  (Ctrl+Shift+Delete)"));
        Add(p, Bullet("현재 선택된 슬라이드를 삭제합니다."));

        Add(p, H2("슬라이드 순서 변경"));
        Add(p, Bullet("슬라이드 패널에서 슬라이드를 드래그하여 원하는 위치로 이동합니다."));
        Add(p, Bullet("메뉴 슬라이드 > 위로 이동 (Ctrl+Shift+↑) / 아래로 이동 (Ctrl+Shift+↓)"));
        return p;
    }

    private static StackPanel PageShapesText()
    {
        var p = P();
        Add(p, H1("도형 및 텍스트"));

        Add(p, H2("도형 선택"));
        Add(p, Bullet("도구 모음에서 화살표(선택) 도구를 선택한 후 도형을 클릭합니다."));
        Add(p, Bullet("빈 영역 드래그로 범위 내 여러 도형을 한꺼번에 선택합니다."));
        Add(p, Bullet("빈 영역 클릭으로 선택을 해제합니다."));

        Add(p, H2("도형 이동 및 크기 조절"));
        Add(p, Bullet("이동: 선택한 도형을 드래그합니다."));
        Add(p, Bullet("미세 이동: 선택 후 방향키 (↑↓←→) 사용."));
        Add(p, Bullet("크기 조절: 선택 후 테두리의 핸들을 드래그합니다."));

        Add(p, H2("텍스트 편집"));
        Add(p, Bullet("텍스트 도형을 더블클릭하면 편집 모드가 활성화됩니다."));
        Add(p, Bullet("편집 완료: Esc 키 또는 도형 바깥 클릭."));
        Add(p, Bullet("줄 바꿈: Enter 키."));
        Add(p, Bullet("굵게 Ctrl+B / 기울임 Ctrl+I / 밑줄 Ctrl+U"));

        Add(p, H2("도형 그리기  (삽입 메뉴 또는 도구 모음)"));
        Add(p, Bullet("텍스트 상자: 삽입 > 텍스트 상자 (Ctrl+Shift+T)"));
        Add(p, Bullet("선 / 사각형 / 삼각형 / 원 / 폴리곤 / 화살표 등: 삽입 > 도형 그리기"));
        Add(p, Bullet("그리기 후 자동으로 선택 도구로 전환됩니다."));

        Add(p, H2("도형 삭제"));
        Add(p, Bullet("도형 선택 후 Delete 키를 누릅니다."));

        Add(p, H2("겹침 순서 변경"));
        Add(p, Bullet("도형 우클릭 > 순서 변경 > 맨 앞 / 앞으로 / 뒤로 / 맨 뒤 선택."));

        Add(p, H2("그룹 및 해제"));
        Add(p, Bullet("여러 도형 선택 후 Ctrl+G로 그룹화합니다."));
        Add(p, Bullet("그룹 선택 후 Ctrl+Shift+G로 그룹을 해제합니다."));

        Add(p, H2("특수 삽입"));
        Add(p, Bullet("수식/과학 기호: 삽입 > 수식/과학 기호 (Ctrl+Shift+M)"));
        Add(p, Bullet("특수 문자: 삽입 > 특수 문자 (Ctrl+Shift+K)"));
        Add(p, Bullet("이모지: 삽입 > 이모지 (Ctrl+Shift+J)"));
        return p;
    }

    private static StackPanel PageFormatting()
    {
        var p = P();
        Add(p, H1("서식 지정"));

        Add(p, H2("글자 속성  (Ctrl+Shift+F)"));
        Add(p, Para("텍스트를 선택하거나 텍스트 편집 모드에서 서식 > 글자 속성을 엽니다."));
        Add(p, Bullet("글꼴 종류, 크기, 굵게 / 기울임 / 밑줄 설정"));
        Add(p, Bullet("글자 색상 변경"));
        Add(p, Bullet("위첨자 / 아래첨자 설정"));

        Add(p, H2("문단 속성  (Ctrl+Shift+P)"));
        Add(p, Para("텍스트 편집 모드에서 서식 > 문단 속성을 엽니다."));
        Add(p, Bullet("가로 정렬: 왼쪽 / 가운데 / 오른쪽 / 양쪽"));
        Add(p, Bullet("줄 간격, 단락 앞/뒤 간격 설정"));
        Add(p, Bullet("들여쓰기 설정"));

        Add(p, H2("도형 속성"));
        Add(p, Para("도형 선택 후 우클릭 > 도형 속성을 엽니다."));
        Add(p, Bullet("채우기 색상 및 투명도 설정"));
        Add(p, Bullet("테두리 색상, 두께, 스타일 설정"));
        Add(p, Bullet("도형 위치 및 크기 직접 입력"));
        Add(p, Bullet("세로 텍스트 정렬 (위 / 가운데 / 아래)"));

        Add(p, H2("문서 정보  (Ctrl+Shift+I)"));
        Add(p, Para("보기 > 문서 정보에서 슬라이드 크기, 방향, 암호를 설정합니다."));
        Add(p, Bullet("표준 16:9, 와이드스크린, A4, 커스텀 크기 등"));
        Add(p, Bullet("열기 암호 설정 및 제거"));
        return p;
    }

    private static StackPanel PageEffects()
    {
        var p = P();
        Add(p, H1("효과 넣기"));

        Add(p, H2("슬라이드 전환 효과"));
        Add(p, Para("메뉴 효과넣기 > 슬라이드 전환 효과에서 설정합니다."));
        Add(p, Bullet("없음: 즉시 전환"));
        Add(p, Bullet("밝기변화 (Fade): 부드럽게 사라지고 나타남"));
        Add(p, Bullet("밀어내기 (Push): 새 슬라이드가 오른쪽에서 밀어냄"));
        Add(p, Bullet("나타나기 (Wipe): 왼쪽에서 닦아내며 전환"));
        Add(p, Bullet("집어내기 (Flip): 현재 슬라이드를 위로 집어내는 효과"));
        Add(p, Bullet("구겨 던지기 (Crumple): 종이를 구겨 던지는 효과"));
        Add(p, Bullet("모핑 (Morph): 개체가 자연스럽게 변형되며 이동"));
        Add(p, Para("▶ 시간(초) 설정으로 전환 속도를 조절합니다 (0.1초 단위)."));
        Add(p, Para("▶ '모든 슬라이드에 적용' 체크 시 전체 슬라이드에 일괄 적용됩니다."));

        Add(p, H2("개체 애니메이션 효과"));
        Add(p, Para("도형 선택 후 메뉴 효과넣기 > 개체 애니메이션 효과에서 설정합니다."));
        Add(p, Bullet("없음: 슬라이드와 함께 즉시 표시"));
        Add(p, Bullet("밝기변화/나타나기 (Fade In): 서서히 나타남"));
        Add(p, Bullet("날아오기 (Fly In): 아래에서 위로 날아오며 등장"));
        Add(p, Bullet("펄스/강조 (Pulse): 커졌다가 원래 크기로 복귀"));
        Add(p, Bullet("팝콘/튕김 (Bounce): 위로 튀어 오르는 효과"));
        Add(p, Bullet("닦아내기 (Wipe In): 왼쪽에서 오른쪽으로 나타남"));
        Add(p, Para("▶ 시간(초): 애니메이션 재생 시간 (0.1초 단위)."));
        Add(p, Para("▶ 반복 횟수: 0 입력 시 무한 반복."));
        Add(p, Para("▶ 자동 실행: 체크 시 슬라이드가 나타날 때 자동 재생, 미체크 시 클릭할 때 재생."));
        return p;
    }

    private static StackPanel PageSlideShow()
    {
        var p = P();
        Add(p, H1("슬라이드 쇼"));

        Add(p, H2("시작  (F5)"));
        Add(p, Bullet("메뉴 슬라이드 > 슬라이드 쇼 또는 F5 키"));
        Add(p, Bullet("현재 슬라이드부터 전체 화면으로 시작됩니다."));

        Add(p, H2("탐색"));
        Add(p, Bullet("다음 / 애니메이션 재생: PageDown, Space, →, ↓, 마우스 왼쪽 클릭"));
        Add(p, Bullet("이전 슬라이드: PageUp, Backspace, ←, ↑, 마우스 오른쪽 클릭"));
        Add(p, Bullet("종료: Esc 키"));

        Add(p, H2("애니메이션 재생 순서"));
        Add(p, Bullet("자동 실행이 설정된 개체는 슬라이드가 나타날 때 바로 재생됩니다."));
        Add(p, Bullet("수동 실행 개체는 다음(클릭/Enter)을 누를 때마다 순서대로 재생됩니다."));
        Add(p, Bullet("모든 애니메이션이 재생된 후 다음을 누르면 다음 슬라이드로 이동합니다."));
        Add(p, Bullet("이전으로 이동 시 진행 중인 애니메이션은 취소됩니다."));

        Add(p, H2("슬라이드 쇼 종료"));
        Add(p, Para("마지막 슬라이드에서 다음으로 이동하거나 Esc를 누르면 종료됩니다."));
        return p;
    }

    private static StackPanel PageShortcuts()
    {
        var p = P();
        Add(p, H1("키보드 단축키"));

        Add(p, H2("파일"));
        Add(p, SC("Ctrl+N",           "새로 만들기"));
        Add(p, SC("Ctrl+O",           "열기"));
        Add(p, SC("Ctrl+S",           "저장"));
        Add(p, SC("Ctrl+Shift+S",     "다른 이름으로 저장"));
        Add(p, SC("Ctrl+Shift+E",     "PDF로 내보내기"));

        Add(p, H2("편집"));
        Add(p, SC("Ctrl+Z",           "실행 취소"));
        Add(p, SC("Ctrl+Y",           "다시 실행"));
        Add(p, SC("Ctrl+C",           "복사"));
        Add(p, SC("Ctrl+X",           "잘라내기"));
        Add(p, SC("Ctrl+V",           "붙여넣기"));

        Add(p, H2("슬라이드"));
        Add(p, SC("Ctrl+M",           "슬라이드 추가"));
        Add(p, SC("Ctrl+D",           "슬라이드 복제"));
        Add(p, SC("Ctrl+Shift+Delete","슬라이드 삭제"));
        Add(p, SC("Ctrl+Shift+↑",     "슬라이드 위로 이동"));
        Add(p, SC("Ctrl+Shift+↓",     "슬라이드 아래로 이동"));
        Add(p, SC("F5",               "슬라이드 쇼 시작"));

        Add(p, H2("도형"));
        Add(p, SC("Delete",           "선택한 도형 삭제"));
        Add(p, SC("↑↓←→",            "선택한 도형 미세 이동"));
        Add(p, SC("Ctrl+G",           "그룹으로 묶기"));
        Add(p, SC("Ctrl+Shift+G",     "그룹 해제"));

        Add(p, H2("삽입"));
        Add(p, SC("Ctrl+Shift+T",     "텍스트 상자"));
        Add(p, SC("Ctrl+Shift+M",     "수식/과학 기호"));
        Add(p, SC("Ctrl+Shift+K",     "특수 문자"));
        Add(p, SC("Ctrl+Shift+J",     "이모지"));

        Add(p, H2("텍스트 편집 (더블클릭 후)"));
        Add(p, SC("Ctrl+B",           "굵게"));
        Add(p, SC("Ctrl+I",           "기울임"));
        Add(p, SC("Ctrl+U",           "밑줄"));
        Add(p, SC("Esc",              "편집 종료"));

        Add(p, H2("서식"));
        Add(p, SC("Ctrl+Shift+F",     "글자 속성"));
        Add(p, SC("Ctrl+Shift+P",     "문단 속성"));
        Add(p, SC("Ctrl+Shift+I",     "문서 정보"));

        Add(p, H2("보기"));
        Add(p, SC("Ctrl++",           "확대"));
        Add(p, SC("Ctrl+−",           "축소"));
        Add(p, SC("Ctrl+0",           "화면에 맞춤"));

        Add(p, H2("슬라이드 쇼 진행 중"));
        Add(p, SC("Space / PageDown / →", "다음 슬라이드/애니메이션"));
        Add(p, SC("Backspace / PageUp / ←", "이전 슬라이드"));
        Add(p, SC("Esc",              "슬라이드 쇼 종료"));
        return p;
    }

    // ── UI helpers ────────────────────────────────────────────────────

    private static StackPanel P() => new();

    private static void Add(StackPanel p, UIElement el) => p.Children.Add(el);

    private static UIElement H1(string t)
    {
        var sp = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };
        sp.Children.Add(new TextBlock
        {
            Text = t, FontSize = 17, FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 0, 0, 6),
        });
        sp.Children.Add(new Separator());
        return sp;
    }

    private static TextBlock H2(string t) => new TextBlock
    {
        Text = t, FontSize = 13, FontWeight = FontWeights.Bold,
        Margin = new Thickness(0, 14, 0, 4),
    };

    private static TextBlock Para(string t) => new TextBlock
    {
        Text = t, TextWrapping = TextWrapping.Wrap,
        Margin = new Thickness(0, 3, 0, 3), LineHeight = 20,
    };

    private static TextBlock Bullet(string t) => new TextBlock
    {
        Text = "• " + t, TextWrapping = TextWrapping.Wrap,
        Margin = new Thickness(14, 2, 0, 2), LineHeight = 20,
    };

    private static Grid SC(string keys, string desc)
    {
        var g = new Grid { Margin = new Thickness(0, 2, 0, 2) };
        g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(190) });
        g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var keyBorder = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xBB, 0xBB, 0xBB)),
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(6, 1, 6, 1),
            HorizontalAlignment = HorizontalAlignment.Left,
            Child = new TextBlock
            {
                Text = keys,
                FontFamily = new FontFamily("Consolas, Courier New"),
                FontSize = 12,
            },
        };
        var descTb = new TextBlock
        {
            Text = desc,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(10, 0, 0, 0),
        };
        Grid.SetColumn(keyBorder, 0);
        Grid.SetColumn(descTb, 1);
        g.Children.Add(keyBorder);
        g.Children.Add(descTb);
        return g;
    }
}
