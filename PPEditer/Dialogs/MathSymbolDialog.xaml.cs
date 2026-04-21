using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PPEditer.Dialogs;

public partial class MathSymbolDialog : Window
{
    public event Action<string>? SymbolInserted;

    private static readonly (string KeyPrefix, (string Sym, string Desc)[] Items)[] _categories =
    [
        ("Dlg_Math_Cat0", // 기본 연산
        [
            ("×","곱셈"), ("÷","나눗셈"), ("±","플러스마이너스"), ("∓","마이너스플러스"),
            ("√","제곱근"), ("∛","세제곱근"), ("∜","네제곱근"), ("∞","무한대"),
            ("∝","비례"), ("∟","직각"), ("∠","각도"), ("°","도"),
            ("‰","퍼밀"), ("‱","퍼만"), ("℃","섭씨"), ("℉","화씨"),
            ("Å","옹스트롬"), ("ℏ","디랙상수"), ("ℓ","리터/길이"), ("Ω","옴"),
            ("μ","마이크로"), ("∑","합"), ("∏","곱"), ("∫","적분"),
        ]),
        ("Dlg_Math_Cat1", // 관계/부등호
        [
            ("=","같다"), ("≠","같지않다"), ("≈","근사"), ("≡","항등"), ("≃","근사같다"),
            ("≅","합동"), ("∼","유사"), ("∝","비례"), ("<","미만"), (">","초과"),
            ("≤","이하"), ("≥","이상"), ("≦","이하(변형)"), ("≧","이상(변형)"),
            ("≪","매우작다"), ("≫","매우크다"), ("≮","미만아님"), ("≯","초과아님"),
            ("⊂","부분집합"), ("⊃","상위집합"), ("⊆","부분집합(등호)"), ("⊇","상위집합(등호)"),
            ("⊄","부분집합아님"), ("⊅","상위집합아님"), ("⊊","진부분집합"), ("⊋","진상위집합"),
        ]),
        ("Dlg_Math_Cat2", // 집합/논리
        [
            ("∈","원소"), ("∉","원소아님"), ("∋","역원소"), ("∌","역원소아님"),
            ("∪","합집합"), ("∩","교집합"), ("∅","공집합"), ("⊕","대칭차"),
            ("ℕ","자연수"), ("ℤ","정수"), ("ℚ","유리수"), ("ℝ","실수"),
            ("ℂ","복소수"), ("ℍ","사원수"), ("∀","전칭한정사"), ("∃","존재한정사"),
            ("∄","존재하지않음"), ("¬","부정"), ("∧","논리곱AND"), ("∨","논리합OR"),
            ("⊻","배타적OR"), ("⊢","증명"), ("⊨","만족"), ("→","함의"),
            ("←","역함의"), ("↔","쌍조건"), ("⇒","쌍조건화살표"), ("⇔","동치"),
        ]),
        ("Dlg_Math_Cat3", // 미적분/해석
        [
            ("∂","편미분"), ("∇","나블라"), ("Δ","델타(변화량)"), ("δ","델타(변분)"),
            ("∫","적분"), ("∬","이중적분"), ("∭","삼중적분"), ("∮","선적분"),
            ("∯","면적분"), ("∰","체적적분"), ("∑","시그마합"), ("∏","파이곱"),
            ("lim","극한"), ("∞","무한대"), ("∝","비례"), ("∼","점근"),
            ("d/dt","시간미분"), ("∂/∂x","편미분x"),
        ]),
        ("Dlg_Math_Cat4", // 그리스 대문자
        [
            ("Α","Alpha"), ("Β","Beta"), ("Γ","Gamma"), ("Δ","Delta"),
            ("Ε","Epsilon"), ("Ζ","Zeta"), ("Η","Eta"), ("Θ","Theta"),
            ("Ι","Iota"), ("Κ","Kappa"), ("Λ","Lambda"), ("Μ","Mu"),
            ("Ν","Nu"), ("Ξ","Xi"), ("Ο","Omicron"), ("Π","Pi"),
            ("Ρ","Rho"), ("Σ","Sigma"), ("Τ","Tau"), ("Υ","Upsilon"),
            ("Φ","Phi"), ("Χ","Chi"), ("Ψ","Psi"), ("Ω","Omega"),
        ]),
        ("Dlg_Math_Cat5", // 그리스 소문자
        [
            ("α","alpha"), ("β","beta"), ("γ","gamma"), ("δ","delta"),
            ("ε","epsilon"), ("ζ","zeta"), ("η","eta"), ("θ","theta"),
            ("ι","iota"), ("κ","kappa"), ("λ","lambda"), ("μ","mu"),
            ("ν","nu"), ("ξ","xi"), ("ο","omicron"), ("π","pi"),
            ("ρ","rho"), ("σ","sigma"), ("τ","tau"), ("υ","upsilon"),
            ("φ","phi"), ("χ","chi"), ("ψ","psi"), ("ω","omega"),
            ("ϕ","phi(변형)"), ("ϑ","theta(변형)"), ("ϖ","pi(변형)"),
            ("ϱ","rho(변형)"), ("ϵ","epsilon(변형)"),
        ]),
        ("Dlg_Math_Cat6", // 위/아래 첨자
        [
            ("⁰","위첨자0"), ("¹","위첨자1"), ("²","위첨자2"), ("³","위첨자3"),
            ("⁴","위첨자4"), ("⁵","위첨자5"), ("⁶","위첨자6"), ("⁷","위첨자7"),
            ("⁸","위첨자8"), ("⁹","위첨자9"), ("ⁿ","위첨자n"), ("⁺","위첨자+"),
            ("⁻","위첨자-"), ("⁼","위첨자="), ("⁽","위첨자("), ("⁾","위첨자)"),
            ("₀","아래첨자0"), ("₁","아래첨자1"), ("₂","아래첨자2"), ("₃","아래첨자3"),
            ("₄","아래첨자4"), ("₅","아래첨자5"), ("₆","아래첨자6"), ("₇","아래첨자7"),
            ("₈","아래첨자8"), ("₉","아래첨자9"), ("₊","아래첨자+"), ("₋","아래첨자-"),
            ("₌","아래첨자="), ("₍","아래첨자("), ("₎","아래첨자)"), ("ₐ","아래첨자a"),
            ("ₑ","아래첨자e"), ("ₒ","아래첨자o"), ("ₓ","아래첨자x"), ("ₙ","아래첨자n"),
        ]),
        ("Dlg_Math_Cat7", // 공식 템플릿
        [
            ("E=mc²","아인슈타인 에너지-질량"),
            ("a²+b²=c²","피타고라스 정리"),
            ("F=ma","뉴턴 제2법칙"),
            ("F=Gm₁m₂/r²","만유인력 법칙"),
            ("PV=nRT","이상기체 방정식"),
            ("v=λf","파동 방정식"),
            ("E=hf","광자 에너지"),
            ("p=mv","운동량"),
            ("W=Fd","일"),
            ("P=F/A","압력"),
            ("ρ=m/V","밀도"),
            ("I=V/R","옴의 법칙"),
            ("Q=mcΔT","열량"),
            ("sin²θ+cos²θ=1","삼각함수 항등식"),
            ("eⁱᵖ+1=0","오일러 항등식"),
            ("x=(-b±√(b²-4ac))/(2a)","근의 공식"),
            ("a=Δv/Δt","가속도"),
            ("v²=v₀²+2aΔx","등가속도 운동"),
        ]),
    ];

    public MathSymbolDialog()
    {
        InitializeComponent();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        string[] catIcons = ["∑","≤","∈","∫","Σ","σ","²","E=mc²"];

        for (int i = 0; i < _categories.Length; i++)
        {
            int idx = i;
            string catKey  = _categories[i].KeyPrefix;
            string catName = Application.Current.TryFindResource(catKey) as string ?? catKey;
            string icon    = catIcons[i];

            var tab = new TabItem { Header = $"{icon} {catName}" };
            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            };
            var wrap = new WrapPanel { Margin = new Thickness(4) };

            bool isFormulas = idx == 7;

            foreach (var (sym, desc) in _categories[idx].Items)
            {
                string s = sym, d = desc;
                var btn = new Button
                {
                    Content         = s,
                    FontSize        = isFormulas ? 13 : 18,
                    Width           = isFormulas ? double.NaN : 44,
                    Height          = isFormulas ? double.NaN : 44,
                    Margin          = new Thickness(2),
                    Padding         = isFormulas ? new Thickness(8, 4, 8, 4) : new Thickness(2),
                    ToolTip         = d,
                    FontFamily      = new FontFamily("Cambria Math, Segoe UI Symbol, Arial Unicode MS"),
                };
                btn.Click       += (_, _) => SymbolInserted?.Invoke(s);
                btn.MouseEnter  += (_, _) => { TbPreview.Text = s; TbInfo.Text = d; };
                wrap.Children.Add(btn);
            }

            scroll.Content = wrap;
            tab.Content    = scroll;
            TabCtrl.Items.Add(tab);
        }
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}
