using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace PPEditer.Dialogs;

public partial class CharMapDialog : Window
{
    public event Action<string>? CharInserted;

    // ── Character data: (character, english name) ──────────────────────
    private static readonly (string[] Keys, (string C, string N)[] Chars)[] _data =
    [
        (["Dlg_CharMap_Cat0"], // General
        [
            ("©","Copyright"), ("®","Registered"), ("™","Trade Mark"),
            ("§","Section"), ("¶","Pilcrow"), ("†","Dagger"), ("‡","Double Dagger"),
            ("•","Bullet"), ("…","Ellipsis"), ("–","En Dash"), ("—","Em Dash"),
            ("“","Left Quote “"), ("”","Right Quote ”"),
            ("‘","Left Quote ‘"), ("’","Right Quote ’"),
            ("«","Left Double Angle"), ("»","Right Double Angle"),
            ("‰","Per Mille"), ("′","Prime"), ("″","Double Prime"), ("‴","Triple Prime"),
            ("№","Numero"), ("℃","Degree Celsius"), ("℉","Degree Fahrenheit"),
            ("☎","Telephone"), ("✉","Envelope"), ("✂","Scissors"),
            ("✏","Pencil"), ("✒","Black Nib"), ("⌚","Watch"), ("⌛","Hourglass"),
            ("☀","Sun"), ("☁","Cloud"), ("☂","Umbrella"), ("☃","Snowman"),
        ]),
        (["Dlg_CharMap_Cat1"], // Math
        [
            ("+","Plus"), ("−","Minus"), ("×","Multiplication"), ("÷","Division"),
            ("=","Equal"), ("≠","Not Equal"), ("<","Less Than"), (">","Greater Than"),
            ("≤","Less Than or Equal"), ("≥","Greater Than or Equal"),
            ("≈","Almost Equal"), ("≡","Identical To"), ("≃","Asymptotically Equal"),
            ("∞","Infinity"), ("√","Square Root"), ("∛","Cube Root"),
            ("∑","Summation"), ("∏","Product"), ("∫","Integral"),
            ("∂","Partial Differential"), ("∇","Nabla"), ("∆","Increment"),
            ("∈","Element Of"), ("∉","Not Element Of"), ("∋","Contains"),
            ("∩","Intersection"), ("∪","Union"), ("⊂","Subset"), ("⊃","Superset"),
            ("⊆","Subset or Equal"), ("⊇","Superset or Equal"),
            ("⊕","Circled Plus"), ("⊗","Circled Times"), ("⊥","Perpendicular"),
            ("∝","Proportional To"), ("∀","For All"), ("∃","There Exists"),
            ("∧","Logical And"), ("∨","Logical Or"), ("¬","Not Sign"),
            ("±","Plus Minus"), ("∓","Minus Plus"), ("∠","Angle"), ("∟","Right Angle"),
            ("°","Degree"), ("∅","Empty Set"), ("π","Pi"),
        ]),
        (["Dlg_CharMap_Cat2"], // Greek
        [
            ("α","Alpha"), ("β","Beta"), ("γ","Gamma"), ("δ","Delta"), ("ε","Epsilon"),
            ("ζ","Zeta"), ("η","Eta"), ("θ","Theta"), ("ι","Iota"), ("κ","Kappa"),
            ("λ","Lambda"), ("μ","Mu"), ("ν","Nu"), ("ξ","Xi"), ("π","Pi"),
            ("ρ","Rho"), ("σ","Sigma"), ("τ","Tau"), ("υ","Upsilon"),
            ("φ","Phi"), ("χ","Chi"), ("ψ","Psi"), ("ω","Omega"),
            ("Α","Alpha UC"), ("Β","Beta UC"), ("Γ","Gamma UC"), ("Δ","Delta UC"),
            ("Ε","Epsilon UC"), ("Ζ","Zeta UC"), ("Η","Eta UC"), ("Θ","Theta UC"),
            ("Ι","Iota UC"), ("Κ","Kappa UC"), ("Λ","Lambda UC"), ("Μ","Mu UC"),
            ("Ν","Nu UC"), ("Ξ","Xi UC"), ("Π","Pi UC"), ("Ρ","Rho UC"),
            ("Σ","Sigma UC"), ("Τ","Tau UC"), ("Υ","Upsilon UC"),
            ("Φ","Phi UC"), ("Χ","Chi UC"), ("Ψ","Psi UC"), ("Ω","Omega UC"),
        ]),
        (["Dlg_CharMap_Cat3"], // Currency
        [
            ("$","Dollar"), ("€","Euro"), ("£","Pound"), ("¥","Yen/Yuan"),
            ("₩","Won"), ("¢","Cent"), ("₽","Ruble"), ("₿","Bitcoin"),
            ("₹","Rupee"), ("₺","Lira"), ("₸","Tenge"), ("₦","Naira"),
            ("₴","Hryvnia"), ("₫","Dong"), ("₭","Kip"), ("₱","Peso"),
            ("₲","Guarani"), ("₼","Manat"), ("₾","Lari"), ("¤","Generic Currency"),
        ]),
        (["Dlg_CharMap_Cat4"], // Arrows
        [
            ("←","Left"), ("→","Right"), ("↑","Up"), ("↓","Down"),
            ("↔","Left Right"), ("↕","Up Down"),
            ("↖","NW"), ("↗","NE"), ("↘","SE"), ("↙","SW"),
            ("↩","Left Hook"), ("↪","Right Hook"),
            ("⇐","Left Double"), ("⇒","Right Double"),
            ("⇑","Up Double"), ("⇓","Down Double"),
            ("⇔","LR Double"), ("⇕","UD Double"),
            ("➡","Right Bold"), ("⬅","Left Bold"),
            ("⬆","Up Bold"), ("⬇","Down Bold"),
            ("↺","Anticlockwise"), ("↻","Clockwise"),
            ("⟵","Long Left"), ("⟶","Long Right"), ("⟷","Long LR"),
            ("↤","Maps From"), ("↦","Maps To"), ("⇢","Right Dotted"),
        ]),
        (["Dlg_CharMap_Cat5"], // Shapes & Symbols
        [
            ("■","Black Square"), ("□","White Square"),
            ("▪","Small Black Sq"), ("▫","Small White Sq"),
            ("▲","Up Triangle"), ("△","White Up Tri"),
            ("▼","Down Triangle"), ("▽","White Down Tri"),
            ("◀","Left Triangle"), ("▶","Right Triangle"),
            ("◆","Black Diamond"), ("◇","White Diamond"),
            ("●","Black Circle"), ("○","White Circle"), ("◎","Bull's Eye"),
            ("★","Black Star"), ("☆","White Star"),
            ("✓","Check Mark"), ("✗","Ballot X"),
            ("✔","Heavy Check"), ("✘","Heavy Ballot X"),
            ("✕","Multiply X"), ("✖","Heavy Multiply X"),
            ("♠","Spade"), ("♥","Heart"), ("♦","Diamond"), ("♣","Club"),
            ("♡","White Heart"), ("♢","White Diamond"),
            ("♪","Eighth Note"), ("♫","Beamed Notes"),
            ("⚡","Lightning"), ("❄","Snowflake"), ("✿","Flower"),
            ("❤","Heavy Heart"), ("❥","Rotated Heart"),
        ]),
        (["Dlg_CharMap_Cat6"], // Fractions & Super/Subscript
        [
            ("½","½"), ("¼","¼"), ("¾","¾"), ("⅓","⅓"), ("⅔","⅔"),
            ("⅛","⅛"), ("⅜","⅜"), ("⅝","⅝"), ("⅞","⅞"),
            ("⅐","1/7"), ("⅑","1/9"), ("⅒","1/10"),
            ("⁰","⁰ Sup"), ("¹","¹ Sup"), ("²","² Sup"), ("³","³ Sup"),
            ("⁴","⁴ Sup"), ("⁵","⁵ Sup"), ("⁶","⁶ Sup"), ("⁷","⁷ Sup"),
            ("⁸","⁸ Sup"), ("⁹","⁹ Sup"), ("⁺","⁺ Sup"), ("⁻","⁻ Sup"),
            ("ⁿ","ⁿ Sup"), ("ⁱ","ⁱ Sup"),
            ("₀","₀ Sub"), ("₁","₁ Sub"), ("₂","₂ Sub"), ("₃","₃ Sub"),
            ("₄","₄ Sub"), ("₅","₅ Sub"), ("₆","₆ Sub"), ("₇","₇ Sub"),
            ("₈","₈ Sub"), ("₉","₉ Sub"), ("₊","₊ Sub"), ("₋","₋ Sub"),
            ("ₐ","ₐ Sub"), ("ₑ","ₑ Sub"), ("ₒ","ₒ Sub"),
        ]),
    ];

    public CharMapDialog()
    {
        InitializeComponent();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        for (int i = 0; i < _data.Length; i++)
        {
            var (keys, chars) = _data[i];
            string header = Application.Current.TryFindResource(keys[0]) as string ?? keys[0];

            var wrap = new WrapPanel { Margin = new Thickness(4) };
            foreach (var (c, name) in chars)
            {
                string ch = c, nm = name;
                var btn = new Button
                {
                    Content    = ch,
                    Width      = 40,
                    Height     = 40,
                    FontSize   = 18,
                    FontFamily = new FontFamily("Segoe UI Symbol, Segoe UI Emoji, Arial Unicode MS"),
                    Margin     = new Thickness(1),
                    ToolTip    = nm,
                };
                btn.Click      += (_, _) => { CharInserted?.Invoke(ch); };
                btn.MouseEnter += (_, _) =>
                {
                    TbPreview.Text = ch;
                    uint cp = (ch.Length >= 2 && char.IsHighSurrogate(ch[0]))
                        ? (uint)char.ConvertToUtf32(ch[0], ch[1])
                        : (uint)ch[0];
                    TbInfo.Text = $"U+{cp:X4}  {nm}";
                };
                wrap.Children.Add(btn);
            }

            TabCtrl.Items.Add(new TabItem
            {
                Header  = header,
                Content = new ScrollViewer
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    Content = wrap,
                },
            });
        }
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}
