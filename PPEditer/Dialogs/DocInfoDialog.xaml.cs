using System.Windows;
using System.Windows.Media;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class DocInfoDialog : Window
{
    // ── Output ────────────────────────────────────────────────────────────

    public DocProperties? Result        { get; private set; }
    public bool           SetProtect    { get; private set; }
    public bool           RemoveProtect { get; private set; }
    public string?        WritePassword { get; private set; }

    // ── State ─────────────────────────────────────────────────────────────

    private readonly bool _initiallyProtected;

    private enum PendingAction { None, Set, Remove }
    private PendingAction _pending = PendingAction.None;
    private string?       _stagedPassword;

    // ── Constructor ───────────────────────────────────────────────────────

    public DocInfoDialog(DocProperties props, bool hasWriteProtection)
    {
        InitializeComponent();
        _initiallyProtected = hasWriteProtection;

        // General tab
        TxTitle.Text   = props.Title;
        TxSubject.Text = props.Subject;
        TxAuthor.Text  = props.Author;
        TxManager.Text = props.Manager;
        TxCompany.Text = props.Company;

        LblLastModifiedBy.Text = string.IsNullOrEmpty(props.LastModifiedBy) ? "-" : props.LastModifiedBy;
        LblCreated.Text        = props.Created?.ToLocalTime().ToString("yyyy-MM-dd HH:mm") ?? "-";
        LblModified.Text       = props.Modified?.ToLocalTime().ToString("yyyy-MM-dd HH:mm") ?? "-";
        LblRevision.Text       = props.Revision > 0 ? props.Revision.ToString() : "-";

        // Security tab
        SyncProtectionUI();
    }

    // ── Security tab helpers ──────────────────────────────────────────────

    private void SyncProtectionUI()
    {
        bool isProtected = _pending == PendingAction.Set
                        || (_initiallyProtected && _pending != PendingAction.Remove);

        string statusText;
        Brush  statusColor;

        if (_pending == PendingAction.Set)
        {
            statusText  = Res("Dlg_DocInfo_Protected") + " (*)";
            statusColor = new SolidColorBrush(Color.FromRgb(0, 128, 0));
        }
        else if (_pending == PendingAction.Remove)
        {
            statusText  = Res("Dlg_DocInfo_NotProtected") + " (*)";
            statusColor = new SolidColorBrush(Colors.OrangeRed);
        }
        else if (_initiallyProtected)
        {
            statusText  = Res("Dlg_DocInfo_Protected");
            statusColor = new SolidColorBrush(Color.FromRgb(0, 128, 0));
        }
        else
        {
            statusText  = Res("Dlg_DocInfo_NotProtected");
            statusColor = new SolidColorBrush(Colors.Gray);
        }

        LblProtectStatus.Text       = statusText;
        LblProtectStatus.Foreground = statusColor;

        BtnSetProtect.Content    = _initiallyProtected
            ? Res("Dlg_DocInfo_ChangePwd")
            : Res("Dlg_DocInfo_SetPwd");
        BtnRemoveProtect.Content   = Res("Dlg_DocInfo_RemovePwd");
        BtnRemoveProtect.IsEnabled = isProtected;
    }

    // ── Security button handlers ──────────────────────────────────────────

    private void OnSetWriteProtect(object sender, RoutedEventArgs e)
    {
        var pw1 = PwNew.Password;
        var pw2 = PwConfirm.Password;

        if (string.IsNullOrEmpty(pw1))
        {
            MessageBox.Show(this, Res("Dlg_DocInfo_PwdEmpty"), Res("Dlg_DocInfo_WriteProtect"),
                            MessageBoxButton.OK, MessageBoxImage.Warning);
            PwNew.Focus();
            return;
        }
        if (pw1 != pw2)
        {
            MessageBox.Show(this, Res("Dlg_PasswordMismatch"), Res("Dlg_DocInfo_WriteProtect"),
                            MessageBoxButton.OK, MessageBoxImage.Warning);
            PwConfirm.Clear();
            PwConfirm.Focus();
            return;
        }

        _pending       = PendingAction.Set;
        _stagedPassword = pw1;
        PwNew.Clear();
        PwConfirm.Clear();
        SyncProtectionUI();
    }

    private void OnRemoveWriteProtect(object sender, RoutedEventArgs e)
    {
        _pending        = PendingAction.Remove;
        _stagedPassword = null;
        PwNew.Clear();
        PwConfirm.Clear();
        SyncProtectionUI();
    }

    // ── Dialog result handlers ────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        Result = new DocProperties
        {
            Title   = TxTitle.Text.Trim(),
            Subject = TxSubject.Text.Trim(),
            Author  = TxAuthor.Text.Trim(),
            Manager = TxManager.Text.Trim(),
            Company = TxCompany.Text.Trim(),
        };

        SetProtect    = _pending == PendingAction.Set;
        RemoveProtect = _pending == PendingAction.Remove;
        WritePassword = SetProtect ? _stagedPassword : null;

        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        Result       = null;
        DialogResult = false;
    }

    // ── Resource helper ───────────────────────────────────────────────────

    private static string Res(string key)
        => Application.Current.TryFindResource(key) is string s ? s : key;
}
