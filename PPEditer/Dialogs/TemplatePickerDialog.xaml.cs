using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace PPEditer.Dialogs;

public partial class TemplatePickerDialog : Window
{
    public string? SelectedPath { get; private set; }

    private readonly string _folder;

    public TemplatePickerDialog(string templatesFolder)
    {
        InitializeComponent();
        _folder = templatesFolder;
        FolderLabel.Text = templatesFolder;
        LoadTemplates();
    }

    private void LoadTemplates()
    {
        TemplateList.Items.Clear();

        if (!Directory.Exists(_folder) || Directory.GetFiles(_folder, "*.pptx").Length == 0)
        {
            NoTemplatesText.Visibility = Visibility.Visible;
            CreateBtn.IsEnabled = false;
            return;
        }

        NoTemplatesText.Visibility = Visibility.Collapsed;
        CreateBtn.IsEnabled = true;

        foreach (var file in Directory.GetFiles(_folder, "*.pptx")
                                      .OrderByDescending(File.GetLastWriteTime))
        {
            var info = new FileInfo(file);
            TemplateList.Items.Add(new TemplateItem
            {
                FilePath = file,
                FileName = info.Name,
                Modified = info.LastWriteTime.ToString("yyyy-MM-dd  HH:mm"),
            });
        }

        if (TemplateList.Items.Count > 0)
            TemplateList.SelectedIndex = 0;
    }

    private void OnBrowse(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter           = "PowerPoint 파일|*.pptx|모든 파일|*.*",
            Title            = "탬플릿 파일 선택",
            InitialDirectory = Directory.Exists(_folder) ? _folder : null,
        };
        if (dlg.ShowDialog(this) == true)
        {
            SelectedPath = dlg.FileName;
            DialogResult = true;
        }
    }

    private void OnCreate(object sender, RoutedEventArgs e)
    {
        if (TemplateList.SelectedItem is TemplateItem item)
        {
            SelectedPath = item.FilePath;
            DialogResult = true;
        }
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    private void OnListDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (TemplateList.SelectedItem is TemplateItem)
            OnCreate(sender, e);
    }
}

file class TemplateItem
{
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public string Modified { get; set; } = "";
}
