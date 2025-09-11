using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using System.Reflection;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class OfficeModule : BaseSimpleModule
    {
        private enum OfficeApp { Word, Excel }
        private OfficeApp _current = OfficeApp.Word;
        private Word.Application _wordApp;
        private Excel.Application _excelApp;

        private WindowsFormsHost _host;
        private System.Windows.Forms.Panel _panel;
        private TextBlock _txtStatus;
        private TextBlock _txtInfo;

        [DllImport("user32.dll")] static extern bool SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        [DllImport("user32.dll")] static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hWnd);

        private const int GWL_STYLE = -16;
        private const int WS_CAPTION = 0x00C00000;
        private const int WS_THICKFRAME = 0x00040000;
        private const int WS_CHILD = 0x40000000;
        private const int SW_MAXIMIZE = 3;

        public OfficeModule()
        {
            InitializeComponent();
            _txtStatus = FindName("txtStatus") as TextBlock;
            _txtInfo = FindName("txtInfo") as TextBlock;
            cmbApp.SelectedIndex = 0;
            InitializeHostContainer();
            UpdateStatus();
        }

        private void InitializeHostContainer()
        {
            try
            {
                _host = new WindowsFormsHost();
                _panel = new System.Windows.Forms.Panel { Dock = System.Windows.Forms.DockStyle.Fill, BackColor = System.Drawing.Color.White };
                _host.Child = _panel;
                if (hostContainer != null)
                {
                    hostContainer.Children.Clear();
                    hostContainer.Children.Add(_host);
                    hostContainer.SizeChanged += (s, e) => ResizeEmbeddedOffice();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Init Host Fehler: {ex.Message}");
                MessageBox.Show($"Fehler beim Initialisieren des Office-Containers: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void cmbApp_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbApp.SelectedIndex < 0) return;
            var old = _current;
            _current = (cmbApp.SelectedIndex == 1) ? OfficeApp.Excel : OfficeApp.Word;
            if (old != _current)
            {
                CleanupExistingOffice();
                UpdateStatus();
            }
        }

        private void btnNew_Click(object sender, RoutedEventArgs e) => CreateNewDocument();
        private void btnOpen_Click(object sender, RoutedEventArgs e) => OpenDocument();
        private void btnSave_Click(object sender, RoutedEventArgs e) => SaveDocument(false);
        private void btnSaveAs_Click(object sender, RoutedEventArgs e) => SaveDocument(true);
        private void btnBold_Click(object sender, RoutedEventArgs e) => ApplyFormatting("Bold");
        private void btnItalic_Click(object sender, RoutedEventArgs e) => ApplyFormatting("Italic");
        private void btnUnderline_Click(object sender, RoutedEventArgs e) => ApplyFormatting("Underline");
        private void btnAlignLeft_Click(object sender, RoutedEventArgs e) => ApplyFormatting("AlignLeft");
        private void btnAlignCenter_Click(object sender, RoutedEventArgs e) => ApplyFormatting("AlignCenter");
        private void btnAlignRight_Click(object sender, RoutedEventArgs e) => ApplyFormatting("AlignRight");

        private void cmbSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbSize.SelectedItem is ComboBoxItem item && double.TryParse(item.Content.ToString(), out double size))
                ApplyFormatting("FontSize", size);
        }

        private void CreateNewDocument()
        {
            try
            {
                EnsureOfficeCompletelyTerminated();
                if (_current == OfficeApp.Word)
                {
                    _wordApp = new Word.Application { Visible = false };
                    _wordApp.Documents.Add();
                    Thread.Sleep(500);
                    _wordApp.Visible = true;
                    Thread.Sleep(200);
                    var hwnd = GetWordHwnd();
                    if (hwnd != IntPtr.Zero) EmbedOfficeWindow(hwnd);
                }
                else
                {
                    _excelApp = new Excel.Application { Visible = false };
                    _excelApp.Workbooks.Add();
                    Thread.Sleep(500);
                    _excelApp.Visible = true;
                    Thread.Sleep(200);
                    var hwnd = GetExcelHwnd();
                    if (hwnd != IntPtr.Zero) PositionExcelWindow(hwnd);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"New Doc Fehler: {ex.Message}");
                MessageBox.Show($"Fehler beim Erstellen: {ex.Message}");
            }
        }

        private void OpenDocument()
        {
            var filter = _current == OfficeApp.Word ? "Word-Dokumente (*.doc;*.docx)|*.doc;*.docx|Alle Dateien (*.*)|*.*" : "Excel-Tabellen (*.xls;*.xlsx)|*.xls;*.xlsx|Alle Dateien (*.*)|*.*";
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = filter };
            if (dlg.ShowDialog() != true) return;
            try
            {
                EnsureOfficeCompletelyTerminated();
                if (_current == OfficeApp.Word)
                {
                    _wordApp = new Word.Application { Visible = false };
                    _wordApp.Documents.Open(dlg.FileName);
                    Thread.Sleep(500);
                    _wordApp.Visible = true;
                    Thread.Sleep(200);
                    var hwnd = GetWordHwnd();
                    if (hwnd != IntPtr.Zero) EmbedOfficeWindow(hwnd);
                }
                else
                {
                    _excelApp = new Excel.Application { Visible = false };
                    _excelApp.Workbooks.Open(dlg.FileName);
                    Thread.Sleep(500);
                    _excelApp.Visible = true;
                    Thread.Sleep(200);
                    var hwnd = GetExcelHwnd();
                    if (hwnd != IntPtr.Zero) PositionExcelWindow(hwnd);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Open Doc Fehler: {ex.Message}");
                MessageBox.Show($"Fehler beim Öffnen: {ex.Message}");
            }
        }

        private void SaveDocument(bool saveAs)
        {
            try
            {
                if (_current == OfficeApp.Word && _wordApp != null)
                {
                    if (saveAs)
                    {
                        var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Word-Dokumente (*.docx)|*.docx|Alle Dateien (*.*)|*.*", DefaultExt = ".docx" };
                        if (dlg.ShowDialog() == true && _wordApp.ActiveDocument != null) _wordApp.ActiveDocument.SaveAs2(dlg.FileName);
                    }
                    else if (_wordApp.ActiveDocument != null) _wordApp.ActiveDocument.Save();
                }
                else if (_current == OfficeApp.Excel && _excelApp != null)
                {
                    if (saveAs)
                    {
                        var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Excel-Tabellen (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*", DefaultExt = ".xlsx" };
                        if (dlg.ShowDialog() == true && _excelApp.ActiveWorkbook != null) _excelApp.ActiveWorkbook.SaveAs(dlg.FileName);
                    }
                    else if (_excelApp.ActiveWorkbook != null) _excelApp.ActiveWorkbook.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Speichern: {ex.Message}");
            }
        }

        private void ApplyFormatting(string format, object value = null)
        {
            try
            {
                if (_current == OfficeApp.Word && _wordApp != null)
                {
                    var sel = _wordApp.Selection;
                    if (sel == null) return;
                    switch (format)
                    {
                        case "Bold": sel.Font.Bold = sel.Font.Bold == 0 ? 1 : 0; break;
                        case "Italic": sel.Font.Italic = sel.Font.Italic == 0 ? 1 : 0; break;
                        case "Underline": sel.Font.Underline = sel.Font.Underline == Word.WdUnderline.wdUnderlineNone ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone; break;
                        case "FontSize": if (value is double sz) sel.Font.Size = (float)sz; break;
                        case "AlignLeft": sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                        case "AlignCenter": sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                        case "AlignRight": sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                    }
                }
                else if (_current == OfficeApp.Excel && _excelApp != null)
                {
                    var cell = _excelApp.ActiveCell; if (cell == null) return;
                    switch (format)
                    {
                        case "Bold": cell.Font.Bold = !cell.Font.Bold; break;
                        case "Italic": cell.Font.Italic = !cell.Font.Italic; break;
                        case "Underline": cell.Font.Underline = cell.Font.Underline == Excel.XlUnderlineStyle.xlUnderlineStyleNone ? Excel.XlUnderlineStyle.xlUnderlineStyleSingle : Excel.XlUnderlineStyle.xlUnderlineStyleNone; break;
                        case "FontSize": if (value is double sz) cell.Font.Size = sz; break;
                        case "AlignLeft": cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; break;
                        case "AlignCenter": cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; break;
                        case "AlignRight": cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; break;
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show($"Fehler bei Formatierung: {ex.Message}"); }
        }

        private IntPtr GetWordHwnd()
        {
            if (_wordApp == null) return IntPtr.Zero;
            try { dynamic w = _wordApp.ActiveWindow; if (w != null) { int h = w.Hwnd; if (h != 0) return new IntPtr(h); } } catch { }
            try { var p = _wordApp.GetType().GetProperty("Hwnd"); if (p != null) { var v = p.GetValue(_wordApp); if (v is int i && i != 0) return new IntPtr(i); } } catch { }
            return IntPtr.Zero;
        }
        private IntPtr GetExcelHwnd()
        {
            if (_excelApp == null) return IntPtr.Zero;
            try { int h = _excelApp.Hwnd; if (h != 0) return new IntPtr(h); } catch { }
            return IntPtr.Zero;
        }

        private void EmbedOfficeWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || _panel?.Handle == IntPtr.Zero) return;
            try
            {
                int style = GetWindowLong(hwnd, GWL_STYLE);
                style &= ~WS_CAPTION; style &= ~WS_THICKFRAME; style |= WS_CHILD;
                SetWindowLong(hwnd, GWL_STYLE, style);
                SetParent(hwnd, _panel.Handle);
                ShowWindow(hwnd, SW_MAXIMIZE);
                ResizeEmbeddedOffice();
                SetForegroundWindow(hwnd);
            }
            catch { }
        }

        private void PositionExcelWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !IsWindow(hwnd)) return;
            try
            {
                var pt = _panel.PointToScreen(new System.Drawing.Point(0, 0));
                MoveWindow(hwnd, pt.X, pt.Y, _panel.Width, _panel.Height, true);
            }
            catch { }
        }

        private void ResizeEmbeddedOffice()
        {
            IntPtr hwnd = IntPtr.Zero;
            if (_current == OfficeApp.Word && _wordApp != null) hwnd = GetWordHwnd();
            else if (_current == OfficeApp.Excel && _excelApp != null) hwnd = GetExcelHwnd();
            if (hwnd != IntPtr.Zero && IsWindow(hwnd))
            {
                try { MoveWindow(hwnd, 0, 0, _panel.Width, _panel.Height, true); } catch { }
            }
        }

        private void EnsureOfficeCompletelyTerminated()
        {
            CleanupExistingOffice();
            Thread.Sleep(500);
        }

        private void CleanupExistingOffice()
        {
            if (_wordApp != null) { try { _wordApp.Quit(false); } catch { } try { System.Runtime.InteropServices.Marshal.ReleaseComObject(_wordApp); } catch { } _wordApp = null; }
            if (_excelApp != null) { try { _excelApp.Quit(); } catch { } try { System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp); } catch { } _excelApp = null; }
        }

        public override void SetVisible(bool visible)
        {
            base.SetVisible(visible);
            if (!visible) CleanupExistingOffice();
        }

        private void UpdateStatus()
        {
            if (_txtStatus == null || _txtInfo == null) return;
            string appName = _current == OfficeApp.Word ? "Word" : "Excel";
            _txtStatus.Text = $"MS Office: {appName}";
            _txtInfo.Text = $"{appName}-Dokument";
        }
    }
}