using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using Microsoft.Win32;
using System.Reflection;
using System.Threading;
using Vivit_Control_Center.Settings;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class OfficeModule : BaseSimpleModule
    {
        private enum OfficeApp { Word, Excel }
        
        // Office-Konfiguration
        private OfficeApp _current = OfficeApp.Word;
        private string _suite;
        private AppSettings _settings;
        
        // Office COM-Objekte
        private Word.Application _wordApp;
        private Excel.Application _excelApp;
        private uint _librePid;
        
        // UI-Container
        private WindowsFormsHost _host;
        private System.Windows.Forms.Panel _panel;
        
        // UI-Elemente
        private TextBlock _txtStatus;
        private TextBlock _txtInfo;
        
        // Win32 APIs für das Fensterhandling
        [DllImport("user32.dll")] static extern bool SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        [DllImport("user32.dll")] static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildCallback lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
        
        private delegate bool EnumChildCallback(IntPtr hwnd, IntPtr lParam);
        
        private const int GWL_STYLE = -16;
        private const int WS_CAPTION = 0x00C00000;
        private const int WS_THICKFRAME = 0x00040000;
        private const int WS_CHILD = 0x40000000;
        private const int SW_MAXIMIZE = 3;
        private const int SW_SHOWNA = 8;

        public OfficeModule()
        {
            InitializeComponent();
            
            // UI-Elemente aus XAML referenzieren
            _txtStatus = FindName("txtStatus") as TextBlock;
            _txtInfo = FindName("txtInfo") as TextBlock;
            
            _settings = AppSettings.Load();
            _suite = _settings.OfficeSuite ?? "MSOffice";
            cmbApp.SelectedIndex = 0;
            
            // Panel erstellen, das als Container für Office dient
            InitializeHostContainer();
            
            // Status aktualisieren
            UpdateStatus();
        }
        
        private void InitializeHostContainer()
        {
            try
            {
                // Windows Forms Host und Panel initialisieren
                _host = new WindowsFormsHost();
                _panel = new System.Windows.Forms.Panel();
                _panel.Dock = System.Windows.Forms.DockStyle.Fill;
                _panel.BackColor = System.Drawing.Color.White;
                
                // Panel als Kind des WindowsFormsHost setzen
                _host.Child = _panel;
                
                // Host zum Container hinzufügen
                if (hostContainer != null)
                {
                    hostContainer.Children.Clear();
                    hostContainer.Children.Add(_host);
                    
                    // Event-Handler für Größenänderungen
                    hostContainer.SizeChanged += (s, e) => ResizeEmbeddedOffice();
                }
                
                Debug.WriteLine("WindowsFormsHost und Panel erfolgreich initialisiert");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler bei der Initialisierung des WindowsFormsHost: {ex.Message}");
                System.Windows.MessageBox.Show($"Fehler beim Initialisieren des Office-Containers: {ex.Message}", 
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        // Benutzeroberflächen-Events
        private void cmbApp_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbApp.SelectedIndex < 0) return;
            
            OfficeApp oldApp = _current;
            _current = (cmbApp.SelectedIndex == 1) ? OfficeApp.Excel : OfficeApp.Word;
            
            // Nur neu initialisieren, wenn sich die App geändert hat
            if (oldApp != _current)
            {
                CleanupExistingOffice();
                UpdateStatus();
            }
        }
        
        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            CreateNewDocument();
        }
        
        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenDocument();
        }
        
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveDocument(false);
        }
        
        private void btnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            SaveDocument(true);
        }
        
        private void btnBold_Click(object sender, RoutedEventArgs e) => ApplyFormatting("Bold");
        private void btnItalic_Click(object sender, RoutedEventArgs e) => ApplyFormatting("Italic");
        private void btnUnderline_Click(object sender, RoutedEventArgs e) => ApplyFormatting("Underline");
        private void btnAlignLeft_Click(object sender, RoutedEventArgs e) => ApplyFormatting("AlignLeft");
        private void btnAlignCenter_Click(object sender, RoutedEventArgs e) => ApplyFormatting("AlignCenter");
        private void btnAlignRight_Click(object sender, RoutedEventArgs e) => ApplyFormatting("AlignRight");
        
        private void cmbSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(cmbSize.SelectedItem is ComboBoxItem item) || 
                !double.TryParse(item.Content.ToString(), out double size))
                return;
                
            ApplyFormatting("FontSize", size);
        }
        
        // Dokument-Operationen
        private void CreateNewDocument()
        {
            try
            {
                // Wichtig: Stelle sicher, dass die alte Office-Instanz komplett beendet ist
                EnsureOfficeCompletelyTerminated();
                
                if (IsLibre())
                {
                    // LibreOffice in einem Panel hosten
                    string app = _current == OfficeApp.Word ? "--writer" : "--calc";
                    LaunchLibreOffice(app);
                    return;
                }
                
                // MS Office als COM-Objekt verwenden
                if (_current == OfficeApp.Word)
                {
                    _wordApp = new Word.Application();
                    _wordApp.Visible = false; // Erst unsichtbar starten
                    _wordApp.Documents.Add();
                    
                    // Kurz warten, damit Word vollständig laden kann
                    Thread.Sleep(500);
                    
                    // Jetzt sichtbar machen und einbetten
                    _wordApp.Visible = true;
                    Thread.Sleep(200);
                    
                    // Word-Fenster in Panel einbetten
                    IntPtr wordHwnd = GetWordHwnd();
                    if (wordHwnd != IntPtr.Zero)
                    {
                        Debug.WriteLine($"Word HWND gefunden: 0x{wordHwnd.ToInt64():X}");
                        EmbedOfficeWindow(wordHwnd);
                    }
                    else
                    {
                        Debug.WriteLine("Word HWND konnte nicht gefunden werden.");
                    }
                }
                else
                {
                    // Für Excel verwenden wir die Position/Größen-Methode statt Einbettung
                    // um das Ribbon-Problem zu vermeiden
                    _excelApp = new Excel.Application();
                    _excelApp.Visible = false; // Erst unsichtbar starten
                    _excelApp.Workbooks.Add();
                    
                    // Kurz warten, damit Excel vollständig laden kann
                    Thread.Sleep(500);
                    
                    // Jetzt sichtbar machen
                    _excelApp.Visible = true;
                    Thread.Sleep(200);
                    
                    // Excel-Fenster finden und als Kind einbetten
                    IntPtr excelHwnd = GetExcelHwnd();
                    if (excelHwnd != IntPtr.Zero)
                    {
                        Debug.WriteLine($"Excel HWND gefunden: 0x{excelHwnd.ToInt64():X}");
                        PositionExcelWindow(excelHwnd);
                    }
                    else
                    {
                        Debug.WriteLine("Excel HWND konnte nicht gefunden werden.");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Erstellen eines neuen Dokuments: {ex.Message}");
                System.Windows.MessageBox.Show($"Fehler beim Erstellen eines neuen Dokuments: {ex.Message}", 
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void OpenDocument()
        {
            string filter = _current == OfficeApp.Word 
                ? "Word-Dokumente (*.doc;*.docx)|*.doc;*.docx|Alle Dateien (*.*)|*.*" 
                : "Excel-Tabellen (*.xls;*.xlsx)|*.xls;*.xlsx|Alle Dateien (*.*)|*.*";
                
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = filter };
            if (dlg.ShowDialog() != true) return;
            
            try
            {
                // Wichtig: Stelle sicher, dass die alte Office-Instanz komplett beendet ist
                EnsureOfficeCompletelyTerminated();
                
                if (IsLibre())
                {
                    // LibreOffice mit Datei starten
                    LaunchLibreOffice($"\"{dlg.FileName}\"");
                    return;
                }
                
                // MS Office-Datei öffnen
                if (_current == OfficeApp.Word)
                {
                    _wordApp = new Word.Application();
                    _wordApp.Visible = false; // Erst unsichtbar starten
                    _wordApp.Documents.Open(dlg.FileName);
                    
                    // Kurz warten, damit Word vollständig laden kann
                    Thread.Sleep(500);
                    
                    // Jetzt sichtbar machen und einbetten
                    _wordApp.Visible = true;
                    Thread.Sleep(200);
                    
                    // Word-Fenster einbetten
                    IntPtr wordHwnd = GetWordHwnd();
                    if (wordHwnd != IntPtr.Zero)
                    {
                        EmbedOfficeWindow(wordHwnd);
                    }
                }
                else
                {
                    // Für Excel verwenden wir die Position/Größen-Methode statt Einbettung
                    _excelApp = new Excel.Application();
                    _excelApp.Visible = false; // Erst unsichtbar starten
                    _excelApp.Workbooks.Open(dlg.FileName);
                    
                    // Kurz warten, damit Excel vollständig laden kann
                    Thread.Sleep(500);
                    
                    // Jetzt sichtbar machen
                    _excelApp.Visible = true;
                    Thread.Sleep(200);
                    
                    // Excel-Fenster finden und positionieren
                    IntPtr excelHwnd = GetExcelHwnd();
                    if (excelHwnd != IntPtr.Zero)
                    {
                        PositionExcelWindow(excelHwnd);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Öffnen der Datei: {ex.Message}");
                System.Windows.MessageBox.Show($"Fehler beim Öffnen der Datei: {ex.Message}", 
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void SaveDocument(bool saveAs)
        {
            if (IsLibre())
            {
                System.Windows.MessageBox.Show("Diese Funktion ist für LibreOffice nicht verfügbar.", 
                    "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                if (_current == OfficeApp.Word && _wordApp != null)
                {
                    if (saveAs)
                    {
                        var dlg = new Microsoft.Win32.SaveFileDialog
                        {
                            Filter = "Word-Dokumente (*.docx)|*.docx|Alle Dateien (*.*)|*.*",
                            DefaultExt = ".docx"
                        };
                        
                        if (dlg.ShowDialog() == true && _wordApp.ActiveDocument != null)
                            _wordApp.ActiveDocument.SaveAs2(dlg.FileName);
                    }
                    else if (_wordApp.ActiveDocument != null)
                    {
                        _wordApp.ActiveDocument.Save();
                    }
                }
                else if (_current == OfficeApp.Excel && _excelApp != null)
                {
                    if (saveAs)
                    {
                        var dlg = new Microsoft.Win32.SaveFileDialog
                        {
                            Filter = "Excel-Tabellen (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*",
                            DefaultExt = ".xlsx"
                        };
                        
                        if (dlg.ShowDialog() == true && _excelApp.ActiveWorkbook != null)
                            _excelApp.ActiveWorkbook.SaveAs(dlg.FileName);
                    }
                    else if (_excelApp.ActiveWorkbook != null)
                    {
                        _excelApp.ActiveWorkbook.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Speichern: {ex.Message}");
                System.Windows.MessageBox.Show($"Fehler beim Speichern: {ex.Message}", 
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void ApplyFormatting(string format, object value = null)
        {
            if (IsLibre())
            {
                System.Windows.MessageBox.Show("Diese Funktion ist für LibreOffice nicht verfügbar.", 
                    "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                if (_current == OfficeApp.Word && _wordApp != null)
                {
                    var selection = _wordApp.Selection;
                    if (selection == null) return;

                    switch (format)
                    {
                        case "Bold":
                            selection.Font.Bold = selection.Font.Bold == 0 ? 1 : 0;
                            break;
                        case "Italic":
                            selection.Font.Italic = selection.Font.Italic == 0 ? 1 : 0;
                            break;
                        case "Underline":
                            selection.Font.Underline = selection.Font.Underline == Word.WdUnderline.wdUnderlineNone
                                ? Word.WdUnderline.wdUnderlineSingle
                                : Word.WdUnderline.wdUnderlineNone;
                            break;
                        case "FontSize":
                            if (value is double size)
                                selection.Font.Size = (float)size;
                            break;
                        case "AlignLeft":
                            selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            break;
                        case "AlignCenter":
                            selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            break;
                        case "AlignRight":
                            selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                            break;
                    }
                }
                else if (_current == OfficeApp.Excel && _excelApp != null)
                {
                    var cell = _excelApp.ActiveCell;
                    if (cell == null) return;

                    switch (format)
                    {
                        case "Bold":
                            cell.Font.Bold = !cell.Font.Bold;
                            break;
                        case "Italic":
                            cell.Font.Italic = !cell.Font.Italic;
                            break;
                        case "Underline":
                            cell.Font.Underline = cell.Font.Underline == Excel.XlUnderlineStyle.xlUnderlineStyleNone
                                ? Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                                : Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                            break;
                        case "FontSize":
                            if (value is double size)
                                cell.Font.Size = size;
                            break;
                        case "AlignLeft":
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            break;
                        case "AlignCenter":
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            break;
                        case "AlignRight":
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler bei Formatierung '{format}': {ex.Message}");
                System.Windows.MessageBox.Show($"Fehler bei Formatierung: {ex.Message}", 
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        // LibreOffice-Integration
        private void LaunchLibreOffice(string args)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = GetLibreOfficePath(),
                    Arguments = args + " --nologo --norestore",
                    UseShellExecute = false
                };
                
                Process process = Process.Start(psi);
                if (process == null) return;
                
                _librePid = (uint)process.Id;
                
                // Warten, bis das Fenster erstellt wird
                try { process.WaitForInputIdle(5000); } catch { }
                Thread.Sleep(2500);  // Extra Wartezeit
                
                if (process.MainWindowHandle != IntPtr.Zero)
                {
                    Debug.WriteLine($"LibreOffice Hauptfenster gefunden: {process.MainWindowHandle.ToInt64():X}");
                    // LibreOffice-Fenster in Panel einbetten
                    EmbedOfficeWindow(process.MainWindowHandle);
                }
                else
                {
                    Debug.WriteLine("LibreOffice Hauptfenster nicht gefunden. Versuche es mit Polling...");
                    // Starte Timer für wiederholte Versuche
                    StartLibreOfficePolling();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Starten von LibreOffice: {ex.Message}");
                System.Windows.MessageBox.Show($"Fehler beim Starten von LibreOffice: {ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void StartLibreOfficePolling()
        {
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(1000);
            int attempts = 0;
            
            timer.Tick += (s, e) => 
            {
                attempts++;
                Debug.WriteLine($"LibreOffice Polling-Versuch {attempts}");
                
                Process[] processes = Process.GetProcessesByName("soffice");
                foreach (var process in processes)
                {
                    if (process.Id == _librePid && process.MainWindowHandle != IntPtr.Zero)
                    {
                        Debug.WriteLine($"LibreOffice Fenster gefunden: {process.MainWindowHandle.ToInt64():X}");
                        EmbedOfficeWindow(process.MainWindowHandle);
                        timer.Stop();
                        return;
                    }
                }
                
                // Nach 20 Sekunden aufgeben
                if (attempts > 20)
                {
                    Debug.WriteLine("LibreOffice Polling abgebrochen - kein Fenster gefunden");
                    timer.Stop();
                }
            };
            
            timer.Start();
        }
        
        private string GetLibreOfficePath()
        {
            string configPath = _settings?.LibreOfficeProgramPath;
            
            if (!string.IsNullOrWhiteSpace(configPath))
            {
                if (System.IO.File.Exists(configPath))
                    return configPath;
                    
                if (System.IO.Directory.Exists(configPath))
                {
                    string exePath = System.IO.Path.Combine(configPath.TrimEnd('\\'), "soffice.exe");
                    if (System.IO.File.Exists(exePath))
                        return exePath;
                }
            }
            
            return "soffice.exe"; // Annahme: im PATH
        }
        
        // HWND-Zugriff über Reflection
        private IntPtr GetWordHwnd()
        {
            if (_wordApp == null) return IntPtr.Zero;
            
            try
            {
                // Methode 1: ActiveWindow.Hwnd
                try
                {
                    dynamic activeWindow = _wordApp.ActiveWindow;
                    if (activeWindow != null)
                    {
                        int hwnd = activeWindow.Hwnd;
                        if (hwnd != 0) return new IntPtr(hwnd);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ActiveWindow.Hwnd Zugriff fehlgeschlagen: {ex.Message}");
                }
                
                // Methode 2: Application.Hwnd über Reflection
                try
                {
                    var hwndProperty = _wordApp.GetType().GetProperty("Hwnd");
                    if (hwndProperty != null)
                    {
                        var value = hwndProperty.GetValue(_wordApp);
                        if (value is int intValue && intValue != 0)
                            return new IntPtr(intValue);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Reflection Hwnd-Zugriff fehlgeschlagen: {ex.Message}");
                }
                
                // Methode 3: Suche nach dem Hauptfenster über Prozess
                try
                {
                    int wordProcessId = 0;
                    Type processIdType = _wordApp.GetType().GetProperty("ProcessID")?.PropertyType;
                    if (processIdType == typeof(int))
                        wordProcessId = (int)_wordApp.GetType().GetProperty("ProcessID").GetValue(_wordApp);
                    else if (_wordApp.GetType().GetProperty("ProcessID") != null)
                        wordProcessId = Convert.ToInt32(_wordApp.GetType().GetProperty("ProcessID").GetValue(_wordApp));
                    
                    if (wordProcessId > 0)
                    {
                        Process wordProcess = Process.GetProcessById(wordProcessId);
                        return wordProcess.MainWindowHandle;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Prozessbasierter Hwnd-Zugriff fehlgeschlagen: {ex.Message}");
                }
                
                Debug.WriteLine("Keine Methode zum Abrufen des Word-HWND erfolgreich!");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Allgemeiner Fehler beim Abrufen des Word-HWND: {ex.Message}");
            }
            
            return IntPtr.Zero;
        }
        
        private IntPtr GetExcelHwnd()
        {
            if (_excelApp == null) return IntPtr.Zero;
            
            try
            {
                // Methode 1: Application.Hwnd direkt
                try
                {
                    int hwnd = _excelApp.Hwnd;
                    if (hwnd != 0) return new IntPtr(hwnd);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Direkter Excel Hwnd-Zugriff fehlgeschlagen: {ex.Message}");
                }
                
                // Methode 2: Application.Hwnd über Reflection
                try
                {
                    var hwndProperty = _excelApp.GetType().GetProperty("Hwnd");
                    if (hwndProperty != null)
                    {
                        var value = hwndProperty.GetValue(_excelApp);
                        if (value is int intValue && intValue != 0)
                            return new IntPtr(intValue);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Reflection Excel Hwnd-Zugriff fehlgeschlagen: {ex.Message}");
                }
                
                // Methode 3: Suche nach dem Hauptfenster über Prozess
                try
                {
                    int excelProcessId = 0;
                    Type processIdType = _excelApp.GetType().GetProperty("ProcessID")?.PropertyType;
                    if (processIdType == typeof(int))
                        excelProcessId = (int)_excelApp.GetType().GetProperty("ProcessID").GetValue(_excelApp);
                    else if (_excelApp.GetType().GetProperty("ProcessID") != null)
                        excelProcessId = Convert.ToInt32(_excelApp.GetType().GetProperty("ProcessID").GetValue(_excelApp));
                    
                    if (excelProcessId > 0)
                    {
                        Process excelProcess = Process.GetProcessById(excelProcessId);
                        return excelProcess.MainWindowHandle;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Prozessbasierter Excel Hwnd-Zugriff fehlgeschlagen: {ex.Message}");
                }
                
                Debug.WriteLine("Keine Methode zum Abrufen des Excel-HWND erfolgreich!");
            }   
            catch (Exception ex)
            {
                Debug.WriteLine($"Allgemeiner Fehler beim Abrufen des Excel-HWND: {ex.Message}");
            }
            
            return IntPtr.Zero;
        }
        
        // Fenster-Einbettung / Positionierung
        private void EmbedOfficeWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || _panel == null || _panel.Handle == IntPtr.Zero)
            {
                Debug.WriteLine($"EmbedOfficeWindow fehlgeschlagen: HWND={hwnd}, Panel={_panel}, Handle={_panel?.Handle}");
                return;
            }
            
            try
            {
                Debug.WriteLine($"Versuche Fenster 0x{hwnd.ToInt64():X} in Panel 0x{_panel.Handle.ToInt64():X} einzubetten");
                
                // Stil des Office-Fensters ändern (Rahmen und Titelleiste entfernen)
                int style = GetWindowLong(hwnd, GWL_STYLE);
                Debug.WriteLine($"Ursprünglicher Fensterstil: 0x{style:X}");
                
                style &= ~WS_CAPTION;
                style &= ~WS_THICKFRAME;
                style |= WS_CHILD;
                
                Debug.WriteLine($"Neuer Fensterstil: 0x{style:X}");
                SetWindowLong(hwnd, GWL_STYLE, style);
                
                // Office-Fenster als Kind des Panels setzen
                Debug.WriteLine("Setze Parent-Fenster");
                bool parentResult = SetParent(hwnd, _panel.Handle);
                if (!parentResult)
                {
                    int error = Marshal.GetLastWin32Error();
                    Debug.WriteLine($"SetParent fehlgeschlagen mit Fehler: {error}");
                }
                
                // Maximiere das Fenster innerhalb des Panels
                ShowWindow(hwnd, SW_MAXIMIZE);
                
                // Größe und Position anpassen
                Debug.WriteLine("Passe Fenstergröße an");
                ResizeEmbeddedOffice();
                
                // Fokus setzen
                Debug.WriteLine("Setze Fensterfokus");
                SetForegroundWindow(hwnd);
                
                Debug.WriteLine("Fenstereinbettung abgeschlossen");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Einbetten des Office-Fensters: {ex.Message}");
            }
        }
        
        // Für Excel verwenden wir ein anderes Verfahren (keine Einbettung, nur Positionierung)
        // um das Ribbon-Problem zu vermeiden
        private void PositionExcelWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !IsWindow(hwnd))
            {
                Debug.WriteLine("Excel-Fenster ist nicht verfügbar für Positionierung");
                return;
            }
            
            try
            {
                // Position und Größe des Panels ermitteln
                var point = _panel.PointToScreen(new System.Drawing.Point(0, 0));
                var screenBounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                
                // Excel-Fenster positionieren
                MoveWindow(hwnd, point.X, point.Y, _panel.Width, _panel.Height, true);
                
                // Fokus setzen
                SetForegroundWindow(hwnd);
                
                // Timer für regelmäßige Neupositionierung starten
                StartExcelRepositionTimer(hwnd);
                
                Debug.WriteLine($"Excel-Fenster positioniert an {point.X},{point.Y} mit Größe {_panel.Width}x{_panel.Height}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Positionieren des Excel-Fensters: {ex.Message}");
            }
        }
        
        private void StartExcelRepositionTimer(IntPtr hwnd)
        {
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(500);
            
            timer.Tick += (s, e) =>
            {
                if (!IsWindow(hwnd) || _excelApp == null)
                {
                    timer.Stop();
                    return;
                }
                
                try
                {
                    // Neupositionieren bei Größenänderung
                    var point = _panel.PointToScreen(new System.Drawing.Point(0, 0));
                    MoveWindow(hwnd, point.X, point.Y, _panel.Width, _panel.Height, false);
                }
                catch
                {
                    timer.Stop();
                }
            };
            
            timer.Start();
        }
        
        private void ResizeEmbeddedOffice()
        {
            IntPtr hwnd = IntPtr.Zero;
            
            // Aktives Fenster bestimmen
            if (IsLibre())
            {
                Process[] processes = Process.GetProcessesByName("soffice");
                foreach (var process in processes)
                {
                    if (process.Id == _librePid && process.MainWindowHandle != IntPtr.Zero)
                    {
                        hwnd = process.MainWindowHandle;
                        break;
                    }
                }
            }
            else if (_current == OfficeApp.Word && _wordApp != null)
            {
                hwnd = GetWordHwnd();
            }
            else if (_current == OfficeApp.Excel && _excelApp != null)
            {
                hwnd = GetExcelHwnd();
                
                // Excel behandeln wir anders (siehe PositionExcelWindow)
                if (hwnd != IntPtr.Zero && IsWindow(hwnd))
                {
                    var point = _panel.PointToScreen(new System.Drawing.Point(0, 0));
                    MoveWindow(hwnd, point.X, point.Y, _panel.Width, _panel.Height, true);
                    return;
                }
            }
            
            // Größe anpassen für eingebettete Fenster
            if (hwnd != IntPtr.Zero && _panel != null && IsWindow(hwnd))
            {
                try
                {
                    int width = Math.Max(1, _panel.Width);
                    int height = Math.Max(1, _panel.Height);
                    
                    Debug.WriteLine($"Passe Fenstergröße an: {width}x{height}");
                    MoveWindow(hwnd, 0, 0, width, height, true);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Fehler beim Anpassen der Fenstergröße: {ex.Message}");
                }
            }
        }
        
        // Hilfsfunktionen
        private bool IsLibre() => string.Equals(_suite, "LibreOffice", StringComparison.OrdinalIgnoreCase);
        
        private void UpdateStatus()
        {
            if (_txtStatus == null || _txtInfo == null) return;
            
            string status, info;
            
            if (IsLibre())
            {
                status = "LibreOffice";
                info = "LibreOffice eingebettet";
            }
            else
            {
                status = $"MS Office: {(_current == OfficeApp.Word ? "Word" : "Excel")}";
                info = $"{(_current == OfficeApp.Word ? "Word" : "Excel")}-Dokument";
            }
            
            _txtStatus.Text = status;
            _txtInfo.Text = info;
        }
        
        private void EnsureOfficeCompletelyTerminated()
        {
            // Zuerst sauber beenden versuchen
            CleanupExistingOffice();
            
            // Zusätzlich kurz warten, um sicherzustellen, dass alle Prozesse beendet wurden
            Thread.Sleep(500);
            
            // Dann alle EXCEL.EXE und WINWORD.EXE Prozesse finden, die wir gestartet haben könnten
            try
            {
                if (_current == OfficeApp.Excel)
                {
                    foreach (var process in Process.GetProcessesByName("EXCEL"))
                    {
                        try
                        {
                            // Nur Prozesse beenden, die keine Dokumente geöffnet haben
                            if (process.MainWindowTitle == "Microsoft Excel" || 
                                string.IsNullOrEmpty(process.MainWindowTitle))
                            {
                                process.Kill();
                            }
                        }
                        catch { /* Ignorieren */ }
                    }
                }
                else
                {
                    foreach (var process in Process.GetProcessesByName("WINWORD"))
                    {
                        try
                        {
                            // Nur Prozesse beenden, die keine Dokumente geöffnet haben
                            if (process.MainWindowTitle == "Microsoft Word" || 
                                string.IsNullOrEmpty(process.MainWindowTitle))
                            {
                                process.Kill();
                            }
                        }
                        catch { /* Ignorieren */ }
                    }
                }
            }
            catch { /* Ignorieren */ }
            
            // Nochmals kurz warten
            Thread.Sleep(500);
        }
        
        private void CleanupExistingOffice()
        {
            // Word aufräumen
            if (_wordApp != null)
            {
                try
                {
                    _wordApp.Quit(false);
                    Marshal.ReleaseComObject(_wordApp);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Fehler beim Beenden von Word: {ex.Message}");
                }
                finally
                {
                    _wordApp = null;
                }
            }
            
            // Excel aufräumen
            if (_excelApp != null)
            {
                try
                {   
                    _excelApp.Quit();
                    Marshal.ReleaseComObject(_excelApp);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Fehler beim Beenden von Excel: {ex.Message}");
                }
                finally
                {
                    _excelApp = null;
                }
            }
        }
        
        public override void SetVisible(bool visible)
        {
            base.SetVisible(visible);
            
            if (!visible)
            {
                // Beim Verstecken aufräumen
                CleanupExistingOffice();
            }
        }
    }
}