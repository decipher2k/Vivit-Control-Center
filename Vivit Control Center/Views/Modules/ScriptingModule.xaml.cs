using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class ScriptingModule : BaseSimpleModule
    {
        private enum ScriptLanguage { Unknown, CSharp, Python }

        private string _currentFilePath;
        private ScriptLanguage _currentLanguage = ScriptLanguage.Unknown;
        private readonly List<string> _csharpKeywords = new List<string>
        {
            // Keywords
            "abstract","as","base","bool","break","byte","case","catch","char","checked","class","const","continue","decimal","default","delegate","do","double","else","enum","event","explicit","extern","false","finally","fixed","float","for","foreach","goto","if","implicit","in","int","interface","internal","is","lock","long","new","null","object","operator","out","override","params","private","protected","public","readonly","ref","return","sbyte","sealed","short","sizeof","stackalloc","static","string","struct","switch","this","throw","true","try","typeof","uint","ulong","unchecked","unsafe","ushort","using","virtual","void","volatile","while",
            // Common types/members
            "Console","Console.WriteLine","Math","DateTime","Guid","List","Dictionary","Task","File","Directory","Path","Environment"
        };
        private readonly List<string> _pythonKeywords = new List<string>
        {
            "and","as","assert","break","class","continue","def","del","elif","else","except","False","finally","for","from","global","if","import","in","is","lambda","None","nonlocal","not","or","pass","raise","return","True","try","while","with","yield",
            // Common modules/types
            "print","math","datetime","os","sys","json","list","dict","str","len","range"
        };

        // Default C# script skeleton
        private const string DefaultCSharpSkeleton =
            "using System;\n" +
            "using System.IO;\n" +
            "using System.Linq;\n" +
            "using System.Collections.Generic;\n" +
            "using System.Net;\n" +
            "using System.Net.Http;\n" +
            "using System.Text;\n" +
            "using System.Text.RegularExpressions;\n\n" +
            "public static class Script\n" +
            "{\n" +
            "    public static void Main()\n" +
            "    {\n" +
            "        Console.WriteLine(\"Hello from C# script.\");\n" +
            "    }\n" +
            "}\n";

        // Access named controls via FindName
        private TextBox ScriptTextBox => (TextBox)FindName("txtScript");
        private TextBlock DetectedLanguageTextBlock => (TextBlock)FindName("txtDetectedLanguage");
        private System.Windows.Controls.Primitives.Popup CompletionPopup => (System.Windows.Controls.Primitives.Popup)FindName("completionPopup");
        private ListBox CompletionList => (ListBox)FindName("completionList");
        private TextBox OutputTextBox => (TextBox)FindName("txtOutput");

        public ScriptingModule()
        {
            InitializeComponent();
            ClearOutput();
            WriteOutputLine("Scripting bereit. Sprache wird automatisch erkannt.");
            // Insert default C# skeleton if editor is empty
            var tb = ScriptTextBox;
            if (tb != null && string.IsNullOrWhiteSpace(tb.Text))
            {
                tb.Text = DefaultCSharpSkeleton;
                tb.CaretIndex = tb.Text.Length;
            }
            UpdateLanguage();
            if (tb != null) tb.Focus();
        }

        // Auto language detection based on simple heuristics
        private ScriptLanguage DetectLanguage(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return ScriptLanguage.Unknown;
            // Heuristics for Python
            if (text.Contains("def ") || text.Contains("import ") || text.Contains("print(") || text.Contains("#!/usr/bin/env python"))
                return ScriptLanguage.Python;
            // Heuristics for C#
            if (text.Contains("using ") || text.Contains("namespace ") || text.Contains("class ") || text.Contains("Console.WriteLine"))
                return ScriptLanguage.CSharp;
            // Fallback: braces and semicolons suggest C#
            if (text.Contains(";") || text.Contains("{") || text.Contains("}")) return ScriptLanguage.CSharp;
            // Colon and no braces often Python
            if (text.Contains(":") && !text.Contains("{")) return ScriptLanguage.Python;
            return ScriptLanguage.Unknown;
        }

        private void UpdateLanguage()
        {
            var tb = ScriptTextBox; var lbl = DetectedLanguageTextBlock;
            var text = tb != null ? tb.Text : string.Empty;
            _currentLanguage = DetectLanguage(text);
            if (lbl != null) lbl.Text = _currentLanguage == ScriptLanguage.Unknown ? "Auto" : _currentLanguage.ToString();
        }

        private void ClearOutput()
        {
            var ob = OutputTextBox; if (ob != null) ob.Text = string.Empty;
        }

        private void WriteOutput(string text)
        {
            var ob = OutputTextBox; if (ob == null) return;
            ob.AppendText(text);
            ob.ScrollToEnd();
        }

        private void WriteOutputLine(string text)
        {
            WriteOutput(text + Environment.NewLine);
        }

        private async void btnRun_Click(object sender, RoutedEventArgs e)
        {
            await RunCurrentAsync();
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Scripts (*.cs;*.py)|*.cs;*.py|C# (*.cs)|*.cs|Python (*.py)|*.py|Alle Dateien (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                _currentFilePath = dlg.FileName;
                var tb = ScriptTextBox; if (tb != null) tb.Text = File.ReadAllText(_currentFilePath);
                UpdateLanguage();
                WriteOutputLine($"Geoeffnet: {_currentFilePath}");
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath)) { btnSaveAs_Click(sender, e); return; }
            var tb = ScriptTextBox; if (tb != null)
            {
                File.WriteAllText(_currentFilePath, tb.Text, Encoding.UTF8);
                WriteOutputLine($"Gespeichert: {_currentFilePath}");
            }
        }

        private void btnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "C# (*.cs)|*.cs|Python (*.py)|*.py|Alle Dateien (*.*)|*.*",
                FileName = _currentLanguage == ScriptLanguage.Python ? "script.py" : "script.cs"
            };
            if (dlg.ShowDialog() == true)
            {
                _currentFilePath = dlg.FileName;
                var tb = ScriptTextBox; if (tb != null)
                {
                    File.WriteAllText(_currentFilePath, tb.Text, Encoding.UTF8);
                    WriteOutputLine($"Gespeichert: {_currentFilePath}");
                }
            }
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            var tb = ScriptTextBox; if (tb != null) tb.Clear();
            _currentFilePath = null;
            UpdateLanguage();
            ClearOutput();
            WriteOutputLine("Editor geleert.");
        }

        private async Task RunCurrentAsync()
        {
            UpdateLanguage();
            var tb = ScriptTextBox; var code = tb != null ? tb.Text : string.Empty;
            ClearOutput();
            WriteOutputLine($"Sprache: {_currentLanguage}");
            switch (_currentLanguage)
            {
                case ScriptLanguage.CSharp:
                    await RunCSharpAsync(code);
                    break;
                case ScriptLanguage.Python:
                    await RunPythonAsync(code);
                    break;
                default:
                    WriteOutputLine("Sprache konnte nicht erkannt werden.");
                    break;
            }
        }

        // Simple in-memory compile and reflect invoke
        private Task RunCSharpAsync(string code)
        {
            return Task.Run(() =>
            {
                try
                {
                    var provider = CodeDomProvider.CreateProvider("CSharp");
                    var parameters = new CompilerParameters
                    {
                        GenerateInMemory = true,
                        GenerateExecutable = false,
                        TreatWarningsAsErrors = false
                    };

                    // Reference common assemblies
                    var refs = new[]
                    {
                        "mscorlib.dll","System.dll","System.Core.dll","Microsoft.CSharp.dll",
                        "System.Data.dll","System.Xml.dll","System.Net.Http.dll","System.Xml.Linq.dll",
                        "System.Data.DataSetExtensions.dll","System.Drawing.dll","System.Windows.Forms.dll"
                    };
                    foreach (var r in refs) parameters.ReferencedAssemblies.Add(r);
                    parameters.ReferencedAssemblies.Add(Assembly.GetExecutingAssembly().Location);

                    // If code is likely a snippet, wrap into a Program class with Main and common usings
                    if (!code.Contains("class "))
                    {
                        var usings = string.Join("\n", new[]
                        {
                            "using System;","using System.IO;","using System.Linq;","using System.Collections.Generic;",
                            "using System.Net;","using System.Net.Http;","using System.Text;","using System.Text.RegularExpressions;"
                        });
                        code = usings + "\npublic static class __Script { public static void Main() { " + code + " } }";
                    }

                    var res = provider.CompileAssemblyFromSource(parameters, new string[] { code });
                    if (res == null || res.Errors.HasErrors)
                    {
                        var sb = new StringBuilder();
                        if (res != null)
                        {
                            foreach (CompilerError err in res.Errors)
                                sb.AppendLine($"[{err.ErrorNumber}] Zeile {err.Line}: {err.ErrorText}");
                        }
                        else
                        {
                            sb.AppendLine("Unbekannter Compilerfehler.");
                        }
                        Dispatcher.Invoke(() => WriteOutputLine(sb.ToString()));
                        return;
                    }

                    var asm = res.CompiledAssembly;

                    // Try to find an entry point to invoke
                    MethodInfo entry = asm.EntryPoint; // if an exe; for library we search known patterns
                    if (entry == null)
                    {
                        var type = asm.GetTypes().FirstOrDefault(t => t.GetMethod("Main", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static) != null)
                                ?? asm.GetTypes().FirstOrDefault(t => t.GetMethod("Run", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static) != null);
                        entry = type?.GetMethod("Main", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                             ?? type?.GetMethod("Run", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);
                    }

                    if (entry == null)
                    {
                        Dispatcher.Invoke(() => WriteOutputLine("Kein Einstiegspunkt gefunden (Main/Run)."));
                        return;
                    }

                    // Invoke static entry point and capture console output
                    var parametersInfo = entry.GetParameters();
                    object result = null;
                    var originalOut = Console.Out;
                    try
                    {
                        using (var sw = new StringWriter())
                        {
                            Console.SetOut(sw);
                            if (parametersInfo.Length == 0)
                                result = entry.Invoke(null, null);
                            else if (parametersInfo.Length == 1 && parametersInfo[0].ParameterType == typeof(string[]))
                                result = entry.Invoke(null, new object[] { new string[0] });
                            else
                                result = entry.Invoke(null, new object[parametersInfo.Length]);

                            var consoleText = sw.ToString();
                            if (!string.IsNullOrEmpty(consoleText))
                                Dispatcher.Invoke(() => WriteOutput(consoleText));
                        }
                    }
                    finally
                    {
                        Console.SetOut(originalOut);
                    }

                    if (result != null)
                    {
                        Dispatcher.Invoke(() => WriteOutputLine("Result: " + Convert.ToString(result)));
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => WriteOutputLine("C# Ausnahme: " + ex));
                }
            });
        }

        private Task RunPythonAsync(string code)
        {
            return Task.Run(() =>
            {
                try
                {
                    // Write to a temp .py and execute with default python
                    var temp = Path.Combine(Path.GetTempPath(), "occ_script_" + Guid.NewGuid().ToString("N") + ".py");
                    File.WriteAllText(temp, code, new UTF8Encoding(false));

                    var psi = new ProcessStartInfo
                    {
                        FileName = "python",
                        Arguments = "\"" + temp + "\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    };

                    var p = Process.Start(psi);
                    var stdout = p.StandardOutput.ReadToEnd();
                    var stderr = p.StandardError.ReadToEnd();
                    p.WaitForExit();

                    Dispatcher.Invoke(() =>
                    {
                        if (!string.IsNullOrEmpty(stdout)) WriteOutput(stdout);
                        if (!string.IsNullOrEmpty(stderr)) WriteOutputLine("[stderr]\n" + stderr);
                        WriteOutputLine($"ExitCode: {p.ExitCode}");
                    });

                    try { File.Delete(temp); } catch { }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => WriteOutputLine("Python Ausnahme: " + ex));
                }
            });
        }

        private void txtScript_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateLanguage();
        }

        private void txtScript_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F5 || (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control))
            {
                e.Handled = true;
                _ = RunCurrentAsync();
                return;
            }

            if (e.Key == Key.Space && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
                ShowCompletion();
                return;
            }
        }

        private void txtScript_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (e.Key == Key.S)
                {
                    e.Handled = true;
                    btnSave_Click(sender, e);
                }
                else if (e.Key == Key.O)
                {
                    e.Handled = true;
                    btnOpen_Click(sender, e);
                }
            }
        }

        // Basic completion: shows keywords per language
        private void ShowCompletion()
        {
            var tb = ScriptTextBox; if (tb == null) return;
            var caretIndex = tb.CaretIndex;
            var text = tb.Text;
            var start = caretIndex - 1;
            while (start >= 0 && (char.IsLetterOrDigit(text[start]) || text[start] == '_')) start--;
            start++;
            var prefix = caretIndex > start ? text.Substring(start, caretIndex - start) : string.Empty;

            IEnumerable<string> items = Enumerable.Empty<string>();
            if (_currentLanguage == ScriptLanguage.CSharp)
            {
                items = _csharpKeywords;
            }
            else if (_currentLanguage == ScriptLanguage.Python)
            {
                items = _pythonKeywords;
            }

            if (!string.IsNullOrEmpty(prefix))
                items = items.Where(i => i.StartsWith(prefix, StringComparison.InvariantCultureIgnoreCase));

            var list = items.Distinct().OrderBy(i => i).ToList();
            var popup = CompletionPopup; var lb = CompletionList;
            if (list.Count == 0 || popup == null || lb == null)
            {
                if (popup != null) popup.IsOpen = false;
                return;
            }

            lb.ItemsSource = list;
            lb.SelectedIndex = 0;

            // Compute caret point for popup placement
            var rect = tb.GetRectFromCharacterIndex(caretIndex, true);
            popup.HorizontalOffset = rect.X + 10;
            popup.VerticalOffset = rect.Y + rect.Height + 4;
            popup.IsOpen = true;
        }

        private void completionList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CommitSelectedCompletion();
        }

        private void completionList_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var popup = CompletionPopup;
            if (e.Key == Key.Enter || e.Key == Key.Tab)
            {
                e.Handled = true;
                CommitSelectedCompletion();
            }
            else if (e.Key == Key.Escape)
            {
                if (popup != null) popup.IsOpen = false;
            }
        }

        private void CommitSelectedCompletion()
        {
            var popup = CompletionPopup; var lb = CompletionList; var tb = ScriptTextBox;
            if (popup == null || lb == null || tb == null) return;
            if (!popup.IsOpen) return;
            var selected = lb.SelectedItem as string;
            if (string.IsNullOrEmpty(selected)) { popup.IsOpen = false; return; }

            var caretIndex = tb.CaretIndex;
            var text = tb.Text;
            var start = caretIndex - 1;
            while (start >= 0 && (char.IsLetterOrDigit(text[start]) || text[start] == '_')) start--;
            start++;

            tb.Text = text.Substring(0, start) + selected + text.Substring(caretIndex);
            tb.CaretIndex = start + selected.Length;
            popup.IsOpen = false;
            tb.Focus();
        }
    }
}