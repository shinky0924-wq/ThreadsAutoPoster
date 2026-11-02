using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

namespace ThreadsAutoPoster
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string scriptPath = Path.Combine(baseDir, "threads_gui_app.py");

                // Multi-location search if not in baseDir
                if (!File.Exists(scriptPath))
                {
                    string trimmedBase = baseDir.TrimEnd('\\', '/');
                    string parentDir = Path.GetDirectoryName(trimmedBase);
                    string[] searchPaths = new string[] {
                        Path.Combine(baseDir, "ThreadsAutoPoster", "threads_gui_app.py"),
                        !string.IsNullOrEmpty(parentDir) ? Path.Combine(parentDir, "threads_gui_app.py") : "",
                        !string.IsNullOrEmpty(parentDir) ? Path.Combine(parentDir, "ThreadsAutoPoster", "threads_gui_app.py") : "",
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "threads_gui_app.py"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "threads_auto_poster_bundle", "threads_gui_app.py")
                    };
                    foreach (string sp in searchPaths)
                    {
                        if (!string.IsNullOrEmpty(sp) && File.Exists(sp))
                        {
                            scriptPath = sp;
                            baseDir = Path.GetDirectoryName(sp);
                            break;
                        }
                    }
                }

                // Candidates for Python executable
                string[] pyCandidates = new string[] {
                    "python.exe",
                    "py.exe",
                    "pythonw.exe",
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python312\python.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python313\python.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python311\python.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python310\python.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python39\python.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python38\python.exe"),
                    @"C:\Program Files\Python312\python.exe",
                    @"C:\Program Files\Python313\python.exe",
                    @"C:\Program Files\Python311\python.exe",
                    @"C:\Program Files\Python310\python.exe",
                    @"C:\Program Files\Python39\python.exe",
                    @"C:\Python312\python.exe",
                    @"C:\Python313\python.exe",
                    @"C:\Python311\python.exe",
                    @"C:\Python310\python.exe",
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python312\pythonw.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python313\pythonw.exe"),
                    @"C:\Program Files\Python312\pythonw.exe",
                    @"C:\Program Files\Python313\pythonw.exe"
                };

                string foundPy = null;
                foreach (string candidate in pyCandidates)
                {
                    try
                    {
                        ProcessStartInfo testInfo = new ProcessStartInfo(candidate, "-c \"import sys\"")
                        {
                            CreateNoWindow = true,
                            UseShellExecute = false
                        };
                        using (Process proc = Process.Start(testInfo))
                        {
                            proc.WaitForExit(2000);
                            if (proc.ExitCode == 0)
                            {
                                foundPy = candidate;
                                break;
                            }
                        }
                    }
                    catch { }
                }

                if (string.IsNullOrEmpty(foundPy))
                {
                    DialogResult res = MessageBox.Show(
                        "Threads AutoPosterの実行に必要なPythonが検出されませんでした。\n\n「0_他のPCでもすぐ動く環境作成.bat」を実行して自動セットアップを行いますか？",
                        "Threads AutoPoster - 初期設定",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );
                    if (res == DialogResult.Yes)
                    {
                        string setupBat = Path.Combine(baseDir, "0_他のPCでもすぐ動く環境作成.bat");
                        if (!File.Exists(setupBat)) setupBat = Path.Combine(baseDir, "0_Python環境セットアップ.bat");
                        if (File.Exists(setupBat))
                        {
                            Process.Start(new ProcessStartInfo(setupBat) { WorkingDirectory = baseDir, UseShellExecute = true });
                        }
                        else
                        {
                            Process.Start("https://www.python.org/downloads/");
                        }
                    }
                    return;
                }

                if (!File.Exists(scriptPath))
                {
                    bool isTempZip = baseDir.IndexOf("Temp", StringComparison.OrdinalIgnoreCase) >= 0;
                    string notFoundMsg = "";
                    if (isTempZip)
                    {
                        notFoundMsg = "【ZIPファイルの「すべて展開（解凍）」が必要です】\n\n"
                            + "現在、ダウンロードしたZIPファイルの中（圧縮フォルダー）から直接実行されています。\n"
                            + "Windowsの仕様上、ZIPの中から直接実行すると他の必須ファイルが読み込めません。\n\n"
                            + "【解決手順（30秒）】\n"
                            + "1. ダウンロードフォルダを開きます\n"
                            + "2. 「threads_auto_poster_bundle.zip」を右クリックします\n"
                            + "3. 「すべて展開」を選択して解凍してください\n"
                            + "4. 展開された新しいフォルダ内にある「ThreadsAutoPoster.vbs」または「start_gui.bat」を開いてください。";
                    }
                    else
                    {
                        notFoundMsg = "threads_gui_app.py が見つかりません。\n\n"
                            + "【原因】\n"
                            + "ThreadsAutoPoster.exe だけを単体でデスクトップ等に移動した可能性があります。\n\n"
                            + "【解決手順】\n"
                            + "解凍されたフォルダ一式（Excelやスクリプトが全て入っているフォルダ）の中で実行するか、\n"
                            + "フォルダ内の「ThreadsAutoPoster.vbs」または「start_gui.bat」から起動してください。";
                    }
                    MessageBox.Show(notFoundMsg, "ファイル未検出エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Check required packages
                bool hasPkg = false;
                try
                {
                    ProcessStartInfo testPkg = new ProcessStartInfo(foundPy, "-c \"import openpyxl\"")
                    {
                        CreateNoWindow = true,
                        UseShellExecute = false
                    };
                    using (Process p = Process.Start(testPkg))
                    {
                        p.WaitForExit(2500);
                        if (p.ExitCode == 0) hasPkg = true;
                    }
                }
                catch { }

                if (!hasPkg)
                {
                    ProcessStartInfo pipInfo = new ProcessStartInfo(foundPy, "-m pip install openpyxl playwright")
                    {
                        WorkingDirectory = baseDir,
                        UseShellExecute = true
                    };
                    try
                    {
                        Process pipProc = Process.Start(pipInfo);
                        pipProc.WaitForExit();
                    }
                    catch { }
                }

                // Launch GUI App
                ProcessStartInfo runInfo = new ProcessStartInfo(foundPy, "\"" + scriptPath + "\"")
                {
                    WorkingDirectory = baseDir,
                    UseShellExecute = false
                };
                Process appProc = Process.Start(runInfo);

                // Wait briefly to detect immediate startup crash
                if (appProc.WaitForExit(1500))
                {
                    if (appProc.ExitCode != 0)
                    {
                        string crashLog = Path.Combine(baseDir, "crash_log.txt");
                        string details = "";
                        if (File.Exists(crashLog))
                        {
                            try { details = File.ReadAllText(crashLog); } catch { }
                        }
                        MessageBox.Show(
                            "GUIアプリの起動時にエラーが発生しました。\n\n終了コード: " + appProc.ExitCode + "\n\n" + details + "\n\n※同梱の「start_gui.bat」をダブルクリックして起動すると、詳細なログを確認できます。",
                            "Threads AutoPoster - 起動エラー",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("起動中にエラーが発生しました:\n\n" + ex.Message + "\n\n※「start_gui.bat」から起動してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
