using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Microsoft.Win32;

namespace ChromeProfileLauncher
{
    internal static class Program
    {
        private static string FindChromeExecutable()
        {
            string localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string progFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string progFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            string[] candidates = new string[]
            {
                @"C:\Program Files\Google\Chrome\Application\chrome.exe",
                @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                Path.Combine(progFiles, @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(progFilesX86, @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(localApp, @"Google\Chrome\Application\chrome.exe")
            };

            foreach (string path in candidates)
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    return path;
                }
            }

            string[] regKeys = new string[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
            };

            foreach (RegistryKey root in new RegistryKey[] { Registry.LocalMachine, Registry.CurrentUser })
            {
                foreach (string subKey in regKeys)
                {
                    try
                    {
                        using (RegistryKey key = root.OpenSubKey(subKey))
                        {
                            if (key != null)
                            {
                                object val = key.GetValue(null);
                                if (val != null)
                                {
                                    string regPath = val.ToString().Trim().Trim('"');
                                    if (File.Exists(regPath))
                                    {
                                        return regPath;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return null;
        }

        private static string SanitizeProfileName(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            string cleaned = raw.Trim().TrimStart('@').Trim();
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                cleaned = cleaned.Replace(c.ToString(), "");
            }
            return cleaned.Trim();
        }

        private static void CleanLockAndCrashFiles(string profileDir)
        {
            string[] locks = new string[] { "SingletonLock", "SingletonCookie", "SingletonSocket" };
            foreach (string lk in locks)
            {
                try
                {
                    string p = Path.Combine(profileDir, lk);
                    if (File.Exists(p))
                    {
                        File.Delete(p);
                    }
                }
                catch
                {
                }
            }

            string[] prefFiles = new string[]
            {
                Path.Combine(profileDir, "Default", "Preferences"),
                Path.Combine(profileDir, "Preferences")
            };
            foreach (string pref in prefFiles)
            {
                try
                {
                    if (File.Exists(pref))
                    {
                        string content = File.ReadAllText(pref, Encoding.UTF8);
                        if (content.Contains("\"exit_type\":\"Crashed\"") || content.Contains("\"exited_cleanly\":false"))
                        {
                            content = content.Replace("\"exit_type\":\"Crashed\"", "\"exit_type\":\"Normal\"");
                            content = content.Replace("\"exited_cleanly\":false", "\"exited_cleanly\":true");
                            File.WriteAllText(pref, content, Encoding.UTF8);
                        }
                    }
                }
                catch
                {
                }
            }
        }

        [STAThread]
        private static void Main(string[] args)
        {
            try {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;
                Console.InputEncoding = Encoding.UTF8;
            }
            catch
            {
            }

            Console.Title = "Chrome カスタムプロファイル起動ツール";
            string baseDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\', '/');
            string profilesRoot = Path.Combine(baseDir, "chrome_profiles");
            Directory.CreateDirectory(profilesRoot);

            string chromeExe = FindChromeExecutable();
            if (string.IsNullOrEmpty(chromeExe))
            {
                Console.WriteLine("=========================================");
                Console.WriteLine("[エラー] Google Chrome (chrome.exe) が見つかりませんでした。");
                Console.WriteLine("Google Chrome をインストールしてから再度お試しください。");
                Console.WriteLine("=========================================");
                Console.WriteLine("何かキーを押すと終了します...");
                Console.ReadKey();
                return;
            }

            while (true)
            {
                Console.WriteLine("=========================================");
                Console.WriteLine("Chrome カスタムプロファイル起動ツール");
                Console.Write("作成・使用するプロファイル名を入力してください: ");

                string inputName = (args != null && args.Length > 0) ? args[0] : Console.ReadLine();
                if (args != null && args.Length > 0) args = null;

                string profileName = SanitizeProfileName(inputName);
                if (string.IsNullOrEmpty(profileName))
                {
                    Console.WriteLine("※ 有効なプロファイル名を入力してください（例: wqwq1320）");
                    Console.WriteLine();
                    continue;
                }

                string profileDir = Path.Combine(profilesRoot, profileName);
                bool isNewProfile = !Directory.Exists(profileDir) || !Directory.Exists(Path.Combine(profileDir, "Default"));
                Directory.CreateDirectory(profileDir);

                CleanLockAndCrashFiles(profileDir);

                Console.WriteLine();
                if (isNewProfile)
                {
                    Console.WriteLine("[初回作成] 新規プロファイルを作成してGoogle Chromeを起動します...");
                }
                else
                {
                    Console.WriteLine("[既存利用] 保存済みプロファイルを読み込んでGoogle Chromeを起動します...");
                }
                Console.WriteLine("プロファイル保存先: " + profileDir);
                Console.WriteLine("※ 勝手に外部サイトへ遷移しません。ご自身でログインしたいサイトを開いて手動ログインしてください。");
                Console.WriteLine("※ ログイン後、Chromeウィンドウを閉じるとログイン情報がプロファイルに完全保存されます。");
                Console.WriteLine("=========================================");

                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo();
                    psi.FileName = chromeExe;
                    // URLは指定せず --user-data-dir のみを渡すことで、通常の「Chrome にログイン」/新規タブで起動
                    psi.Arguments = string.Format("--user-data-dir=\"{0}\"", profileDir);
                    psi.UseShellExecute = false;

                    using (Process proc = Process.Start(psi))
                    {
                        if (proc != null)
                        {
                            Console.WriteLine("Google Chrome を起動しました。Chromeを閉じると保存が完了します...");
                            proc.WaitForExit();
                        }
                    }

                    Console.WriteLine();
                    Console.WriteLine("✅ プロファイル [" + profileName + "] の情報を保持・保存しました！");
                    Console.WriteLine("2回目以降は同じプロファイル名、または自動投稿アプリからログイン済みのまま利用できます。");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("[エラー] Chromeの起動に失敗しました: " + ex.Message);
                }

                Console.WriteLine();
                Console.Write("続けて別のプロファイルを起動しますか？ (y=続行 / Enter=終了): ");
                string again = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(again) || !again.Trim().Equals("y", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
                Console.WriteLine();
            }
            }
            catch (Exception topEx)
            {
                Console.WriteLine("=========================================");
                Console.WriteLine("[エラー] 予期しないエラーが発生しました:");
                Console.WriteLine(topEx.ToString());
                Console.WriteLine("=========================================");
                Console.WriteLine("Enterキーを押すと終了します...");
                Console.ReadLine();
            }
        }
    }
}