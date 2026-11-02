
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Threads AutoPoster - 原因診断＆自動修復ツール"

Clear-Host
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  Threads AutoPoster - 原因診断＆自動修復ツール" -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "このツールは、起動エラーの原因を自動特定し、1クリックで修復します。" -ForegroundColor Gray
Write-Host ""

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $baseDir) { $baseDir = Get-Location }
Set-Location $baseDir

$reportFile = Join-Path $baseDir "診断レポート.txt"
$nowStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
"Threads AutoPoster 診断レポート ($nowStr)" | Out-File -FilePath $reportFile -Encoding utf8

# [1/5] ファイル存在チェック
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "[1/5] 必須ファイルの存在チェック" -ForegroundColor White
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray

$files = @("threads_gui_app.py", "ThreadsAutoPoster.exe", "アカウント情報.xlsx")
foreach ($f in $files) {
    if (Test-Path (Join-Path $baseDir $f)) {
        Write-Host "  [OK] $f : 存在します" -ForegroundColor Green
        "  [OK] $f : 存在します" | Out-File -FilePath $reportFile -Append -Encoding utf8
    } else {
        Write-Host "  [NG] $f : 見つかりません" -ForegroundColor Red
        "  [NG] $f : 見つかりません" | Out-File -FilePath $reportFile -Append -Encoding utf8
    }
}
Write-Host ""

# [2/5] ログ確認
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "[2/5] 起動ログ (startup_log.txt) の確認" -ForegroundColor White
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray

$logPath = Join-Path $baseDir "startup_log.txt"
if (Test-Path $logPath) {
    Write-Host "  [発見] startup_log.txt の直近ログ:" -ForegroundColor Cyan
    Get-Content $logPath -Tail 15 | ForEach-Object { Write-Host ("    " + $_) -ForegroundColor Gray }
    "" | Out-File -FilePath $reportFile -Append -Encoding utf8
    "--- startup_log.txt ---" | Out-File -FilePath $reportFile -Append -Encoding utf8
    Get-Content $logPath | Out-File -FilePath $reportFile -Append -Encoding utf8
} else {
    Write-Host "  [未生成] まだ startup_log.txt は生成されていません。" -ForegroundColor Yellow
}

$crashPath = Join-Path $baseDir "crash_log.txt"
if (Test-Path $crashPath) {
    Write-Host "  [警告] crash_log.txt が見つかりました:" -ForegroundColor Red
    Get-Content $crashPath -Tail 10 | ForEach-Object { Write-Host ("    " + $_) -ForegroundColor Magenta }
}
Write-Host ""

# [3/5] Python環境の徹底診断
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "[3/5] Python環境の徹底診断" -ForegroundColor White
Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray

$validPy = $null
$hasTk = $false

$candidates = @()
$localApp = $env:LOCALAPPDATA
if ($localApp) {
    foreach ($v in @("313", "312", "311", "310", "39", "38")) {
        $candidates += "$localApp\Programs\Python\Python$v\python.exe"
    }
}
$prog = $env:ProgramFiles
if ($prog) {
    foreach ($v in @("313", "312", "311", "310", "39", "38")) {
        $candidates += "$prog\Python$v\python.exe"
    }
}
$candidates += "C:\Python312\python.exe"
$candidates += "C:\Python311\python.exe"
$candidates += "python.exe"
$candidates += "py.exe"

foreach ($cand in $candidates) {
    try {
        $testRes = & $cand -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $testRes -notmatch "WindowsApps") {
            $pyPath = $testRes.Trim()
            $tkTest = & $cand -c "import tkinter" 2>$null
            if ($LASTEXITCODE -eq 0) {
                $validPy = $cand
                $hasTk = $true
                Write-Host "  [OK] 有効なPython（GUI対応）を検出: $pyPath" -ForegroundColor Green
                break
            } elseif (-not $validPy) {
                $validPy = $cand
                $hasTk = $false
            }
        }
    } catch { }
}

if (-not $validPy) {
    Write-Host "  [NG] 有効なPythonがこのPCにインストールされていません！" -ForegroundColor Red
    Write-Host "       ※これが ThreadsAutoPoster.exe が起動しない原因です。" -ForegroundColor Yellow
} else {
    if ($hasTk) {
        Write-Host "  [OK] Tkinter (GUI描画): 正常動作" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Tkinter (GUI描画): 未検出 (公式インストーラーでのセットアップを推奨)" -ForegroundColor Yellow
    }

    # [4/5] ライブラリテスト
    Write-Host ""
    Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "[4/5] 必須ライブラリの動作テスト" -ForegroundColor White
    Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray

    $openpyxlTest = & $validPy -c "import openpyxl" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  [OK] Excel読み込み (openpyxl): 正常動作" -ForegroundColor Green
    } else {
        Write-Host "  [NG] Excel読み込み (openpyxl): 未インストール (自動修復可能)" -ForegroundColor Yellow
    }

    $playwrightTest = & $validPy -c "import playwright" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  [OK] ブラウザ自動化 (playwright): 正常動作" -ForegroundColor Green
    } else {
        Write-Host "  [NG] ブラウザ自動化 (playwright): 未インストール (自動修復可能)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "[5/5] 診断完了！次の操作を選択してください" -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  [1] 【おすすめ】ワンクリック自動修復" -ForegroundColor Green
Write-Host "      → Python・全ライブラリ・ブラウザを全自動導入して準備完了にします" -ForegroundColor White
Write-Host "  [2] GUIアプリ（threads_gui_app.py）をテスト起動してみる" -ForegroundColor White
Write-Host "  [3] startup_log.txt をメモ帳で開く" -ForegroundColor White
Write-Host "  [4] 診断レポート（診断レポート.txt）をメモ帳で開く" -ForegroundColor White
Write-Host "  [5] 終了する" -ForegroundColor Gray
Write-Host "========================================================" -ForegroundColor Cyan

$choice = Read-Host "番号を入力してください [1, 2, 3, 4, 5]"

switch ($choice) {
    "1" {
        Write-Host ""
        Write-Host "【自動修復を開始します】" -ForegroundColor Cyan
        if (-not $validPy -or -not $hasTk) {
            Write-Host "Pythonのインストール（winget）を試行中..." -ForegroundColor Yellow
            winget install --id Python.Python.3.12 --exact --accept-package-agreements --accept-source-agreements
            Start-Sleep -Seconds 3
            $validPy = "python.exe"
        }
        Write-Host "ライブラリ (openpyxl, playwright) をインストール中 (--user)..." -ForegroundColor Yellow
        & $validPy -m pip install --upgrade pip 2>$null
        & $validPy -m pip install --user openpyxl playwright
        Write-Host "Playwright Chromium ブラウザをインストール中..." -ForegroundColor Yellow
        & $validPy -m playwright install chromium
        Write-Host ""
        Write-Host "[OK] 全自動修復が完了しました！ThreadsAutoPoster.exe をダブルクリックして起動できます。" -ForegroundColor Green
    }
    "2" {
        if ($validPy) {
            Write-Host "GUIアプリをテスト起動中..." -ForegroundColor Cyan
            & $validPy threads_gui_app.py
        } else {
            Write-Host "Pythonが検出されていないため起動できません。[1]で自動修復してください。" -ForegroundColor Red
        }
    }
    "3" {
        if (Test-Path $logPath) { notepad.exe $logPath }
        else { Write-Host "startup_log.txt はまだ存在しません。" -ForegroundColor Yellow }
    }
    "4" {
        notepad.exe $reportFile
    }
}

Write-Host ""
Write-Host "キーボードの任意のキーを押すと画面を閉じます..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
