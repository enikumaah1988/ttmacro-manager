<#
ttmacro-manager の exe ビルド + 配布用 zip 化スクリプト。

使い方（プロジェクトルートで PowerShell 実行）:
    .\build.ps1

事前に .venv が作成済み + pip install -e ".[dev]" 済みであること。
#>

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

# pyproject.toml から version を抽出
$Version = (Select-String 'version\s*=\s*"([^"]+)"' pyproject.toml |
    Select-Object -First 1).Matches.Groups[1].Value

Write-Host "ttmacro-manager v$Version を build します..." -ForegroundColor Cyan

# 1. PyInstaller で exe を生成（成果物は bin/*.exe）
Write-Host "[1/3] exe ビルド中..."
.venv\Scripts\pyinstaller.exe --clean --distpath bin packaging\ttmacro-launcher.spec
.venv\Scripts\pyinstaller.exe --clean --distpath bin packaging\ttmacro-generate.spec

# 2. deploy/ フォルダを組み立て（毎回クリーンに作り直す）
Write-Host "[2/3] deploy/ を組み立て中..."
$Deploy = "deploy"
if (Test-Path $Deploy) { Remove-Item $Deploy -Recurse -Force }
$null = New-Item -ItemType Directory -Path `
    "$Deploy\bin", "$Deploy\data", "$Deploy\macros\templates", `
    "$Deploy\keys", "$Deploy\logs"

Copy-Item bin\ttmacro-launcher.exe "$Deploy\bin\"
Copy-Item bin\ttmacro-generate.exe "$Deploy\bin\"
Copy-Item data\servers_template.xlsx "$Deploy\data\"
Copy-Item macros\templates\*.ttl "$Deploy\macros\templates\"

# zip に空ディレクトリを残すための占位ファイル
$null = New-Item -ItemType File -Path `
    "$Deploy\keys\.gitkeep", "$Deploy\logs\.gitkeep"

# 3. zip 化
Write-Host "[3/3] zip 化中..."
$Zip = "ttmacro-manager-v$Version.zip"
if (Test-Path $Zip) { Remove-Item $Zip -Force }
Compress-Archive -Path "$Deploy\*" -DestinationPath $Zip

Write-Host ""
Write-Host "完了: $Zip" -ForegroundColor Green
Write-Host "  中身は deploy/ にも残っています（手動で再 zip 化に使えます）" -ForegroundColor Gray
