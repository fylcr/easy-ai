Add-Type -AssemblyName System.Windows.Forms

function Select-InstallPath {
    $topForm = New-Object System.Windows.Forms.Form
    $topForm.TopMost = $true
    $topForm.ShowInTaskbar = $false
    $topForm.FormBorderStyle = 'None'
    $topForm.Size = New-Object System.Drawing.Size(1, 1)
    $topForm.StartPosition = 'CenterScreen'
    $topForm.Show()
    $topForm.Hide()

    while ($true) {
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "请选择安装路径"
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop
        $folderBrowser.ShowNewFolderButton = $true

        $result = $folderBrowser.ShowDialog($topForm)

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderBrowser.SelectedPath

            $confirm = [System.Windows.Forms.MessageBox]::Show(
                $topForm,
                "您选择的安装路径为：`n`n$selectedPath`n`n程序将安装到：`n$selectedPath\OmniVoice`n`n是否继续安装？",
                "确认安装路径",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                $topForm.Dispose()
                return $selectedPath
            }
        }
        else {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                $topForm,
                "您确定要退出安装吗？",
                "退出确认",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                $topForm.Dispose()
                Write-Host "用户已取消安装，程序退出。" -ForegroundColor Yellow
                exit
            }
        }
    }
}

function Install-OmniVoice {
    param (
        [string]$InstallPath
    )

    # ==================== 路径定义 ====================
    $downloadUrl    = "https://gh-proxy.org/https://github.com/HildaM/LongCat-AudioDiT-Web/archive/refs/heads/main.zip"
    $uvDownloadUrl  = "https://releases.astral.sh/github/uv/releases/download/0.11.3/uv-x86_64-pc-windows-msvc.zip"

    $tempZip        = Join-Path $env:TEMP "LongCat-AudioDiT-Web-main"
    $tempUvZip      = Join-Path $env:TEMP "uv-temp.zip"
    $tempUvDir      = Join-Path $env:TEMP "uv-temp"

    $oldName        = Join-Path $InstallPath "LongCat-AudioDiT-Web-main"
    $newName        = Join-Path $InstallPath "LongCat-AudioDiT"

    # ==================== 检查目标文件夹 ====================
    if (Test-Path $newName) {
        $topForm = New-Object System.Windows.Forms.Form
        $topForm.TopMost = $true
        $topForm.ShowInTaskbar = $false
        $topForm.FormBorderStyle = 'None'
        $topForm.Size = New-Object System.Drawing.Size(1, 1)
        $topForm.Show()
        $topForm.Hide()

        $overwrite = [System.Windows.Forms.MessageBox]::Show(
            $topForm,
            "目标文件夹已存在：`n`n$newName`n`n是否覆盖安装？",
            "文件夹已存在",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $topForm.Dispose()

        if ($overwrite -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "正在删除旧版本..." -ForegroundColor Yellow
            Remove-Item -Path $newName -Recurse -Force
        }
        else {
            Write-Host "安装已取消。" -ForegroundColor Yellow
            exit
        }
    }

    # ==================== [1/7] 下载 OmniVoice ====================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " [1/7] 正在下载 LongCat-AudioDiT 源码 ..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "下载地址: $downloadUrl" -ForegroundColor Gray
    Write-Host ""

    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $downloadUrl -OutFile $tempZip -UseBasicParsing
        Write-Host "✅ LongCat-AudioDiT 源码下载完成！" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ 下载失败: $_" -ForegroundColor Red
        if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
        Read-Host "按回车键退出"
        exit 1
    }

    # ==================== [2/7] 解压并重命名 ====================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " [2/7] 正在解压并安装 LongCat-AudioDiT ..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "安装目标: $newName" -ForegroundColor Gray
    Write-Host ""

    try {
        if (Test-Path $oldName) { Remove-Item -Path $oldName -Recurse -Force }

        Expand-Archive -Path $tempZip -DestinationPath $InstallPath -Force
        Rename-Item -Path $oldName -NewName "OmniVoice" -Force

        Remove-Item $tempZip -Force

        Write-Host "✅ OmniVoice 解压安装完成！" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ 解压/重命名失败: $_" -ForegroundColor Red
        if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
        Read-Host "按回车键退出"
        exit 1
    }

    # ==================== [3/7] 下载 uv 工具（仅临时使用） ====================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " [3/7] 正在下载 uv 包管理工具 ..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "下载地址: $uvDownloadUrl" -ForegroundColor Gray
    Write-Host ""

    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $uvDownloadUrl -OutFile $tempUvZip -UseBasicParsing

        if (Test-Path $tempUvDir) { Remove-Item -Path $tempUvDir -Recurse -Force }
        Expand-Archive -Path $tempUvZip -DestinationPath $tempUvDir -Force

        $uvExe = Get-ChildItem -Path $tempUvDir -Filter "uv.exe" -Recurse | Select-Object -First 1
        if (-not $uvExe) { throw "在解压目录中未找到 uv.exe" }
        $uvPath = $uvExe.FullName

        Remove-Item $tempUvZip -Force

        Write-Host "✅ uv 工具准备就绪 (临时)" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ uv 下载/解压失败: $_" -ForegroundColor Red
        if (Test-Path $tempUvZip) { Remove-Item $tempUvZip -Force }
        if (Test-Path $tempUvDir) { Remove-Item $tempUvDir -Recurse -Force }
        Read-Host "按回车键退出"
        exit 1
    }

    # ==================== [4/7] 安装 Python 依赖 ====================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " [4/7] 正在安装 Python 依赖 (uv sync) ..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "工作目录: $newName" -ForegroundColor Gray
    Write-Host "镜像源:   https://mirrors.aliyun.com/pypi/simple" -ForegroundColor Gray
    Write-Host ""

    try {
        Push-Location $newName

        & $uvPath sync --default-index "https://mirrors.aliyun.com/pypi/simple" --link-mode=copy

        if ($LASTEXITCODE -ne 0) { throw "uv sync 返回错误代码: $LASTEXITCODE" }

        Write-Host "✅ Python 依赖安装完成！" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ 依赖安装失败: $_" -ForegroundColor Red
        Pop-Location
        if (Test-Path $tempUvDir) { Remove-Item $tempUvDir -Recurse -Force }
        Read-Host "按回车键退出"
        exit 1
    }

    # ==================== [5/7] 下载模型文件（带重试） ====================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " [5/7] 正在下载 HuggingFace 模型 ..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "模型仓库: k2-fsa/OmniVoice" -ForegroundColor Gray
    Write-Host "镜像站:   https://hf-mirror.com" -ForegroundColor Gray
    Write-Host "缓存目录: .\hf_cache" -ForegroundColor Gray
    Write-Host ""

    $env:HF_HOME = ".\hf_cache"
    $env:HF_ENDPOINT = "https://hf-mirror.com"

    $pyScript = @"
import time
import sys
from huggingface_hub import snapshot_download

max_retries = 5

for attempt in range(1, max_retries + 1):
    try:
        print(f"\n>>> 下载尝试 {attempt}/{max_retries} ...")
        snapshot_download(
            repo_id="k2-fsa/OmniVoice",
            resume_download=True,
            max_workers=1,
        )
        print("\n>>> 模型下载成功！")
        sys.exit(0)
    except Exception as e:
        print(f"\n>>> 下载出错: {e}")
        if attempt < max_retries:
            wait = 5 * attempt
            print(f">>> {wait} 秒后重试 ...")
            time.sleep(wait)
        else:
            print(f"\n>>> 已重试 {max_retries} 次，全部失败。")
            sys.exit(1)
"@

    $pyScriptPath = Join-Path $newName "_download_model.py"
    $pyScript | Out-File -FilePath $pyScriptPath -Encoding UTF8

    try {
        & $uvPath run python $pyScriptPath

        if ($LASTEXITCODE -ne 0) { throw "模型下载失败，已重试多次" }

        Write-Host "✅ 模型下载完成！" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ 模型下载失败: $_" -ForegroundColor Red
        Pop-Location
        if (Test-Path $pyScriptPath) { Remove-Item $pyScriptPath -Force }
        if (Test-Path $tempUvDir) { Remove-Item $tempUvDir -Recurse -Force }
        Read-Host "按回车键退出"
        exit 1
    }

    if (Test-Path $pyScriptPath) { Remove-Item $pyScriptPath -Force }

    Pop-Location

    # ==================== [6/7] 创建启动脚本 ====================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " [6/7] 正在创建启动脚本 ..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $batContent = @"
@echo off
cd /d "%~dp0"
set HF_HOME=%~dp0hf_cache

echo ========================================
echo   OmniVoice 正在启动...
echo   HF_HOME: %HF_HOME%
echo   访问地址: http://0.0.0.0:8001
echo ========================================

call .venv\Scripts\activate.bat
omnivoice-demo --ip 0.0.0.0 --port 8001
pause
"@

    $batPath = Join-Path $newName "start.bat"
    $batContent | Out-File -FilePath $batPath -Encoding ASCII

    Write-Host "✅ 启动脚本已创建: $batPath" -ForegroundColor Green

    # ==================== [7/7] 清理临时文件 ====================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " [7/7] 正在清理临时文件 ..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    # uv 仅安装时使用，安装完全部删除
    if (Test-Path $tempUvDir)  { Remove-Item $tempUvDir -Recurse -Force }
    if (Test-Path $tempUvZip)  { Remove-Item $tempUvZip -Force }
    if (Test-Path $tempZip)    { Remove-Item $tempZip -Force }

    Write-Host "✅ 清理完成！" -ForegroundColor Green

    # ==================== 安装完成 ====================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  ✅ OmniVoice 安装成功！" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  安装路径:   $newName" -ForegroundColor Cyan
    Write-Host "  模型缓存:   $newName\hf_cache" -ForegroundColor Cyan
    Write-Host "  启动脚本:   $batPath" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  双击 start.bat 即可启动服务" -ForegroundColor Yellow
    Write-Host ""
}

# ==================== 主程序 ====================
Clear-Host
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "       OmniVoice 安装程序" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  本程序将执行以下步骤：" -ForegroundColor White
Write-Host "  [1/7] 下载 OmniVoice 源码" -ForegroundColor Gray
Write-Host "  [2/7] 解压并安装 OmniVoice" -ForegroundColor Gray
Write-Host "  [3/7] 下载 uv 包管理工具 (临时)" -ForegroundColor Gray
Write-Host "  [4/7] 安装 Python 依赖" -ForegroundColor Gray
Write-Host "  [5/7] 下载 HuggingFace 模型" -ForegroundColor Gray
Write-Host "  [6/7] 创建启动脚本" -ForegroundColor Gray
Write-Host "  [7/7] 清理临时文件" -ForegroundColor Gray
Write-Host ""

$installPath = Select-InstallPath
Install-OmniVoice -InstallPath $installPath

Read-Host "按回车键退出安装程序"
