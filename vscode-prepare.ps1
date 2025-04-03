<#
.SYNOPSIS
Visual Studio Code 自动化测试环境一键配置脚本

.NOTES
需要以管理员权限运行，确保网络连接正常
#>

# 强制以管理员权限运行
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Install-VSCode(){
    # 下载 VSCode 安装程序
    $vscodeUrl = "https://update.code.visualstudio.com/latest/win32-x64/stable"
    $installer = "$env:TEMP\VSCodeSetup.exe"
    Write-Host "正在下载 Visual Studio Code..." -ForegroundColor Green
    Invoke-WebRequest -Uri $vscodeUrl -OutFile $installer

    # 静默安装 VSCode
    Write-Host "正在安装 Visual Studio Code..." -ForegroundColor Green
    Start-Process -FilePath $installer -ArgumentList "/VERYSILENT /NORESTART /MERGETASKS=!runcode,desktopicon" -Wait

    # 检查 VSCode 是否安装成功
    $vscodeExe = "${env:ProgramFiles}\Microsoft VS Code\Code.exe"
    if (-not (Test-Path $vscodeExe)) {
        Write-Host "Visual Studio Code 安装失败！" -ForegroundColor Red
        exit 1
    }
    Write-Host "Visual Studio Code 安装完成。" -ForegroundColor Green

    # 定义要安装的扩展列表
    $extensions = @(
        "esbenp.prettier-vscode",
        "MS-CEINTL.vscode-language-pack-zh-hans",
        "ms-python.python"
    )

    Write-Host "准备安装一些通用扩展..." -ForegroundColor Green
    Start-Sleep 3
    foreach ($extension in $extensions) {
        Write-Host "正在安装扩展：$extension..." -ForegroundColor Green

        $codePath = "${env:ProgramFiles}\Microsoft VS Code\bin\code.cmd"
        Start-Process -FilePath $codePath -ArgumentList `
            "--install-extension", $extension, `
            "--extensions-dir", "$env:USERPROFILE\.vscode\extensions", `
            "--user-data-dir", "$env:APPDATA\Code", `
            "--force" -NoNewWindow -Wait
    }

    Write-Host "所有扩展已安装完成。" -ForegroundColor Green
    Remove-Item $installer -Force -ErrorAction SilentlyContinue
    Write-Host "Visual Studio Code 已成功安装并配置完成！" -ForegroundColor Green
}

# 检查 VSCode 是否已安装
function Check-VSCodeInstalled {
    $vscodeExe = "${env:ProgramFiles}\Microsoft VS Code\Code.exe"
    $pathResult = Get-Command code -ErrorAction SilentlyContinue
    if (Test-Path $vscodeExe) {
        Write-Host "检测到 Visual Studio Code 已安装。" -ForegroundColor Yellow
        Write-Host $vscodeExe -ForegroundColor Yellow
        return $true
    }

    # 如果未找到默认路径，尝试从 PATH 环境变量中查找
    if ($pathResult) {
        Write-Host "检测到 Visual Studio Code 已安装。" -ForegroundColor Yellow
        Write-Host $pathResult.Source -ForegroundColor Yellow
        return $true
    }

    Write-Host "检测到 Visual Studio Code 未安装。" -ForegroundColor Green
    return $false
}

if (Check-VSCodeInstalled) {
    Start-Sleep 3
}
else{
    Install-VSCode
    Start-Sleep 3
}