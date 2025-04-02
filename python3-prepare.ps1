<#
.SYNOPSIS
Python 自动化测试环境一键配置脚本

.NOTES
需要以管理员权限运行，确保网络连接正常
#>

# 强制以管理员权限运行
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Install-Python(){
    param (
        [string]$Version = "3.12.9"
    )

    try {
        Write-Host "正在下载 Python$Version 安装程序..." -ForegroundColor Cyan
        $url = "https://www.python.org/ftp/python/$Version/python-$Version-amd64.exe"
        $installer = "$env:TEMP\python_installer.exe"

        Invoke-WebRequest -Uri $url -OutFile $installer -ErrorAction Stop
        Write-Host "正在静默安装 Python..." -ForegroundColor Cyan
        Start-Process $installer -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait -NoNewWindow

        $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')
    
        # 二次验证安装结果
        if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
            throw "Python 安装后验证失败，可能需要手动重启终端或检查Python环境变量设置"
        }
        Write-Host "Python $Version 安装成功" -ForegroundColor Green
        Remove-Item $installer -Force
    }
    catch {
        Write-Host "Python $Version 安装失败：$_" -ForegroundColor Red
        exit 1
    }
}

# 检查 PyCharm Community Edition 是否安装
function Test-PyCharmInstalled {
    # 1. 检查注册表中的安装信息（Windows）
    $regPaths = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $pycharmReg = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
                  Where-Object { $_.DisplayName -match "PyCharm Community Edition (\d+\.\d+)?" }

    # 2. 检查默认安装路径
    $programFilesPaths = @(
        "${env:ProgramFiles}\JetBrains\PyCharm Community Edition *",
        "${env:ProgramFiles(x86)}\JetBrains\PyCharm Community Edition *"
    )
    $pycharmDir = Get-Item -Path $programFilesPaths -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match "PyCharm Community Edition \d+\.\d+" }

    # 3. 检查 JetBrains Toolbox 安装路径（用户级）
    $toolboxPath = "${env:APPDATA}\JetBrains\Toolbox\apps\PyCharm-C\*\ch-0\*\bin\pycharm64.exe"
    $pycharmExe = Get-Item -Path $toolboxPath -ErrorAction SilentlyContinue

    # 汇总结果
    if ($pycharmReg -or $pycharmDir -or $pycharmExe) {
        $installPath = if ($pycharmDir) { $pycharmDir.FullName }
                       elseif ($pycharmExe) { $pycharmExe.DirectoryName }
                       else { $pycharmReg.InstallLocation }
        return @{
            Installed = $true
            Path      = $installPath
        }
    } else {
        return @{ Installed = $false }
    }
}

function Install-PyCharm(){
    param(
        [string]$pycharmVersion = "2024.3.5"
    )
    Write-Host "正在下载 PyCharm Community 安装程序..." -ForegroundColor Cyan
    $pycharmUrl = "https://download.jetbrains.com/python/pycharm-community-$pycharmVersion.exe"
    $installer = "$env:TEMP\pycharm_installer.exe"

    Invoke-WebRequest -Uri $pycharmUrl -OutFile $installer -ErrorAction Stop
    Write-Host "正在静默安装 PyCharm Community..." -ForegroundColor Cyan
    Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow

    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath "PyCharm Community.lnk"
    $programFilesPaths = @(
        "${env:ProgramFiles}\JetBrains\PyCharm Community Edition $pycharmVersion",
        "${env:ProgramFiles(x86)}\JetBrains\PyCharm Community Edition $pycharmVersion"
    )
    $pycharmDir = Get-Item -Path $programFilesPaths -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match "PyCharm Community Edition \d+\.\d+" }
    $targetPath = $pycharmDir.FullName
    if (Test-Path $targetPath) {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = "$targetPath\bin\pycharm64.exe"
        $Shortcut.Save()
    }

    Remove-Item $installer -Force -ErrorAction SilentlyContinue
    Write-Host "PyCharm Community $pycharmVersion 安装成功" -ForegroundColor Green
}

# 1. Python客户端安装
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "检测到 Python 未安装，开始自动安装..." -ForegroundColor Yellow
    Install-Python -Version "3.12.9"
}else{
    Write-Host "检测到 Python3 已安装。" -ForegroundColor Green
}

# 2. PyCharm Community 安装
$result = Test-PyCharmInstalled
if ($result.Installed) {
    Write-Host "PyCharm Community Edition 已安装。路径: $($result.Path)" -ForegroundColor Green
} else {
    Write-Host "未检测到 PyCharm Community Edition，开始自动安装..." -ForegroundColor Red
    Install-PyCharm -pycharmVersion "2024.3.5"
}

# 等待退出
Read-Host -Prompt "按 Enter 键继续..."