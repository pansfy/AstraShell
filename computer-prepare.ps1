<#
.SYNOPSIS
基本办公环境 自动化测试环境一键配置脚本

.NOTES
需要以管理员权限运行，确保网络连接正常
#>

# 强制以管理员权限运行
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# 检测微信是否已安装
function Test-WeChatInstalled {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty $regPaths | Where-Object {
        $_.DisplayName -like "微信"
    }

    if ($installed) { return $true }

    return $false
}

# 静默安装微信
function Install-WeChatSilently {
    $installer = "$env:TEMP\WeChatSetup.exe"
    $downloadUrl = "https://dldir1.qq.com/weixin/Windows/WeChatSetup.exe"

    if (Test-WeChatInstalled) {
        Write-Host "检测到 微信 已安装，跳过安装步骤。" -ForegroundColor Blue
        return
    }

    Write-Host "检测到 微信 未安装，开始执行静默安装..." -ForegroundColor Yellow
    try {
        Write-Host "正在下载微信安装程序..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        Write-Host "正在静默安装微信..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow
        Write-Host "微信安装完成。" -ForegroundColor Green
    } catch {
        Write-Host "微信安装失败: $_" -ForegroundColor Red
    } finally {
        if(Test-Path $installer){
            Remove-Item $installer -Force -ErrorAction SilentlyContinue
        }
    }
}

# 检查 QQ 是否已安装
function Test-QQInstalled {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "QQ" }

    if ($installed) { return $true }

    return $false
}

# 静默安装 QQ
function Install-QQSilently {
    $installer = "$env:TEMP\QQInstaller.exe"
    $downloadUrl ="https://dldir1.qq.com/qqfile/qq/QQNT/Windows/QQ_9.9.18_250318_x64_01.exe"

    if (Test-QQInstalled) {
        Write-Host "检测到 QQ 已安装，跳过安装步骤。" -ForegroundColor Blue
        return
    }

    Write-Host "检测到 QQ 未安装，开始执行静默安装..." -ForegroundColor Yellow
    try {
        Write-Host "正在下载 QQ 安装程序..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        Write-Host "正在静默安装 QQ..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow
        Write-Host "QQ 安装完成。" -ForegroundColor Green
    } catch {
        Write-Host "QQ 安装失败: $_" -ForegroundColor Red
    } finally {
        if(Test-Path $installer){
            Remove-Item $installer -Force -ErrorAction SilentlyContinue
        }
    }
}

# 检查 WPS Office 是否已安装
function Test-OfficeInstalled {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "Office" }
    
    if ($installed) { return $true }

    return $false
}

# 静默安装 WPS Office
function Install-OfficeSilently {
    $installer = "$env:TEMP\WPSOffice.exe"
    $downloadUrl = "https://official-package.wpscdn.cn/wps/download/WPS_Setup_20305.exe"

    if (Test-OfficeInstalled) {
        Write-Host "检测到 Office 已安装，跳过安装步骤。" -ForegroundColor Blue
        return
    }

    Write-Host "检测到 Office 未安装，开始执行静默安装..." -ForegroundColor Yellow
    try {
        Write-Host "正在下载 WPS Office..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        Write-Host "正在静默安装 WPS Office..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/S -agreelicense" -Wait -NoNewWindow
        Write-Host "WPS Office 安装完成。" -ForegroundColor Green
    } catch {
        Write-Host "WPS Office 安装失败: $_" -ForegroundColor Red
    } finally {
        if(Test-Path $installer){
            Remove-Item $installer -Force -ErrorAction SilentlyContinue
        }
    }
}

# 检查 PDF 是否已安装
function Test-PDFInstalled {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "PDF" }

    if ($installed) { return $true }

    return $false
}

# 静默安装 PDF
function Install-PDFSilently {
    $installer = "$env:TEMP\SumatraPDF.exe"
    $downloadUrl = "https://www.sumatrapdfreader.org/dl/rel/3.5.2/SumatraPDF-3.5.2-64-install.exe"

    if (Test-PDFInstalled) {
        Write-Host "检测到 PDF 已安装，跳过安装步骤。" -ForegroundColor Blue
        return
    }

    Write-Host "检测到 PDF 未安装，开始执行静默安装..." -ForegroundColor Yellow
    try {
        Write-Host "正在下载 SumatraPDF..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        Write-Host "正在静默安装 SumatraPDF..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "-s -all-users -install" -Wait -NoNewWindow
        Write-Host "SumatraPDF 安装完成。" -ForegroundColor Green
    } catch {
        Write-Host "SumatraPDF 安装失败: $_" -ForegroundColor Red
    } finally {
        if(Test-Path $installer){
            Remove-Item $installer -Force -ErrorAction SilentlyContinue
        }
    }
}

# 检查 BandiView 是否已安装
function Test-BandiViewInstalled {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "BandiView" }
    
    if ($installed) { return $true }

    return $false
}

# 静默安装 BandiView
function Install-BandiViewSilently {
    $installer = "$env:TEMP\BandiView.exe"
    $downloadUrl = "https://bandisoft.okconnect.cn/bandiview/BANDIVIEW-SETUP-X64.EXE"

    if (Test-BandiViewInstalled) {
        Write-Host "检测到图片查看器已安装，跳过安装步骤。" -ForegroundColor Blue
        return
    }

    Write-Host "检测到图片查看器未安装，开始执行静默安装..." -ForegroundColor Yellow
    try {
        Write-Host "正在下载 BandiView..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        Write-Host "正在静默安装 BandiView..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow
        Write-Host "BandiView 安装完成。" -ForegroundColor Green
    } catch {
        Write-Host "BandiView 安装失败: $_" -ForegroundColor Red
    } finally {
        if(Test-Path $installer){
            Remove-Item $installer -Force -ErrorAction SilentlyContinue
        }
    }
}

# 检查 腾讯会议 是否已安装
function Test-MettingInstalled {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "腾讯会议" }

    if ($installed) { return $true }

    return $false
}

# 静默安装腾讯会议
function Install-MettingSilently {
    $installer = "$env:TEMP\TencentMeeting.exe"
    $downloadUrl = "https://updatecdn.meeting.qq.com/cos/2a15c280af093abad029dc0383d2f184/TencentMeeting_0300000000_3.32.2.413_x86_64.publish.officialwebsite.exe"

    if (Test-MettingInstalled) {
        Write-Host "检测到腾讯会议已安装，跳过安装步骤。" -ForegroundColor Blue
        return
    }

    Write-Host "检测到腾讯会议未安装，开始执行静默安装..." -ForegroundColor Yellow
    try {
        Write-Host "正在下载腾讯会议..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        Write-Host "正在静默安装腾讯会议..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/SilentInstall=0" -Wait -NoNewWindow

    } catch {
        Write-Host "腾讯会议安装失败: $_" -ForegroundColor Red
    } finally {
        if(Test-Path $installer){
            Remove-Item $installer -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-7ZipInstalled {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "7-Zip" }

    if ($installed) { return $true }

    return $false
}

# 安装 7-Zip
function Install-7ZipSilently {
    $installer = "$env:TEMP\7Zip.exe"
    $downloadUrl = "https://nchc.dl.sourceforge.net/project/sevenzip/7-Zip/24.09/7z2409-x64.exe"

    if (Test-7ZipInstalled) {
        Write-Host "检测到7-Zip解压缩工具已安装，跳过安装步骤。" -ForegroundColor Blue
        return
    }

    Write-Host "检测到7-Zip解压缩工具未安装，开始执行静默安装..." -ForegroundColor Yellow
    try {
        Write-Host "正在下载7-Zip解压缩工具..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        Write-Host "正在静默安装7-Zip解压缩工具..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow

    } catch {
        Write-Host "7-Zip解压缩工具安装失败: $_" -ForegroundColor Red
    } finally {
        if(Test-Path $installer){
            Remove-Item $installer -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-aDriveInstalled {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "阿里云盘" }

    if ($installed) { return $true }

    return $false
}

# 安装 阿里云盘
function Install-aDriveSilently{
    $installer = "$env:TEMP\aDrive.exe"
    $downloadUrl = "https://cdn.aliyundrive.net/downloads/apps/desktop/aDrive-6.8.6.exe"

    if (Test-MettingInstalled) {
        Write-Host "检测到阿里云盘已安装，跳过安装步骤。" -ForegroundColor Blue
        return
    }

    Write-Host "检测到阿里云盘未安装，开始执行静默安装..." -ForegroundColor Yellow
    try {
        Write-Host "正在下载阿里云盘..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        Write-Host "正在静默安装阿里云盘..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow

    } catch {
        Write-Host "阿里云盘安装失败: $_" -ForegroundColor Red
    } finally {
        if(Test-Path $installer){
            Remove-Item $installer -Force -ErrorAction SilentlyContinue
        }
    }
}

Install-WeChatSilently
Install-QQSilently
Install-OfficeSilently
Install-PDFSilently
Install-BandiViewSilently
Install-7ZipSilently
Install-aDriveSilently

# 等待退出
Read-Host -Prompt "已完成，按任意键退出..."