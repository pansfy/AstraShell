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
    # 通过注册表检测（64位和32位路径）
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installed = Get-ItemProperty $regPaths | Where-Object {
        $_.DisplayName -like "微信"
    }
    if ($installed) { return $true }

    # 通过文件路径二次验证（默认安装目录）
    $installPath = "${env:ProgramFiles(x86)}\Tencent\WeChat"
    if (Test-Path $installPath) { return $true }

    return $false
}

# 静默安装微信
function Install-WeChatSilently {
    $tempDir = $env:TEMP
    $installer = "$tempDir\WeChatSetup.exe"

    try {
        # 下载微信官方安装包（需替换为可靠链接）
        Write-Host "正在下载微信安装程序..." -ForegroundColor Cyan
        $WXUrl = "https://dldir1.qq.com/weixin/Windows/WeChatSetup.exe"
        Invoke-WebRequest -Uri $WXUrl -OutFile $installer -UseBasicParsing

        # 静默安装（微信默认静默参数为 `/S`）
        Write-Host "正在静默安装微信..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow

        # 等待安装完成并清理临时文件
        Start-Sleep -Seconds 5
        Remove-Item $installer -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "微信安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 检查 QQ 是否已安装
function Test-QQInstalled {
    # 1. 注册表检测（64位/32位路径）
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $qqReg = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "QQ" }

    # 2. 默认安装目录检测
    $installPaths = @(
        "${env:ProgramFiles}\Tencent\QQ",
        "${env:ProgramFiles}\Tencent\QQNT",
        "${env:ProgramFiles(x86)}\Tencent\QQ",
        "${env:ProgramFiles(x86)}\Tencent\QQNT"
    )
    $qqDir = $installPaths | Where-Object { Test-Path $_ }

    return ($qqReg -ne $null) -or ($qqDir -ne $null)
}

# 静默安装 QQ
function Install-QQSilently {
    $tempDir = $env:TEMP
    $installer = "$tempDir\QQInstaller.exe"
    $downloadUrl ="https://dldir1.qq.com/qqfile/qq/QQNT/Windows/QQ_9.9.18_250318_x64_01.exe"

    try {
        # 下载安装包
        Write-Host "正在下载 QQ 安装程序..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing

        # 静默安装（QQ 默认静默参数为 `/S`）
        Write-Host "正在静默安装..." -ForegroundColor Cyan
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow

        # 等待安装完成并清理
        Start-Sleep -Seconds 5
        Remove-Item $installer -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "QQ安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 检查 WPS Office 是否已安装
function Test-WPSInstalled {
    # 1. 注册表检测（适配新旧版本）
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $wpsReg = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "WPS Office" }

    # 2. 默认安装目录检测
    $installPaths = @(
        "${env:ProgramFiles}\WPS Office",
        "${env:ProgramFiles(x86)}\WPS Office"
    )
    $wpsDir = $installPaths | Where-Object { Test-Path $_ }

    return ($wpsReg -ne $null) -or ($wpsDir -ne $null)
}

# 静默安装 WPS Office
function Install-WPSOfficeSilently {
    $tempDir = $env:TEMP
    $installerPath = Join-Path $tempDir "WPSOffice.exe"
    $downloadUrl = "https://official-package.wpscdn.cn/wps/download/WPS_Setup_X64_20288.exe"

    try {
        # 下载安装包
        Write-Host "正在下载 WPS Office..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing

        # 静默安装参数（/S 或 /silent）
        Write-Host "正在静默安装 WPS Office..." -ForegroundColor Cyan
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait -NoNewWindow

        # 等待安装完成（建议根据实际情况调整等待时间）
        Start-Sleep -Seconds 5

        # 清理临时文件
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "WPS Office安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 检查 PDF 是否已安装
function Test-PDFInstalled {
    # 1. 注册表检测（适配新旧版本）
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $pdfReg = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "PDF" }

    # 2. 默认安装目录检测
    $installPaths = @(
        "${env:ProgramFiles}\SumatraPDF",
        "${env:ProgramFiles(x86)}\SumatraPDF"
    )
    $pdfDir = $installPaths | Where-Object { Test-Path $_ }

    return ($pdfReg -ne $null) -or ($pdfDir -ne $null)
}

# 静默安装 PDF
function Install-PDFSilently {
    $tempDir = $env:TEMP
    $installerPath = Join-Path $tempDir "SumatraPDF.exe"
    $downloadUrl = "https://www.sumatrapdfreader.org/dl/rel/3.5.2/SumatraPDF-3.5.2-64-install.exe"

    try {
        # 下载安装包
        Write-Host "正在下载 SumatraPDF..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing

        # 静默安装参数（/S 或 /silent）
        Write-Host "正在静默安装 SumatraPDF..." -ForegroundColor Cyan
        Start-Process -FilePath $installerPath -ArgumentList "-s -all-users" -Wait -NoNewWindow

        # 等待安装完成（建议根据实际情况调整等待时间）
        Start-Sleep -Seconds 5

        # 清理临时文件
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "SumatraPDF安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 检查 BandiView 是否已安装
function Test-BandiViewInstalled {
    # 1. 注册表检测（适配新旧版本）
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $viewReg = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "BandiView" }

    # 2. 默认安装目录检测
    $installPaths = @(
        "${env:ProgramFiles}\\BandiView",
        "${env:ProgramFiles(x86)}\\BandiView"
    )
    $viewDir = $installPaths | Where-Object { Test-Path $_ }

    return ($viewReg -ne $null) -or ($viewDir -ne $null)
}

# 静默安装 BandiView
function Install-BandiViewSilently {
    $tempDir = $env:TEMP
    $installerPath = Join-Path $tempDir "BandiView.exe"
    $downloadUrl = "https://bandisoft.okconnect.cn/bandiview/BANDIVIEW-SETUP-X64.EXE"

    try {
        # 下载安装包
        Write-Host "正在下载 BandiView..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing

        # 静默安装参数（/S 或 /silent）
        Write-Host "正在静默安装 BandiView..." -ForegroundColor Cyan
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait -NoNewWindow

        # 等待安装完成（建议根据实际情况调整等待时间）
        Start-Sleep -Seconds 5

        # 清理临时文件
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "BandiView 安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 检查 腾讯会议 是否已安装
function Test-MettingInstalled {
    # 1. 注册表检测（适配新旧版本）
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $mettingReg = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "腾讯会议" }

    # 2. 默认安装目录检测
    $installPaths = @(
        "${env:ProgramFiles}\\WeMeet",
        "${env:ProgramFiles(x86)}\\WeMeet"
    )
    $mettingDir = $installPaths | Where-Object { Test-Path $_ }

    return ($mettingReg -ne $null) -or ($mettingDir -ne $null)
}

# 静默安装腾讯会议
function Install-MettingSilently {
    $tempDir = $env:TEMP
    $installerPath = Join-Path $tempDir "TencentMeeting.exe"
    $downloadUrl = ""

    try {
        # 下载安装包
        Write-Host "正在下载腾讯会议..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing

        # 静默安装参数（/SilentInstall=0）
        Write-Host "正在静默安装腾讯会议..." -ForegroundColor Cyan
        Start-Process -FilePath $installerPath -ArgumentList "/SilentInstall=0" -Wait -NoNewWindow

        # 等待安装完成（建议根据实际情况调整等待时间）
        Start-Sleep -Seconds 5

        # 清理临时文件
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "腾讯会议安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 1. 检查微信安装
if (-not (Test-WeChatInstalled)) {
    Write-Host "检测到微信未安装，开始执行静默安装..." -ForegroundColor Yellow
    Install-WeChatSilently
    Write-Host "微信安装完成！" -ForegroundColor Green
}
else {
    Write-Host "检测到微信已安装，跳过安装步骤。" -ForegroundColor Blue
}

# 2. 检查QQ安装
if (-not (Test-QQInstalled)) {
    Write-Host "检测到QQ未安装，开始执行静默安装..." -ForegroundColor Yellow
    Install-QQSilently
    Write-Host "QQ安装完成！" -ForegroundColor Green
}
else {
    Write-Host "检测到QQ已安装，跳过安装步骤。" -ForegroundColor Blue
}

# 3. 检查WPS安装
if (-not (Test-WPSInstalled)) {
    Write-Host "检测到WPS未安装，开始执行静默安装..." -ForegroundColor Yellow
    Install-WPSOfficeSilently
    Write-Host "WPS安装完成！" -ForegroundColor Green
}
else {
    Write-Host "检测到WPS已安装，跳过安装步骤。" -ForegroundColor Blue
}

# 4. 检查SumatraPDF安装
if (-not (Test-PDFInstalled)) {
    Write-Host "检测到PDF未安装，开始执行静默安装..." -ForegroundColor Yellow
    Install-PDFSilently
    Write-Host "PDF安装完成！" -ForegroundColor Green
}
else {
    Write-Host "检测到PDF已安装，跳过安装步骤。" -ForegroundColor Blue
}

# 5. 检查图片查看器安装
if (-not (Test-BandiViewInstalled)) {
    Write-Host "检测到图片查看器未安装，开始执行静默安装..." -ForegroundColor Yellow
    Install-BandiViewSilently
    Write-Host "图片查看器安装完成！" -ForegroundColor Green
}
else {
    Write-Host "检测到图片查看器已安装，跳过安装步骤。" -ForegroundColor Blue
}

# 等待退出
Read-Host -Prompt "按 Enter 键继续..."