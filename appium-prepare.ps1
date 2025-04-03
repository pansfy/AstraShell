<#
.SYNOPSIS
Appium自动化测试环境一键配置脚本

.NOTES
需要以管理员权限运行，确保网络连接正常
#>

# 强制以管理员权限运行
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# 1. Node.js环境自动安装
function Install-NodeJS {
    param (
        [string]$Version = "v20.19.0",
        [string]$Arch = "x64"
    )

    $nodeInstallerUrl = "https://nodejs.org/dist/$Version/node-$Version-$Arch.msi"
    $nodeInstallerPath = "$env:TEMP\nodejs-$Version-$Arch.msi"
    $nodePath = "C:\Program Files\nodejs"

    try {
        Write-Host "正在下载Node.js $Version..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $nodeInstallerUrl -OutFile $nodeInstallerPath -UseBasicParsing

        Write-Host "正在静默安装Node.js..." -ForegroundColor Cyan
        $installArgs = "/i `"$nodeInstallerPath`" /quiet /norestart ADDLOCAL=ALL"
        Start-Process msiexec.exe -Wait -ArgumentList $installArgs

        # 更新系统环境变量
        $systemPath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
        if ($systemPath -notlike "*$nodePath*") {
            $env:Path += ";$nodePath"
            [Environment]::SetEnvironmentVariable('Path', "$systemPath$nodePath", 'Machine')
        }

        # 二次验证安装结果
        if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
            throw "Node.js安装后仍无法识别，可能需要手动重启终端"
        }

        Write-Host "Node.js安装验证成功" -ForegroundColor Green
        Remove-Item $nodeInstallerPath -Force
    }
    catch {
        Write-Host "Node.js安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 2. JDK环境自动安装
function Install-JDK {
    param (
        [string]$Version = "17"
    )

    try {
        $jdkBaseUrl = "https://mirror.tuna.tsinghua.edu.cn/Adoptium/$Version/jdk/x64/windows"
        $jdkParentPath = "C:\DevTools\Java"

        # 提取 JDK 下载路径
        $response = Invoke-WebRequest -Uri $jdkBaseUrl -ErrorAction Stop -UseBasicParsing
        $fileLinks = $response.Links | Where-Object { $_.href -match ".*$Version.*\.(zip)$" }
        $jdkUrl = $jdkBaseUrl + "/" + $fileLinks.href

        # 创建JDK父级目录
        if (-not (Test-Path $jdkParentPath)) {
            New-Item -Path $jdkParentPath -ItemType Directory -Force | Out-Null
        }

        # 下载并解压命令行工具
        $tempFile = "$env:TEMP\OpenJDK$Version.zip"
        Invoke-WebRequest -Uri $jdkUrl -OutFile $tempFile -UseBasicParsing
        Expand-Archive -Path $tempFile -DestinationPath "$jdkParentPath\" -Force

        # 获取解压后的目录名
        $subDirs = Get-ChildItem -Path $jdkParentPath -Directory
        $matchedDirs = Get-ChildItem -Path $jdkParentPath -Directory | Where-Object { $_.Name -match "^jdk.*$Version.*" }
        $firstMatchedDir = $matchedDirs | Select-Object -First 1
        if ($firstMatchedDir) {
           $jdkPath = $($firstMatchedDir.FullName)
        } else {
            throw "OpenJDK$Version 解压后路径获取失败，需要手动配置 JAVA_HOME"
        }
        
        # 设置环境变量
        [Environment]::SetEnvironmentVariable("JAVA_HOME", $jdkPath, "Machine")
        $env:JAVA_HOME = $jdkPath
    
        # 更新PATH变量
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($currentPath -notlike "*$jdkPath\bin*") {
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$jdkPath\bin", "Machine")
            $env:Path += ";$jdkPath\bin"
        }
    
        # 二次验证安装结果
        if (-not (Get-Command java -ErrorAction SilentlyContinue)) {
            throw "OpenJDK 安装后验证失败，可能需要手动重启终端或检查JAVA_HOME环境变量设置"
        }
        Write-Host "OpenJDK $Version 安装成功" -ForegroundColor Green
        Remove-Item $tempFile -Force
    }
    catch {
        Write-Host "OpenJDK $Version 安装失败：$_" -ForegroundColor Red
        sleep 5
        exit 1
    }
}

# 3. Android SDK安装配置
function Install-AndroidSDK {
    $androidSdkParentPath = "C:\DevTools"
    $androidSdkPath = "C:\DevTools\android-sdk-windows"
    $sdkUrl = "https://dl.google.com/android/android-sdk_r24.4.1-windows.zip"
    $platformToolsUrl = "https://googledownloads.cn/android/repository/platform-tools-latest-windows.zip"

    try {
        # 创建 Android SDK 目录
        if (-not (Test-Path $androidSdkParentPath)) {
            New-Item -Path $androidSdkParentPath -ItemType Directory -Force | Out-Null
        }
    
        # 下载并解压命令行工具
        $sdkTempFile = "$env:TEMP\android-sdk-windows.zip"
        Invoke-WebRequest -Uri $sdkUrl -OutFile $sdkTempFile -UseBasicParsing
        Expand-Archive -Path $sdkTempFile -DestinationPath "$androidSdkParentPath\" -Force

        $platformTempFile = "$env:TEMP\platform-tools-windows.zip"
        Invoke-WebRequest -Uri $platformToolsUrl -OutFile $platformTempFile -UseBasicParsing
        Expand-Archive -Path $platformTempFile -DestinationPath "$androidSdkPath\" -Force

        # 配置环境变量
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($currentPath -notlike "*$androidSdkPath") {
            $env:ANDROID_HOME = $androidSdkPath
            [Environment]::SetEnvironmentVariable("ANDROID_HOME", $androidSdkPath, "Machine")
            $env:Path += ";$androidSdkPath\platform-tools"
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$androidSdkPath\platform-tools", "Machine")
        }
    
        # 二次验证安装结果
        if (-not (Get-Command adb -ErrorAction SilentlyContinue)) {
            throw "Android SDK 安装后验证失败，可能需要手动重启终端或检查ANDROID_HOME环境变量设置"
        }
    
        Write-Host "Android SDK 安装成功" -ForegroundColor Green
        Remove-Item $sdkTempFile -Force
        Remove-Item $platformTempFile -Force
    }
    catch {
        Write-Host "Android SDK安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 4. 安装 Appium
function Install-Appium(){
    try{
        npm  -g install appium --registry=https://registry.npmmirror.com

        # 二次验证安装结果
        if (-not (Get-Command appium -ErrorAction SilentlyContinue)) {
            throw "Appium安装后验证失败，可能需要手动安装。"
        }

        Write-Host "Appium安装成功" -ForegroundColor Green
    }
    catch{
        Write-Host "Appium安装失败: $_" -ForegroundColor Red
        exit 1
    }
}

# 5. 安装 Appium UiAutomator2 驱动
function Install-UiAutomator2(){
    try{
        appium driver install uiautomator2

        # 二次验证安装结果
        $drivers = appium driver list --installed 2>&1
        if ($drivers -match "uiautomator2") {
            Write-Host "UiAutomator2 驱动安装成功" -ForegroundColor Green
        }
        else{
            throw "UiAutomator2 驱动安装失败, 可重新运行该脚本"
        }
    }
    catch{
        Write-Host "UiAutomator2 驱动安装失败: $_" -ForegroundColor Red
    }
}

# 5. 安装 Python3
function Install-Python(){
    param (
        [string]$Version = "3"
    )

    try {
        Write-Host "正在下载 Python$Version 安装程序..." -ForegroundColor Cyan
        $url = "https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe"
        $installer = "$env:TEMP\python_installer.exe"

        Invoke-WebRequest -Uri $url -OutFile $installer -ErrorAction Stop -UseBasicParsing
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

# 1. Node.js 环境检查
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host "检测到 Node.js 未安装，开始自动安装..." -ForegroundColor Yellow
    Install-NodeJS
}
else {
    Write-Host "检测到 Node.js 已安装: $(node -v)" -ForegroundColor Green
}

# 2. OpenJDK 环境检查
$javaCheck = @(
    (Get-Command java -ErrorAction SilentlyContinue),
    (Test-Path "C:\DevTools\Java\jdk*"),
    [Environment]::GetEnvironmentVariable("JAVA_HOME", "Machine")
) | Where-Object { $_ -ne $null }

if (-not $javaCheck) {
    Write-Host "检测到 JDK 未安装，开始自动安装..." -ForegroundColor Yellow
    Install-JDK -Version 17
}
else {
    Write-Host "检测到 JDK 已安装: $((Get-Command java).Source)" -ForegroundColor Green
}

# 3. ADB环境验证
if (-not (Get-Command adb -ErrorAction SilentlyContinue)) {
    Write-Host "检测到 ADB 环境未配置，开始自动安装..." -ForegroundColor Yellow
    Install-AndroidSDK
}
else {
    Write-Host "检测到 ADB 环境已配置: $((Get-Command adb).Source)" -ForegroundColor Green
}

# 4. Appium 检测安装
$globalPackages = npm list -g --json | ConvertFrom-Json
$hasAppium = $globalPackages.dependencies.PSObject.Properties.Name -contains "appium"
if ($hasAppium) {
    Write-Host "检测到 Appium 已全局安装，版本：" $globalPackages.dependencies.appium.version -ForegroundColor Green
} else {
    Write-Host "检测到 Appium 未全局安装，开始自动安装..." -ForegroundColor Red
    Install-Appium
}

# 5. appium driver: uiautomator2 检测安装
try {
    $drivers = appium driver list --installed 2>&1
    if ($drivers -match "uiautomator2") {
        Write-Host "检测到 UiAutomator2 驱动已安装。" -ForegroundColor Green
    } else {
        Write-Host "检测到 UiAutomator2 驱动未安装。" -ForegroundColor Red
        Install-UiAutomator2
    }
} catch {
    Write-Host "检查失败：请确保 Appium 已正确安装且命令可用。" -ForegroundColor Yellow
}

# 6. Python客户端安装
if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
    Write-Host "检测到 Python 未安装，开始自动安装..." -ForegroundColor Yellow
    Install-Python -Version 3
    pip install Appium-Python-Client -i https://pypi.tuna.tsinghua.edu.cn/simple
    Write-Host "Appium-Python-Client 模块安装成功" -ForegroundColor Green
}else{
    $packages = pip list 2>&1
    if ($packages -match "Appium-Python-Client") {
        Write-Host "检测到 Appium-Python-Client 模块已安装。" -ForegroundColor Green
    } else {
        Write-Host "检测到 Appium-Python-Client 模块未安装，开始自动安装..." -ForegroundColor Yellow
        pip install Appium-Python-Client -i https://pypi.tuna.tsinghua.edu.cn/simple
        Write-Host "Appium-Python-Client 模块安装成功" -ForegroundColor Green
    }
}

# 最终验证
Write-Host "`n环境验证一览：" -ForegroundColor Cyan
Write-Host "Java 环境: $((Get-Command java).Source)" -ForegroundColor Green
Write-Host "Node 环境: $((Get-Command node).Source)" -ForegroundColor Green
Write-Host "Appium 环境: $((Get-Command appium).Source)" -ForegroundColor Green
Write-Host "Python 环境: $((Get-Command python).Source)" -ForegroundColor Green

# 等待退出
Read-Host -Prompt "按 Enter 键继续..."