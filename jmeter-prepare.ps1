<#
.SYNOPSIS
JMeter 自动化测试环境一键配置脚本
https://archive.apache.org/dist/jmeter/binaries/

.NOTES
需要以管理员权限运行，确保网络连接正常
#>

# 强制以管理员权限运行
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# OpenJDK 安装配置
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

# 1. OpenJDK 环境检查
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

# JMeter 安装配置
function Install-JMeter(){
    param(
        [string]$Version = "5.6.2"
    )
    try{
        $JMeterParentPath = "C:\DevTools"
        $JMeterPath = "$JMeterParentPath\apache-jmeter-$Version"
        $JMeterDownloadUrl = "https://dlcdn.apache.org//jmeter/binaries/apache-jmeter-$Version.zip"

        # 创建JDK父级目录
        if (-not (Test-Path $JMeterParentPath)) {
            New-Item -Path $JMeterParentPath -ItemType Directory -Force | Out-Null
        }

        # 下载并解压命令行工具
        $tempFile = "$env:TEMP\jmeter-$Version.zip"
        Invoke-WebRequest -Uri $JMeterDownloadUrl -OutFile $tempFile -UseBasicParsing
        Expand-Archive -Path $tempFile -DestinationPath "$JMeterParentPath\" -Force

        $env:Path += ";$JMeterPath\bin"
        [Environment]::SetEnvironmentVariable("Path", "$currentPath;$JMeterPath\bin", "Machine")

        # 二次验证安装结果
        if (-not (Get-Command jmeter -ErrorAction SilentlyContinue)) {
            throw "JMeter 安装后验证失败，可检查安装情况：$JMeterPath"
        }
        Write-Host "JMeter $Version 安装成功" -ForegroundColor Green
        Remove-Item $tempFile -Force
    }
    catch {
        Write-Host "JMeter $Version 安装失败：$_" -ForegroundColor Red
        sleep 5
        exit 1
    }
}

# 2. JMeter 环境检查
if (-not (Get-Command jmeter -ErrorAction SilentlyContinue)) {
   Write-Host "检测到 JMeter 未安装，开始自动安装..." -ForegroundColor Yellow
   Install-JMeter -Version "5.6.3"
}
else{
    Write-Host "检测到 JMeter 已安装: $((Get-Command jmeter).Source)" -ForegroundColor Green
}

# 等待退出
Read-Host -Prompt "按 Enter 键继续..."