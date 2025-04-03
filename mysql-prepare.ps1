<#
.SYNOPSIS
MySQL 自动化测试环境一键配置脚本

.NOTES
需要以管理员权限运行，确保网络连接正常
#>

# 强制以管理员权限运行
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# 检查 Visual C++ Redistributable 是否已安装
function Test-VCRedist() {
    $vcredistInstalled = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Microsoft Visual C++ 2022*" } | Select-Object -First 1
    if (-not $vcredistInstalled) {
        Write-Host "未检测到 Visual C++ Redistributable，准备安装..." -ForegroundColor Yellow
        return $false
    }
    Write-Host "Visual C++ Redistributable 2015-2022 已安装。" -ForegroundColor Green
    return $true
}

# 下载并安装 Visual C++ Redistributable
function Install-VCRedist() {
    $vcredistUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    $vcredistInstaller = "$env:TEMP\vc_redist.x64.exe"
    Write-Host "正在下载 Visual C++ Redistributable..." -ForegroundColor Green
    Invoke-WebRequest -Uri $vcredistUrl -OutFile $vcredistInstaller

    Write-Host "正在安装 Visual C++ Redistributable..." -ForegroundColor Green
    Start-Process -FilePath $vcredistInstaller -ArgumentList "/install /quiet /norestart" -Wait

    Remove-Item $vcredistInstaller -Force
    Write-Host "Visual C++ Redistributable 安装完成。" -ForegroundColor Green
}

# 设置工作目录
$workDir = "C:\MySQL"
if (-not (Test-Path $workDir)) {
    New-Item $workDir -ItemType Directory -Force | Out-Null
}

# 提供可选的 MySQL 版本列表
$mysqlVersions = @("8.4.3", "8.0.40", "5.7.44")
Write-Host "请选择要安装的 MySQL 版本："
for ($i = 0; $i -lt $mysqlVersions.Count; $i++) {
    Write-Host "$($i + 1). $($mysqlVersions[$i])"
}
$choice = Read-Host "请输入版本编号（例如 1 表示 8.4.3）"

if ($choice -lt 1 -or $choice -gt $mysqlVersions.Count) {
    Write-Host "未提供的版本，请重试" -ForegroundColor Red
    Start-Sleep 3
    exit 1
}

$selectedVersion = $mysqlVersions[$choice - 1]
Write-Host "您选择了 MySQL 版本：$selectedVersion"

# 检查本地是否已安装 MySQL
$mysqlCheck = (Get-Command mysqld -ErrorAction SilentlyContinue)
if ($mysqlCheck -ne $null) {
    Write-Host "检测到本地已安装 MySQL，请先卸载后再继续。" -ForegroundColor Yellow
    Start-Sleep 3
    exit 1
}

# 检查并安装运行库
if (-not (Test-VCRedist)) {
    Install-VCRedist
}

try {
    # 下载 MySQL 解压版
    $major = "$($selectedVersion.Split(".")[0]).$($selectedVersion.Split(".")[1])"
    $downloadUrl = "https://cdn.mysql.com/archives/mysql-$major/mysql-$selectedVersion-winx64.zip"
    $zipFile = "$workDir\mysql-$selectedVersion-winx64.zip"
    $mysqlPath = "$workDir\mysql-$selectedVersion-winx64"
    Write-Host "正在下载 MySQL $selectedVersion..." -ForegroundColor Green
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFile  -UseBasicParsing

    # 解压文件
    Write-Host "正在解压文件..." -ForegroundColor Green
    Expand-Archive -Path $zipFile -DestinationPath "$workDir\" -Force

    # 配置 MySQL
    $configFile = "$mysqlPath\my.ini"
    New-Item -ItemType File -Path $configFile -Force | Out-Null
    $content = @(
        "[mysqld]",
        "basedir=$mysqlPath",
        "datadir=$mysqlPath\data",
        "port=3306",
        "character-set-server=utf8mb4",
        "collation-server=utf8mb4_general_ci"
    )
    $content | Set-Content -Path $configFile

    # 初始化 MySQL 数据目录
    Write-Host "正在初始化 MySQL 数据目录..." -ForegroundColor Green
    Start-Process -FilePath "$mysqlPath\bin\mysqld.exe" -ArgumentList "--initialize-insecure --console" -NoNewWindow -Wait

    # 安装 MySQL 服务
    Write-Host "正在安装 MySQL 服务..." -ForegroundColor Green
    Start-Process -FilePath "$mysqlPath\bin\mysqld.exe" -ArgumentList "--install MySQL$selectedVersion" -NoNewWindow -Wait

    # 启动 MySQL 服务
    Write-Host "正在启动 MySQL 服务..." -ForegroundColor Green
    Start-Service -Name "MySQL$selectedVersion"

    # 添加MySQL环境变量
    $currentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    if (-not $currentPath.Contains($mysqlPath)) {
        $env:PATH += ";$mysqlPath\bin"
        [Environment]::SetEnvironmentVariable("PATH", $currentPath + ";$mysqlPath\bin", "Machine")
    }

    # 完成提示
    Write-Host "MySQL $selectedVersion 已成功安装并启动！默认无密码" -ForegroundColor Green
    Write-Host "可使用【 mysql -uroot -p 】测试连接..." -ForegroundColor Green
}
catch{
    Write-Host "MySQL $selectedVersion 安装失败：$_" -ForegroundColor Red
}
finnaly{
    # 等待退出
    Read-Host -Prompt "按 Enter 键继续..."
}