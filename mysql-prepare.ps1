<#
.SYNOPSIS
MySQL �Զ������Ի���һ�����ýű�

.NOTES
��Ҫ�Թ���ԱȨ�����У�ȷ��������������
#>

# ǿ���Թ���ԱȨ������
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# ��� Visual C++ Redistributable �Ƿ��Ѱ�װ
function Test-VCRedist() {
    $vcredistInstalled = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Microsoft Visual C++ 2022*" } | Select-Object -First 1
    if (-not $vcredistInstalled) {
        Write-Host "δ��⵽ Visual C++ Redistributable��׼����װ..." -ForegroundColor Yellow
        return $false
    }
    Write-Host "Visual C++ Redistributable 2015-2022 �Ѱ�װ��" -ForegroundColor Green
    return $true
}

# ���ز���װ Visual C++ Redistributable
function Install-VCRedist() {
    $vcredistUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    $vcredistInstaller = "$env:TEMP\vc_redist.x64.exe"
    Write-Host "�������� Visual C++ Redistributable..." -ForegroundColor Green
    Invoke-WebRequest -Uri $vcredistUrl -OutFile $vcredistInstaller

    Write-Host "���ڰ�װ Visual C++ Redistributable..." -ForegroundColor Green
    Start-Process -FilePath $vcredistInstaller -ArgumentList "/install /quiet /norestart" -Wait

    Remove-Item $vcredistInstaller -Force
    Write-Host "Visual C++ Redistributable ��װ��ɡ�" -ForegroundColor Green
}

# ���ù���Ŀ¼
$workDir = "C:\MySQL"
if (-not (Test-Path $workDir)) {
    New-Item $workDir -ItemType Directory -Force | Out-Null
}

# �ṩ��ѡ�� MySQL �汾�б�
$mysqlVersions = @("8.4.3", "8.0.40", "5.7.44")
Write-Host "��ѡ��Ҫ��װ�� MySQL �汾��"
for ($i = 0; $i -lt $mysqlVersions.Count; $i++) {
    Write-Host "$($i + 1). $($mysqlVersions[$i])"
}
$choice = Read-Host "������汾��ţ����� 1 ��ʾ 8.4.3��"

if ($choice -lt 1 -or $choice -gt $mysqlVersions.Count) {
    Write-Host "δ�ṩ�İ汾��������" -ForegroundColor Red
    Start-Sleep 3
    exit 1
}

$selectedVersion = $mysqlVersions[$choice - 1]
Write-Host "��ѡ���� MySQL �汾��$selectedVersion"

# ��鱾���Ƿ��Ѱ�װ MySQL
$mysqlCheck = (Get-Command mysqld -ErrorAction SilentlyContinue)
if ($mysqlCheck -ne $null) {
    Write-Host "��⵽�����Ѱ�װ MySQL������ж�غ��ټ�����" -ForegroundColor Yellow
    Start-Sleep 3
    exit 1
}

# ��鲢��װ���п�
if (-not (Test-VCRedist)) {
    Install-VCRedist
}

try {
    # ���� MySQL ��ѹ��
    $major = "$($selectedVersion.Split(".")[0]).$($selectedVersion.Split(".")[1])"
    $downloadUrl = "https://cdn.mysql.com/archives/mysql-$major/mysql-$selectedVersion-winx64.zip"
    $zipFile = "$workDir\mysql-$selectedVersion-winx64.zip"
    $mysqlPath = "$workDir\mysql-$selectedVersion-winx64"
    Write-Host "�������� MySQL $selectedVersion..." -ForegroundColor Green
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFile  -UseBasicParsing

    # ��ѹ�ļ�
    Write-Host "���ڽ�ѹ�ļ�..." -ForegroundColor Green
    Expand-Archive -Path $zipFile -DestinationPath "$workDir\" -Force

    # ���� MySQL
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

    # ��ʼ�� MySQL ����Ŀ¼
    Write-Host "���ڳ�ʼ�� MySQL ����Ŀ¼..." -ForegroundColor Green
    Start-Process -FilePath "$mysqlPath\bin\mysqld.exe" -ArgumentList "--initialize-insecure --console" -NoNewWindow -Wait

    # ��װ MySQL ����
    Write-Host "���ڰ�װ MySQL ����..." -ForegroundColor Green
    Start-Process -FilePath "$mysqlPath\bin\mysqld.exe" -ArgumentList "--install MySQL$selectedVersion" -NoNewWindow -Wait

    # ���� MySQL ����
    Write-Host "�������� MySQL ����..." -ForegroundColor Green
    Start-Service -Name "MySQL$selectedVersion"

    # ���MySQL��������
    $currentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    if (-not $currentPath.Contains($mysqlPath)) {
        $env:PATH += ";$mysqlPath\bin"
        [Environment]::SetEnvironmentVariable("PATH", $currentPath + ";$mysqlPath\bin", "Machine")
    }

    # �����ʾ
    Write-Host "MySQL $selectedVersion �ѳɹ���װ��������Ĭ��������" -ForegroundColor Green
    Write-Host "��ʹ�á� mysql -uroot -p ����������..." -ForegroundColor Green
}
catch{
    Write-Host "MySQL $selectedVersion ��װʧ�ܣ�$_" -ForegroundColor Red
}
finnaly{
    # �ȴ��˳�
    Read-Host -Prompt "�� Enter ������..."
}