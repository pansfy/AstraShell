<#
.SYNOPSIS
Appium�Զ������Ի���һ�����ýű�

.NOTES
��Ҫ�Թ���ԱȨ�����У�ȷ��������������
#>

# ǿ���Թ���ԱȨ������
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# 1. Node.js�����Զ���װ
function Install-NodeJS {
    param (
        [string]$Version = "v20.19.0",
        [string]$Arch = "x64"
    )

    $nodeInstallerUrl = "https://nodejs.org/dist/$Version/node-$Version-$Arch.msi"
    $nodeInstallerPath = "$env:TEMP\nodejs-$Version-$Arch.msi"
    $nodePath = "C:\Program Files\nodejs"

    try {
        Write-Host "��������Node.js $Version..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $nodeInstallerUrl -OutFile $nodeInstallerPath -UseBasicParsing

        Write-Host "���ھ�Ĭ��װNode.js..." -ForegroundColor Cyan
        $installArgs = "/i `"$nodeInstallerPath`" /quiet /norestart ADDLOCAL=ALL"
        Start-Process msiexec.exe -Wait -ArgumentList $installArgs

        # ����ϵͳ��������
        $systemPath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
        if ($systemPath -notlike "*$nodePath*") {
            $env:Path += ";$nodePath"
            [Environment]::SetEnvironmentVariable('Path', "$systemPath$nodePath", 'Machine')
        }

        # ������֤��װ���
        if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
            throw "Node.js��װ�����޷�ʶ�𣬿�����Ҫ�ֶ������ն�"
        }

        Write-Host "Node.js��װ��֤�ɹ�" -ForegroundColor Green
        Remove-Item $nodeInstallerPath -Force
    }
    catch {
        Write-Host "Node.js��װʧ��: $_" -ForegroundColor Red
        exit 1
    }
}

# 2. JDK�����Զ���װ
function Install-JDK {
    param (
        [string]$Version = "17"
    )

    try {
        $jdkBaseUrl = "https://mirror.tuna.tsinghua.edu.cn/Adoptium/$Version/jdk/x64/windows"
        $jdkParentPath = "C:\DevTools\Java"

        # ��ȡ JDK ����·��
        $response = Invoke-WebRequest -Uri $jdkBaseUrl -ErrorAction Stop -UseBasicParsing
        $fileLinks = $response.Links | Where-Object { $_.href -match ".*$Version.*\.(zip)$" }
        $jdkUrl = $jdkBaseUrl + "/" + $fileLinks.href

        # ����JDK����Ŀ¼
        if (-not (Test-Path $jdkParentPath)) {
            New-Item -Path $jdkParentPath -ItemType Directory -Force | Out-Null
        }

        # ���ز���ѹ�����й���
        $tempFile = "$env:TEMP\OpenJDK$Version.zip"
        Invoke-WebRequest -Uri $jdkUrl -OutFile $tempFile -UseBasicParsing
        Expand-Archive -Path $tempFile -DestinationPath "$jdkParentPath\" -Force

        # ��ȡ��ѹ���Ŀ¼��
        $subDirs = Get-ChildItem -Path $jdkParentPath -Directory
        $matchedDirs = Get-ChildItem -Path $jdkParentPath -Directory | Where-Object { $_.Name -match "^jdk.*$Version.*" }
        $firstMatchedDir = $matchedDirs | Select-Object -First 1
        if ($firstMatchedDir) {
           $jdkPath = $($firstMatchedDir.FullName)
        } else {
            throw "OpenJDK$Version ��ѹ��·����ȡʧ�ܣ���Ҫ�ֶ����� JAVA_HOME"
        }
        
        # ���û�������
        [Environment]::SetEnvironmentVariable("JAVA_HOME", $jdkPath, "Machine")
        $env:JAVA_HOME = $jdkPath
    
        # ����PATH����
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($currentPath -notlike "*$jdkPath\bin*") {
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$jdkPath\bin", "Machine")
            $env:Path += ";$jdkPath\bin"
        }
    
        # ������֤��װ���
        if (-not (Get-Command java -ErrorAction SilentlyContinue)) {
            throw "OpenJDK ��װ����֤ʧ�ܣ�������Ҫ�ֶ������ն˻���JAVA_HOME������������"
        }
        Write-Host "OpenJDK $Version ��װ�ɹ�" -ForegroundColor Green
        Remove-Item $tempFile -Force
    }
    catch {
        Write-Host "OpenJDK $Version ��װʧ�ܣ�$_" -ForegroundColor Red
        sleep 5
        exit 1
    }
}

# 3. Android SDK��װ����
function Install-AndroidSDK {
    $androidSdkParentPath = "C:\DevTools"
    $androidSdkPath = "C:\DevTools\android-sdk-windows"
    $sdkUrl = "https://dl.google.com/android/android-sdk_r24.4.1-windows.zip"
    $platformToolsUrl = "https://googledownloads.cn/android/repository/platform-tools-latest-windows.zip"

    try {
        # ���� Android SDK Ŀ¼
        if (-not (Test-Path $androidSdkParentPath)) {
            New-Item -Path $androidSdkParentPath -ItemType Directory -Force | Out-Null
        }
    
        # ���ز���ѹ�����й���
        $sdkTempFile = "$env:TEMP\android-sdk-windows.zip"
        Invoke-WebRequest -Uri $sdkUrl -OutFile $sdkTempFile -UseBasicParsing
        Expand-Archive -Path $sdkTempFile -DestinationPath "$androidSdkParentPath\" -Force

        $platformTempFile = "$env:TEMP\platform-tools-windows.zip"
        Invoke-WebRequest -Uri $platformToolsUrl -OutFile $platformTempFile -UseBasicParsing
        Expand-Archive -Path $platformTempFile -DestinationPath "$androidSdkPath\" -Force

        # ���û�������
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($currentPath -notlike "*$androidSdkPath") {
            $env:ANDROID_HOME = $androidSdkPath
            [Environment]::SetEnvironmentVariable("ANDROID_HOME", $androidSdkPath, "Machine")
            $env:Path += ";$androidSdkPath\platform-tools"
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$androidSdkPath\platform-tools", "Machine")
        }
    
        # ������֤��װ���
        if (-not (Get-Command adb -ErrorAction SilentlyContinue)) {
            throw "Android SDK ��װ����֤ʧ�ܣ�������Ҫ�ֶ������ն˻���ANDROID_HOME������������"
        }
    
        Write-Host "Android SDK ��װ�ɹ�" -ForegroundColor Green
        Remove-Item $sdkTempFile -Force
        Remove-Item $platformTempFile -Force
    }
    catch {
        Write-Host "Android SDK��װʧ��: $_" -ForegroundColor Red
        exit 1
    }
}

# 4. ��װ Appium
function Install-Appium(){
    try{
        npm  -g install appium --registry=https://registry.npmmirror.com

        # ������֤��װ���
        if (-not (Get-Command appium -ErrorAction SilentlyContinue)) {
            throw "Appium��װ����֤ʧ�ܣ�������Ҫ�ֶ���װ��"
        }

        Write-Host "Appium��װ�ɹ�" -ForegroundColor Green
    }
    catch{
        Write-Host "Appium��װʧ��: $_" -ForegroundColor Red
        exit 1
    }
}

# 5. ��װ Appium UiAutomator2 ����
function Install-UiAutomator2(){
    try{
        appium driver install uiautomator2

        # ������֤��װ���
        $drivers = appium driver list --installed 2>&1
        if ($drivers -match "uiautomator2") {
            Write-Host "UiAutomator2 ������װ�ɹ�" -ForegroundColor Green
        }
        else{
            throw "UiAutomator2 ������װʧ��, ���������иýű�"
        }
    }
    catch{
        Write-Host "UiAutomator2 ������װʧ��: $_" -ForegroundColor Red
    }
}

# 5. ��װ Python3
function Install-Python(){
    param (
        [string]$Version = "3"
    )

    try {
        Write-Host "�������� Python$Version ��װ����..." -ForegroundColor Cyan
        $url = "https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe"
        $installer = "$env:TEMP\python_installer.exe"

        Invoke-WebRequest -Uri $url -OutFile $installer -ErrorAction Stop -UseBasicParsing
        Write-Host "���ھ�Ĭ��װ Python..." -ForegroundColor Cyan
        Start-Process $installer -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait -NoNewWindow

        $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')
    
        # ������֤��װ���
        if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
            throw "Python ��װ����֤ʧ�ܣ�������Ҫ�ֶ������ն˻���Python������������"
        }
        Write-Host "Python $Version ��װ�ɹ�" -ForegroundColor Green
        Remove-Item $installer -Force
    }
    catch {
        Write-Host "Python $Version ��װʧ�ܣ�$_" -ForegroundColor Red
        exit 1
    }
}

# 1. Node.js �������
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host "��⵽ Node.js δ��װ����ʼ�Զ���װ..." -ForegroundColor Yellow
    Install-NodeJS
}
else {
    Write-Host "��⵽ Node.js �Ѱ�װ: $(node -v)" -ForegroundColor Green
}

# 2. OpenJDK �������
$javaCheck = @(
    (Get-Command java -ErrorAction SilentlyContinue),
    (Test-Path "C:\DevTools\Java\jdk*"),
    [Environment]::GetEnvironmentVariable("JAVA_HOME", "Machine")
) | Where-Object { $_ -ne $null }

if (-not $javaCheck) {
    Write-Host "��⵽ JDK δ��װ����ʼ�Զ���װ..." -ForegroundColor Yellow
    Install-JDK -Version 17
}
else {
    Write-Host "��⵽ JDK �Ѱ�װ: $((Get-Command java).Source)" -ForegroundColor Green
}

# 3. ADB������֤
if (-not (Get-Command adb -ErrorAction SilentlyContinue)) {
    Write-Host "��⵽ ADB ����δ���ã���ʼ�Զ���װ..." -ForegroundColor Yellow
    Install-AndroidSDK
}
else {
    Write-Host "��⵽ ADB ����������: $((Get-Command adb).Source)" -ForegroundColor Green
}

# 4. Appium ��ⰲװ
$globalPackages = npm list -g --json | ConvertFrom-Json
$hasAppium = $globalPackages.dependencies.PSObject.Properties.Name -contains "appium"
if ($hasAppium) {
    Write-Host "��⵽ Appium ��ȫ�ְ�װ���汾��" $globalPackages.dependencies.appium.version -ForegroundColor Green
} else {
    Write-Host "��⵽ Appium δȫ�ְ�װ����ʼ�Զ���װ..." -ForegroundColor Red
    Install-Appium
}

# 5. appium driver: uiautomator2 ��ⰲװ
try {
    $drivers = appium driver list --installed 2>&1
    if ($drivers -match "uiautomator2") {
        Write-Host "��⵽ UiAutomator2 �����Ѱ�װ��" -ForegroundColor Green
    } else {
        Write-Host "��⵽ UiAutomator2 ����δ��װ��" -ForegroundColor Red
        Install-UiAutomator2
    }
} catch {
    Write-Host "���ʧ�ܣ���ȷ�� Appium ����ȷ��װ��������á�" -ForegroundColor Yellow
}

# 6. Python�ͻ��˰�װ
if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
    Write-Host "��⵽ Python δ��װ����ʼ�Զ���װ..." -ForegroundColor Yellow
    Install-Python -Version 3
    pip install Appium-Python-Client -i https://pypi.tuna.tsinghua.edu.cn/simple
    Write-Host "Appium-Python-Client ģ�鰲װ�ɹ�" -ForegroundColor Green
}else{
    $packages = pip list 2>&1
    if ($packages -match "Appium-Python-Client") {
        Write-Host "��⵽ Appium-Python-Client ģ���Ѱ�װ��" -ForegroundColor Green
    } else {
        Write-Host "��⵽ Appium-Python-Client ģ��δ��װ����ʼ�Զ���װ..." -ForegroundColor Yellow
        pip install Appium-Python-Client -i https://pypi.tuna.tsinghua.edu.cn/simple
        Write-Host "Appium-Python-Client ģ�鰲װ�ɹ�" -ForegroundColor Green
    }
}

# ������֤
Write-Host "`n������֤һ����" -ForegroundColor Cyan
Write-Host "Java ����: $((Get-Command java).Source)" -ForegroundColor Green
Write-Host "Node ����: $((Get-Command node).Source)" -ForegroundColor Green
Write-Host "Appium ����: $((Get-Command appium).Source)" -ForegroundColor Green
Write-Host "Python ����: $((Get-Command python).Source)" -ForegroundColor Green

# �ȴ��˳�
Read-Host -Prompt "�� Enter ������..."