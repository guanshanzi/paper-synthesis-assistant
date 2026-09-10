param(
    [switch]$Repair
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$RuntimeDir = Join-Path $Root ".runtime\python"
$VenvDir = Join-Path $Root ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$ReqFile = Join-Path $Root "requirements.txt"
$StampFile = Join-Path $VenvDir ".requirements.sha256"
$LogDir = Join-Path $Root "bootstrap_logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogFile = Join-Path $LogDir ("bootstrap_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

function Log([string]$Message) {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Test-Python([string]$Exe) {
    if (-not (Test-Path $Exe)) { return $false }
    try {
        & $Exe -c "import sys, venv, tkinter; raise SystemExit(0 if (3,10) <= sys.version_info[:2] < (3,14) else 1)" *> $null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Resolve-SystemPython {
    $candidates = New-Object System.Collections.Generic.List[string]

    $pythonCmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pythonCmd -and $pythonCmd.Source) {
        try {
            $exe = (& $pythonCmd.Source -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1).Trim()
            if ($exe) { $candidates.Add($exe) }
        } catch {}
    }

    $pyCmd = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyCmd -and $pyCmd.Source) {
        try {
            $exe = (& $pyCmd.Source -3 -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1).Trim()
            if ($exe) { $candidates.Add($exe) }
        } catch {}
    }

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (Test-Python $candidate) { return $candidate }
    }
    return $null
}

function Install-LocalPython {
    $version = "3.12.10"
    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "amd64" }
    $installerName = "python-$version-$arch.exe"
    $url = "https://www.python.org/ftp/python/$version/$installerName"
    $installer = Join-Path $env:TEMP $installerName

    Write-Host ""
    Write-Host "首次运行需要 Python 3.12。" -ForegroundColor Yellow
    Write-Host "程序可以从 Python 官方网站自动下载，并安装到本软件目录：" -ForegroundColor Yellow
    Write-Host "  $RuntimeDir"
    Write-Host "不会加入系统 PATH，不需要管理员权限，也不会修改你现有的 Python。"
    Write-Host ""
    $answer = Read-Host "输入 Y 同意自动下载和安装 Python；输入其他内容取消"
    if ($answer -notmatch '^[Yy]$') {
        throw "用户取消了 Python 自动安装。"
    }

    Log "正在从 Python.org 下载 Python $version ($arch)..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing

    New-Item -ItemType Directory -Force -Path $RuntimeDir | Out-Null
    $args = @(
        "/quiet",
        "InstallAllUsers=0",
        ('TargetDir="{0}"' -f $RuntimeDir),
        "PrependPath=0",
        "Include_launcher=0",
        "Include_pip=1",
        "Include_tcltk=1",
        "Include_test=0",
        "Include_doc=0",
        "Shortcuts=0",
        "AssociateFiles=0"
    )

    Log "正在把 Python 安装到软件目录..."
    $proc = Start-Process -FilePath $installer -ArgumentList $args -Wait -PassThru
    Remove-Item $installer -Force -ErrorAction SilentlyContinue

    if ($proc.ExitCode -notin @(0, 3010)) {
        throw "Python 安装失败，退出码：$($proc.ExitCode)"
    }

    $localPython = Join-Path $RuntimeDir "python.exe"
    if (-not (Test-Python $localPython)) {
        throw "Python 安装完成，但运行环境自检失败。请查看 $LogFile"
    }
    return $localPython
}

try {
    Log "论文自动综合助手 v1.0.0 启动。"

    foreach ($dir in @("pdfs", "output", "logs", "cards_input", "archive")) {
        New-Item -ItemType Directory -Force -Path (Join-Path $Root $dir) | Out-Null
    }

    if ($Repair -and (Test-Path $VenvDir)) {
        Log "修复模式：删除旧的 .venv 后重新创建。"
        Remove-Item -Recurse -Force $VenvDir
    }

    $basePython = $null
    $localPython = Join-Path $RuntimeDir "python.exe"
    if (Test-Python $localPython) {
        $basePython = $localPython
        Log "使用软件自带的本地 Python。"
    } else {
        $basePython = Resolve-SystemPython
        if ($basePython) {
            Log "检测到可用的系统 Python：$basePython"
        } else {
            $basePython = Install-LocalPython
            Log "本地 Python 安装完成。"
        }
    }

    if (-not (Test-Path $VenvPython)) {
        Log "首次配置：正在创建独立虚拟环境 .venv ..."
        & $basePython -m venv $VenvDir 2>&1 | Tee-Object -FilePath $LogFile -Append
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $VenvPython)) {
            throw "创建虚拟环境失败。"
        }
    }

    $currentHash = (Get-FileHash $ReqFile -Algorithm SHA256).Hash
    $savedHash = if (Test-Path $StampFile) { (Get-Content $StampFile -Raw).Trim() } else { "" }

    if ($Repair -or $currentHash -ne $savedHash) {
        Log "正在安装/更新运行依赖，首次运行可能需要几分钟..."
        $env:PIP_DISABLE_PIP_VERSION_CHECK = "1"
        & $VenvPython -m pip install --upgrade pip 2>&1 | Tee-Object -FilePath $LogFile -Append
        if ($LASTEXITCODE -ne 0) { throw "pip 更新失败。" }

        & $VenvPython -m pip install -r $ReqFile 2>&1 | Tee-Object -FilePath $LogFile -Append
        if ($LASTEXITCODE -ne 0) { throw "Python 依赖安装失败。" }

        Set-Content -Path $StampFile -Value $currentHash -Encoding ASCII
    } else {
        Log "依赖环境已就绪，无需重复安装。"
    }

    Log "正在执行运行环境自检..."
    & $VenvPython -c "import tkinter, fitz, docx, openai, pandas, openpyxl; print('environment-ok')" 2>&1 | Tee-Object -FilePath $LogFile -Append
    if ($LASTEXITCODE -ne 0) {
        throw "运行环境自检失败。可双击“REPAIR_ENV.bat”重建依赖。"
    }

    Log "环境正常，正在启动论文自动综合助手。"
    $app = Join-Path $Root "app.py"
    Start-Process -FilePath $VenvPython -ArgumentList ('"{0}"' -f $app) -WorkingDirectory $Root
    exit 0
}
catch {
    Log ("启动失败：" + $_.Exception.Message)
    Write-Host ""
    Write-Host "启动失败。详细日志：$LogFile" -ForegroundColor Red
    Write-Host "可以先双击“REPAIR_ENV.bat”重试。"
    Read-Host "按 Enter 关闭"
    exit 1
}
