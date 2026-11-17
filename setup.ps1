param(
    [switch]$RunApp
)

$ErrorActionPreference = "Stop"

$pythonVersion = "3.13.14"
$runtimeDirectory = Join-Path $PSScriptRoot ".python"
$pythonExecutable = Join-Path $runtimeDirectory "python.exe"
$requirementsFile = Join-Path $PSScriptRoot "requirements.txt"
$mainFile = Join-Path $PSScriptRoot "main.py"

if (-not (Test-Path $requirementsFile)) {
    throw "Cannot find requirements.txt in $PSScriptRoot."
}
if (-not (Test-Path $mainFile)) {
    throw "Cannot find main.py in $PSScriptRoot."
}

if (-not (Test-Path $pythonExecutable)) {
    $architecture = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture
    switch ($architecture.ToString()) {
        "X64" { $installerArchitecture = "amd64" }
        "Arm64" { $installerArchitecture = "arm64" }
        default { throw "Unsupported Windows architecture: $architecture. A 64-bit x64 or ARM64 Windows PC is required." }
    }

    $installerName = "python-$pythonVersion-$installerArchitecture.exe"
    $installerUrl = "https://www.python.org/ftp/python/$pythonVersion/$installerName"
    $temporaryInstallerName = "{0}-{1}" -f [guid]::NewGuid().ToString(), $installerName
    $installerPath = Join-Path ([System.IO.Path]::GetTempPath()) $temporaryInstallerName

    Write-Host "Downloading Python $pythonVersion from python.org..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath

    try {
        Write-Host "Installing the project-local Python runtime..."
        $arguments = @(
            "/quiet",
            "InstallAllUsers=0",
            "`"TargetDir=$runtimeDirectory`"",
            "Include_pip=1",
            "Include_tcltk=1",
            "Include_test=0",
            "Include_launcher=0",
            "PrependPath=0",
            "Shortcuts=0",
            "Include_doc=0"
        )
        $process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            throw "Python installer failed with exit code $($process.ExitCode)."
        }
    }
    finally {
        if (Test-Path $installerPath) {
            Remove-Item -LiteralPath $installerPath -Force
        }
    }
}

if (-not (Test-Path $pythonExecutable)) {
    throw "Python setup did not create $pythonExecutable."
}

$installedVersion = & $pythonExecutable -c "import platform; print(platform.python_version())"
if ($LASTEXITCODE -ne 0 -or $installedVersion.Trim() -ne $pythonVersion) {
    throw "Expected Python $pythonVersion at $pythonExecutable, found '$installedVersion'."
}

& $pythonExecutable -c "import tkinter"
if ($LASTEXITCODE -ne 0) {
    throw "The project-local Python runtime is missing Tkinter."
}

Write-Host "Installing project dependencies..."
& $pythonExecutable -m pip install --disable-pip-version-check -r $requirementsFile
if ($LASTEXITCODE -ne 0) {
    throw "Dependency installation failed. Check your internet connection and try again."
}

if ($RunApp) {
    Write-Host "Starting Experimental Plan Builder..."
    & $pythonExecutable $mainFile
    if ($LASTEXITCODE -ne 0) {
        throw "The application exited with code $LASTEXITCODE."
    }
}
