Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
$PythonExe = Join-Path $RepoRoot ".venv/Scripts/python.exe"
$BackendRoot = Join-Path $RepoRoot "backend"
$FrontendRoot = Join-Path $RepoRoot "frontend"

function Get-BackendPython {
    if (-not (Test-Path $PythonExe)) {
        throw "Python executable not found at $PythonExe. Create the virtual environment first."
    }

    return $PythonExe
}

function Install-BackendDependencies {
    $python = Get-BackendPython
    & $python -m pip install --upgrade pip
    & $python -m pip install -r (Join-Path $BackendRoot "requirements.txt")
}

function Invoke-BackendMigrations {
    $python = Get-BackendPython
    & $python (Join-Path $BackendRoot "manage.py") migrate
}

function New-BackendSuperUser {
    $python = Get-BackendPython
    & $python (Join-Path $BackendRoot "manage.py") createsuperuser
}

function Start-BackendServer {
    $python = Get-BackendPython
    & $python (Join-Path $BackendRoot "manage.py") runserver
}

function Start-FrontendServer {
    Push-Location $FrontendRoot
    try {
        npm run dev
    }
    finally {
        Pop-Location
    }
}

function Invoke-BackendTests {
    $python = Get-BackendPython
    Push-Location $BackendRoot
    try {
        & $python -m pytest -q
    }
    finally {
        Pop-Location
    }
}

Write-Host "Loaded helper commands: Install-BackendDependencies, Invoke-BackendMigrations, New-BackendSuperUser, Start-BackendServer, Start-FrontendServer, Invoke-BackendTests"
