$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Tablero ISSSTE · Instalación de entorno" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# ─── 1. Detectar Python disponible ─────────────────────────────────────
$pyExe = $null
$pyArgs = @()

foreach ($v in @("3.14", "3.13", "3.12", "3.11", "3.10")) {
    try {
        & py -$v --version 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) {
            $pyExe = "py"
            $pyArgs = @("-$v")
            Write-Host "[OK] Python $v detectado" -ForegroundColor Green
            break
        }
    } catch { }
}

if (-not $pyExe) {
    try {
        & python --version 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) {
            $pyExe = "python"
            Write-Host "[OK] Python del PATH detectado" -ForegroundColor Green
        }
    } catch { }
}

if (-not $pyExe) {
    Write-Host ""
    Write-Host "[ERROR] No se encontró Python." -ForegroundColor Red
    Write-Host "Instale Python 3.12 o superior desde https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "Marque 'Add python.exe to PATH' durante la instalación." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

# ─── 2. Crear entorno virtual ─────────────────────────────────────────
if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    Write-Host "[*] Creando entorno virtual .venv ..." -ForegroundColor Yellow
    & $pyExe @pyArgs -m venv .venv
} else {
    Write-Host "[OK] Entorno virtual .venv ya existe" -ForegroundColor Green
}

# ─── 3. Actualizar pip ────────────────────────────────────────────────
Write-Host "[*] Actualizando pip ..." -ForegroundColor Yellow
& .\.venv\Scripts\python.exe -m pip install --upgrade pip --quiet

# ─── 4. Instalar dependencias ─────────────────────────────────────────
if (-not (Test-Path ".\requirements.txt")) {
    Write-Host "[ERROR] No se encontró requirements.txt" -ForegroundColor Red
    exit 1
}

Write-Host "[*] Instalando dependencias (esto puede tardar unos minutos) ..." -ForegroundColor Yellow
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt

# ─── 5. Verificación final ────────────────────────────────────────────
Write-Host ""
Write-Host "[*] Verificando instalación ..." -ForegroundColor Yellow
& .\.venv\Scripts\python.exe -c @"
import sys
print('Python:', sys.version.split()[0])
import streamlit, pandas, plotly, openpyxl, numpy
print('streamlit:', streamlit.__version__)
print('pandas   :', pandas.__version__)
print('plotly   :', plotly.__version__)
print('openpyxl :', openpyxl.__version__)
print('numpy    :', numpy.__version__)
print('OK - Todas las dependencias están disponibles')
"@

Write-Host ""
Write-Host "[LISTO] Instalación completa." -ForegroundColor Green
Write-Host "Ejecute ahora:  .\iniciar.ps1" -ForegroundColor Cyan
Write-Host ""