$ErrorActionPreference = "Continue"
Set-Location $PSScriptRoot

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Diagnóstico del entorno · Tablero ISSSTE" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

$todo_ok = $true

# 1. Python en .venv
if (Test-Path ".\.venv\Scripts\python.exe") {
    $v = & .\.venv\Scripts\python.exe --version
    Write-Host "[OK] Entorno virtual .venv ($v)" -ForegroundColor Green
} else {
    Write-Host "[X] Falta el entorno virtual .venv" -ForegroundColor Red
    Write-Host "    -> Ejecute .\instalar_dependencias.ps1" -ForegroundColor Yellow
    $todo_ok = $false
}

# 2. Dependencias
if (Test-Path ".\.venv\Scripts\python.exe") {
    try {
        & .\.venv\Scripts\python.exe -c "import streamlit, pandas, plotly, openpyxl" 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Dependencias instaladas" -ForegroundColor Green
        } else {
            Write-Host "[X] Faltan dependencias" -ForegroundColor Red
            Write-Host "    -> Ejecute .\instalar_dependencias.ps1" -ForegroundColor Yellow
            $todo_ok = $false
        }
    } catch {
        Write-Host "[X] Error al verificar dependencias" -ForegroundColor Red
        $todo_ok = $false
    }
}

# 3. Archivo de datos
if (Test-Path ".\data\RESUMEN_ISSSTE.xlsx") {
    $sz = [math]::Round((Get-Item ".\data\RESUMEN_ISSSTE.xlsx").Length / 1MB, 2)
    Write-Host "[OK] RESUMEN_ISSSTE.xlsx ($sz MB)" -ForegroundColor Green
} else {
    Write-Host "[X] Falta .\data\RESUMEN_ISSSTE.xlsx" -ForegroundColor Red
    $todo_ok = $false
}

# 4. app.py
if (Test-Path ".\app.py") {
    Write-Host "[OK] app.py presente" -ForegroundColor Green
} else {
    Write-Host "[X] Falta app.py" -ForegroundColor Red
    $todo_ok = $false
}

# 5. Puertos ocupados
$puerto = Get-NetTCPConnection -LocalPort 8501 -ErrorAction SilentlyContinue
if ($puerto) {
    Write-Host "[!] El puerto 8501 ya está en uso" -ForegroundColor Yellow
    Write-Host "    Cierre el proceso o use otro puerto" -ForegroundColor Yellow
} else {
    Write-Host "[OK] Puerto 8501 libre" -ForegroundColor Green
}

Write-Host ""
if ($todo_ok) {
    Write-Host "[LISTO] Todo correcto. Ejecute .\iniciar.ps1" -ForegroundColor Green
} else {
    Write-Host "[ATENCIÓN] Corrija los errores marcados con [X]" -ForegroundColor Yellow
}
Write-Host ""