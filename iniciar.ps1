$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    Write-Host ""
    Write-Host "[ERROR] No existe el entorno virtual .venv" -ForegroundColor Red
    Write-Host "Ejecute primero:  .\instalar_dependencias.ps1" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

if (-not (Test-Path ".\data\RESUMEN_ISSSTE.xlsx")) {
    Write-Host ""
    Write-Host "[ERROR] No se encontró el archivo de datos:" -ForegroundColor Red
    Write-Host "         .\data\RESUMEN_ISSSTE.xlsx" -ForegroundColor Yellow
    Write-Host "Coloque el libro en la carpeta data\ y vuelva a intentar." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Iniciando tablero ISSSTE..." -ForegroundColor Cyan
Write-Host " Abra http://localhost:8501 en su navegador" -ForegroundColor Cyan
Write-Host " Presione Ctrl+C para detener" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

& .\.venv\Scripts\python.exe -m streamlit run app.py