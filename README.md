# Tablero ISSSTE
1. Copie el contenido a `D:\DAERAS\Dashboard\Tableros_py`.
2. En PowerShell: `Set-ExecutionPolicy -Scope Process Bypass`.
3. Ejecute `.\instalar_dependencias.ps1`.
4. Ejecute `.\iniciar.ps1`.
5. Abra `http://localhost:8501`.

Incluye las seis páginas solicitadas. Usa las familias reales de hojas; para Plazas el libro contiene `ISSSTE_TRAB_NPLAZAS` y `ISSSTE_PLAZ_NOMBRA_PROM` (no existe un prefijo literal `ISSSTE_TRAB_PLAZ`). `TABLAS_DIM` aporta etiquetas y coordenadas. Se usa `scatter_geo`, no `scatter_mapbox`. El nombre correcto del libro es `.xlsx`.

# Tablero ejecutivo ISSSTE · Streamlit

Dashboard interactivo de estadísticas ISSSTE (Trabajadores, Plazas,
Pensionados, Deudos y Familiares) basado en `RESUMEN_ISSSTE.xlsx`.

## 🖥️ Requisitos

- Windows 10/11
- Python 3.12 o superior (probado en 3.14)
- PowerShell 5.1 o superior
- Navegador moderno (Edge, Chrome o Firefox)

## 📁 Estructura del proyecto
