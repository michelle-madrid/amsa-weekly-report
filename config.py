"""Configuración y constantes del generador de informe semanal."""

import os
from pathlib import Path

# Ignora warnings de formato condicional de openpyxl que no afectan el flujo.
import warnings
warnings.filterwarnings("ignore", message="Conditional Formatting extension is not supported")

# Activa o desactiva el modo debug para usar rutas fijas o seleccionar archivos manualmente.
MODO_DEBUG = True

# Define un preset de rutas de debug para la semana 13.
RUTAS_DEBUG_SEMANA_13 = {
  "excel_madre": r"C:\Users\amsmmadrida\Downloads\13_Semana- 27 mar al 01 abr\Semana 13 -  27 mar al 01 abr.xlsx",
  "excel_indicadores": r"C:\Users\amsmmadrida\Downloads\13_Semana- 27 mar al 01 abr\06 -SSO\BDatos_01_abril_2026.xlsx",
  "carpeta_destino": r"C:\Users\amsmmadrida\Downloads\13_Semana- 27 mar al 01 abr",
  "nombre_archivo": r"Prueba_Informe_Refactorizado",
  "informes": {
    "MLP": r"C:\Users\amsmmadrida\Downloads\13_Semana- 27 mar al 01 abr\01 -MLP\Informe Semanal 26-0401.docx",
    "CEN": r"C:\Users\amsmmadrida\Downloads\13_Semana- 27 mar al 01 abr\02 -CEN\20260327 al 20260401 CEN_Informe Semanal.docx",
    "ANT": r"C:\Users\amsmmadrida\Downloads\13_Semana- 27 mar al 01 abr\03 -ANT\20260401 Informe Semanal.docx",
    "CMZ": r"C:\Users\amsmmadrida\Downloads\13_Semana- 27 mar al 01 abr\04 -CMZ\CMZ 2026 - 26 04 01 - 01 proyectado.docx",
    "FCAB": r"C:\Users\amsmmadrida\Downloads\13_Semana- 27 mar al 01 abr\05 -FCAB\AMSA_SEM 13 DEL 27 DE MARZO AL 01 DE ABRIL2026.docx",
  }
}

# Define un preset de rutas de debug para la semana 12.
RUTAS_DEBUG_SEMANA_12 = {
  "excel_madre": r"C:\Users\amsmmadrida\Downloads\12_Semana- 20 mar al 26 mar\Semana 12 -  20 mar al 26 mar.xlsx",
  "excel_indicadores": r"C:\Users\amsmmadrida\Downloads\12_Semana- 20 mar al 26 mar\06 -SSO\BDatos_26_marzo_2026.xlsx.xlsx",
  "carpeta_destino": r"C:\Users\amsmmadrida\Downloads\12_Semana- 20 mar al 26 mar",
  "nombre_archivo": r"Prueba_Informe_Refactorizado",
  "informes": {
    "MLP": r"C:\Users\amsmmadrida\Downloads\12_Semana- 20 mar al 26 mar\01 -MLP\Informe Semanal 26-0326.docx",
    "CEN": r"C:\Users\amsmmadrida\Downloads\12_Semana- 20 mar al 26 mar\02 -CEN\20260320 al 20260326 CEN_Informe Semanal.docx",
    "ANT": r"C:\Users\amsmmadrida\Downloads\12_Semana- 20 mar al 26 mar\03 -ANT\20260326 Informe Semanal.docx",
    "CMZ": r"C:\Users\amsmmadrida\Downloads\12_Semana- 20 mar al 26 mar\04 -CMZ\CMZ 2026 - 26 03 26 - 26 proyectado.docx",
    "FCAB": r"C:\Users\amsmmadrida\Downloads\12_Semana- 20 mar al 26 mar\05 -FCAB\AMSA_SEM 12 DEL 20 AL 26 DE MARZO 2026.docx",
  }
}

# Selecciona el preset activo de rutas debug.
RUTAS_DEBUG = RUTAS_DEBUG_SEMANA_12

# Controla si se incluye la página de estado de fases de desarrollo.
INCLUIR_ESTADO_FASES_DESARROLLO = False

# Define el orden oficial de las faenas dentro del informe.
ORDEN_OFICIAL = ["MLP", "CEN", "ANT", "CMZ", "FCAB"]

# Define la configuración base por compañía para exportar sus tablas.
CONFIG_COMPANIAS = {
    "MLP": {"nombre": "Los Pelambres", "rango": "B3:AD33", "alto": 7.69},
    "ANT": {"nombre": "Antucoya", "rango": "A3:AC45", "alto": 10.13},
    "CEN": {"nombre": "Centinela", "rango": "A3:AC85", "alto": 21.41},
    "CMZ": {"nombre": "Zaldívar", "rango": "A3:AC35", "alto": 6.85},
    "FCAB": {"nombre": "FCAB", "rango": "A3:V19", "alto": 3.21},
}

# Define el orden esperado de subtítulos para las principales desviaciones por compañía.
ORDEN_PRINCIPALES_DESVIACIONES = {
    "MLP": {
        "Principales Desviaciones": ["?"],
        "Mina": ["Movimiento Mina", "Total Extracción", "Extracción", "Remanejo"],
        "Concentradora": [""],
        "Planta Desaladora": ["?"],
        "Gestión Hídrica": [""],
    },
    "CEN": {
        "Principales Desviaciones": ["?"],
        "Mina": [
            "Movimiento Mina",
            "Movimiento en Rajo Tesoro",
            "Movimiento en Rajo Esperanza",
            "Movimiento en Rajo Óxido Encuentro",
            "Movimiento en Rajo Esperanza Sur:",
            "Movimiento en Rajo Encuentro Sulfuros",
        ],
        "Sulfuros": [""],
        "Cátodos": ["Planta Hidro MET", "Planta Hidro OXE"],
    },
    "ANT": {
    "Principales Desviaciones": ["?"],
    "Mina": [
      "Movimiento Mina",
      "Extracción Mina",
      "Extracción de Mineral",
      "Extracción de lastre",
      "Remanejo",
      "Extracción a desarrollo",
    ],
    "Planta": [""],
    },
    "CMZ": {
        "Principales Desviaciones": ["?"],
        "Mina": ["Movimiento Mina", "Extracción", "Extracción Mineral", "Extracción Lastre", "Remanejo"],
        "Planta": [""],
    },
    "FCAB": {
        "Principales Desviaciones": ["?"],
        "Tren": ["#Transporte Total de Tren", "Transporte de ácido", "Transporte de Cobre", "Transporte de Concentrados"],
        "Camión": ["Transporte Total de Camión"],
    },
}

NIVEL_BASE_POR_SECCION = {
    "Principales Desviaciones": 2,
    "Mina": 2,
    "Detalle por fases": 2,
    "Planta": 1,
    "Sulfuros": 1,
    "Cátodos": 1,
    "Concentradora": 1,
    "Planta Desaladora": 2,
    "Gestión Hídrica": 1,
    "Tren": 2,
    "Camión": 2,
}

NIVEL_POR_COMPANIA_SECCION_SUBTITULO = {
    "MLP": {"Mina": {"Movimiento Mina": 1}},
    "CEN": {"Mina": {"Movimiento Mina": 1}},
    "ANT": {"Mina": {"Movimiento Mina": 1}},
    "CMZ": {"Mina": {"Movimiento Mina": 1}},
}


# Guarda la ruta de la plantilla Word usada para construir el informe final.
BASE_DIR = Path(__file__).resolve().parent
RUTA_PLANTILLA = BASE_DIR / "Template Viñetas Python.docx"

# Guarda el marcador del encabezado de las tablas SSO de respaldo.
SSO_MARCADOR_TABLA = "id del incidente"
