# SIRTOD - INEI Regional Panel Data Scraper & Extractor 🇵🇪

Este repositorio contiene el código y la base de datos completa extraída del **Sistema de Información Regional para la Toma de Decisiones (SIRTOD)** del **Instituto Nacional de Estadística e Informática (INEI)** del Perú.

Mediante ingeniería inversa a los servicios REST/JSON internos de la plataforma SIRTOD, este proyecto permite descargar y transformar en **pocos minutos** todas las series históricas de datos regionales en un formato estandarizado de **panel de datos (Tidy Data)** ideal para investigaciones económicas, sociales y de políticas públicas.

---

## 📊 Descripción del Dataset Incluido

* **Archivo**: [`sirtod_panel_regional_long.xlsx`](sirtod_panel_regional_long.xlsx) (42.7 MB)
* **Formato**: Excel (`.xlsx`) dividida en 2 pestañas ordenadas (`Parte_1` y `Parte_2`) para no exceder el límite de 1,048,576 filas de Excel y garantizar que el 100% de los datos carguen sin truncarse.
* **Total de Observaciones**: **1,343,258 registros panel**
* **Variables Únicas**: **3,830 indicadores regionales**
* **Cobertura Territorial**: **27 unidades geográficas** (25 Departamentos, Lima Metropolitana y Callao)
* **Rango Temporal**: **1940 a 2026**

### Estructura de Columnas
- `ubigeo`: Código estándar de Ubigeo INEI (`01`, `02`, ..., `25`).
- `departamento`: Nombre oficial del departamento.
- `id_indicador`: Código numérico del indicador en SIRTOD.
- `nombre_indicador`: Descripción oficial de la variable.
- `categoria_principal`: Dimensión/Tema (ej. *DEMOGRÁFICO*, *ECONÓMICO*, *SOCIAL*, etc.).
- `subcategoria`: Subtema específico.
- `anio`: Año de la observación.
- `valor`: Valor numérico procesado y limpio.

---

## 🛠️ Instalación y Uso del Extractor

### 1. Requisitos Previos
Tener instalado Python 3.8+ y Git.

```bash
git clone https://github.com/andyterr170796/SIRTOD_scraping.git
cd SIRTOD_scraping
pip install -r requirements.txt
```

### 2. Ejecutar la Extracción
Para actualizar la base de datos o descargar nuevamente todos los datos regionales:

```bash
python sirtod_extractor.py
```

El script consultará automáticamente la API de SIRTOD, descargará los lotes en paralelo y actualizará el archivo `sirtod_panel_regional_long.xlsx`.

---

## 💻 Carga Rápida en Python / Stata / R

### Python (Pandas)
```python
import pandas as pd

# Leer primera pestaña
df_part1 = pd.read_excel('sirtod_panel_regional_long.xlsx', sheet_name='Parte_1')
# Leer segunda pestaña
df_part2 = pd.read_excel('sirtod_panel_regional_long.xlsx', sheet_name='Parte_2')

# Unir dataset completo
df_completo = pd.concat([df_part1, df_part2], ignore_index=True)
print(f"Total registros cargados: {len(df_completo):,}")
```

### R
```R
library(readxl)

part1 <- read_excel("sirtod_panel_regional_long.xlsx", sheet = "Parte_1")
part2 <- read_excel("sirtod_panel_regional_long.xlsx", sheet = "Parte_2")

df_panel <- rbind(part1, part2)
```

---

## 📜 Licencia y Fuentes
Los datos provienen del portal oficial del INEI ([SIRTOD INEI](https://systems.inei.gob.pe/SIRTOD/app/consulta)). Este repositorio tiene fines de investigación pública, académica y científica.
