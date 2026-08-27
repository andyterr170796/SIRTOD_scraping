#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
SIRTOD - INEI Data Extractor & Regional Panel Generator (Excel Long Format)
Extractor de Datos Panel Regional del Sistema de Información Regional para la Toma de Decisiones (SIRTOD - INEI)

Formato de Salida Único: Excel (.xlsx) en Formato Panel Long (Tidy Data)
"""

import sys
import os
import time
import json
import argparse
import logging
import requests
from typing import List, Dict, Any
import pandas as pd
import numpy as np
from tqdm import tqdm

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger('SIRTODExtractor')

BASE_URL = 'https://systems.inei.gob.pe/SIRTOD/app/consulta/'

DEFAULT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8'
}

STANDARD_DEPARTMENTS = {
    '0': {'ubigeo': '00', 'nombre': 'NACIONAL'},
    '1': {'ubigeo': '01', 'nombre': 'AMAZONAS'},
    '2': {'ubigeo': '02', 'nombre': 'ÁNCASH'},
    '3': {'ubigeo': '03', 'nombre': 'APURÍMAC'},
    '4': {'ubigeo': '04', 'nombre': 'AREQUIPA'},
    '5': {'ubigeo': '05', 'nombre': 'AYACUCHO'},
    '6': {'ubigeo': '06', 'nombre': 'CAJAMARCA'},
    '7': {'ubigeo': '07', 'nombre': 'CALLAO'},
    '8': {'ubigeo': '08', 'nombre': 'CUSCO'},
    '9': {'ubigeo': '09', 'nombre': 'HUANCAVELICA'},
    '10': {'ubigeo': '10', 'nombre': 'HUÁNUCO'},
    '11': {'ubigeo': '11', 'nombre': 'ICA'},
    '12': {'ubigeo': '12', 'nombre': 'JUNÍN'},
    '13': {'ubigeo': '13', 'nombre': 'LA LIBERTAD'},
    '14': {'ubigeo': '14', 'nombre': 'LAMBAYEQUE'},
    '15': {'ubigeo': '15', 'nombre': 'LIMA'},
    '16': {'ubigeo': '16', 'nombre': 'LORETO'},
    '17': {'ubigeo': '17', 'nombre': 'MADRE DE DIOS'},
    '18': {'ubigeo': '18', 'nombre': 'MOQUEGUA'},
    '19': {'ubigeo': '19', 'nombre': 'PASCO'},
    '20': {'ubigeo': '20', 'nombre': 'PIURA'},
    '21': {'ubigeo': '21', 'nombre': 'PUNO'},
    '22': {'ubigeo': '22', 'nombre': 'SAN MARTÍN'},
    '23': {'ubigeo': '23', 'nombre': 'TACNA'},
    '24': {'ubigeo': '24', 'nombre': 'TUMBES'},
    '25': {'ubigeo': '25', 'nombre': 'UCAYALI'},
    '26': {'ubigeo': '1501', 'nombre': 'LIMA METROPOLITANA'},
    '27': {'ubigeo': '1500', 'nombre': 'LIMA PROVINCIA / REGIÓN'}
}


class SIRTODExtractor:
    def __init__(self, output_dir: str = '.', batch_size: int = 50, retries: int = 3):
        self.output_dir = output_dir
        self.batch_size = batch_size
        self.retries = retries
        self.session = requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)
        os.makedirs(self.output_dir, exist_ok=True)

    def _post(self, endpoint: str, data: dict = None) -> Any:
        url = BASE_URL + endpoint
        for attempt in range(1, self.retries + 1):
            try:
                response = self.session.post(url, data=data or {}, timeout=30)
                response.raise_for_status()
                return response.json()
            except Exception as e:
                logger.warning(f"Error al consultar {endpoint} (Intento {attempt}/{self.retries}): {e}")
                time.sleep(2 * attempt)
        raise RuntimeError(f"Fallo al consultar {endpoint} tras {self.retries} intentos.")

    def fetch_arbol_tematico(self) -> List[Dict[str, Any]]:
        logger.info("Descargando árbol temático completo de SIRTOD INEI...")
        data = self._post('arboltematico')
        logger.info(f"Árbol temático recuperado: {len(data):,} nodos.")
        return data

    def build_indicator_catalog(self, tree: List[Dict[str, Any]], level: int = 1) -> pd.DataFrame:
        """
        Extrae y reconstruye la taxonomía de TODOS los indicadores regionales disponibles (departamentos == 1).
        """
        logger.info(f"Procesando y filtrando catálogo de indicadores (Nivel geográfico {level})...")
        node_map = {item['idTema']: item for item in tree if 'idTema' in item}
        
        indicators = []
        for item in tree:
            if item.get('tipo') == 'default' and item.get('codigoIndicador'):
                if level == 1 and item.get('departamentos') != '1':
                    continue
                elif level == 2 and item.get('provincias') != '1':
                    continue
                elif level == 3 and (item.get('distritos1') != '1' and item.get('distritos2') != '1' and item.get('distritos3') != '1'):
                    continue

                path = []
                curr_id = item.get('idPadre')
                while curr_id and curr_id in node_map and curr_id != '1':
                    parent = node_map[curr_id]
                    parent_name = parent.get('nombreTema') or parent.get('nombreTemaIngles')
                    if parent_name:
                        path.append(parent_name.strip())
                    curr_id = parent.get('idPadre')
                
                path.reverse()
                cat_n1 = path[0] if len(path) > 0 else 'GENERAL'
                cat_n2 = path[1] if len(path) > 1 else (path[0] if len(path) > 0 else 'GENERAL')
                cat_full = ' > '.join(path)

                indicators.append({
                    'id_indicador': str(item['codigoIndicador']).strip(),
                    'nombre_indicador': (item.get('nombreIndicador') or item.get('nombreTema', '')).strip(),
                    'categoria_principal': cat_n1,
                    'subcategoria': cat_n2,
                    'ruta_categoria': cat_full
                })

        df_cat = pd.DataFrame(indicators)
        df_cat = df_cat.drop_duplicates(subset=['id_indicador']).reset_index(drop=True)
        logger.info(f"Catálogo completo de variables: {len(df_cat):,} indicadores regionales únicos.")
        return df_cat

    def fetch_panel_data(self, indicator_ids: List[str], level: int = 1, from_year: str = '1940', to_year: str = '2026') -> pd.DataFrame:
        """
        Extrae datos anuales por lotes para la lista de indicadores solicitada.
        """
        logger.info(f"Iniciando extracción masiva de datos panel para {len(indicator_ids):,} indicadores...")
        chunks = [indicator_ids[i:i + self.batch_size] for i in range(0, len(indicator_ids), self.batch_size)]
        
        all_rows = []
        for chunk in tqdm(chunks, desc="Descargando lotes de datos"):
            batch_str = ','.join(chunk)
            payload = {
                'indicador_listado': batch_str,
                'tipo_ubigeo': str(level),
                'anio_desde': str(from_year),
                'anio_hasta': str(to_year),
                'ubigeo_listado': ''
            }
            try:
                res = self._post('dato_anual', payload)
                if isinstance(res, list):
                    all_rows.extend(res)
            except Exception as e:
                logger.error(f"Error en lote ({chunk[0]}..{chunk[-1]}): {e}")

        logger.info(f"Descarga finalizada. Registros brutos recuperados: {len(all_rows):,}")
        if not all_rows:
            return pd.DataFrame()

        return pd.DataFrame(all_rows)

    def clean_and_format_panel(self, df_raw: pd.DataFrame, df_cat: pd.DataFrame) -> pd.DataFrame:
        """
        Normaliza y estructura los datos a formato Panel Long (Tidy Data).
        """
        logger.info("Transformando y estructurando dataset a Panel Long (Tidy Data)...")
        df = df_raw.copy()
        
        df['id_indicador'] = df['indicadorId'].astype(str).str.strip()
        df['anio'] = pd.to_numeric(df['anioIntervalo'], errors='coerce')
        df['id_dep'] = df['departamentoId'].astype(str).str.strip()
        
        def parse_val(v):
            if pd.isna(v) or v is None:
                return np.nan
            v_str = str(v).strip().replace(' ', '').replace(',', '.')
            try:
                return float(v_str)
            except ValueError:
                return np.nan

        df['valor'] = df['datoAnual'].apply(parse_val)
        df = df.dropna(subset=['anio', 'valor']).copy()
        df['anio'] = df['anio'].astype(int)

        def get_ubigeo(did):
            return STANDARD_DEPARTMENTS.get(did, {}).get('ubigeo', str(did).zfill(2))

        def get_dep_name(did):
            return STANDARD_DEPARTMENTS.get(did, {}).get('nombre', f'DEPARTAMENTO_{did}')

        df['ubigeo'] = df['id_dep'].apply(get_ubigeo)
        df['departamento'] = df['id_dep'].apply(get_dep_name)

        df_merged = df.merge(
            df_cat[['id_indicador', 'nombre_indicador', 'categoria_principal', 'subcategoria']],
            on='id_indicador',
            how='left'
        )

        final_cols = ['ubigeo', 'departamento', 'id_indicador', 'nombre_indicador', 'categoria_principal', 'subcategoria', 'anio', 'valor']
        df_final = df_merged[final_cols].sort_values(by=['id_indicador', 'ubigeo', 'anio']).reset_index(drop=True)
        logger.info(f"Dataset Panel Long final: {len(df_final):,} observaciones completas.")
        return df_final

    def save_excel_output(self, df_long: pd.DataFrame, chunk_size: int = 700000):
        """
        Guarda el dataset Panel Long en Excel (.xlsx).
        Divide automáticamente en pestañas ('Parte_1', 'Parte_2') si supera los 700,000 registros
        para respetar el límite de 1,048,576 filas de Excel y evitar truncamientos de datos.
        """
        output_excel = os.path.join(self.output_dir, 'sirtod_panel_regional_long.xlsx')
        logger.info(f"Generando archivo Excel (.xlsx) único: {os.path.abspath(output_excel)}")
        
        total_rows = len(df_long)
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            if total_rows <= chunk_size:
                df_long.to_excel(writer, sheet_name='Datos_Panel_Long', index=False)
            else:
                for i in range(0, total_rows, chunk_size):
                    part_num = (i // chunk_size) + 1
                    sheet_name = f"Parte_{part_num}"
                    sub_df = df_long.iloc[i:i + chunk_size]
                    logger.info(f"  Escribiendo pestaña '{sheet_name}': filas {i:,} a {i + len(sub_df):,}...")
                    sub_df.to_excel(writer, sheet_name=sheet_name, index=False)

        size_mb = os.path.getsize(output_excel) / (1024 * 1024)
        logger.info(f"  [OK] Archivo Excel generado con éxito ({size_mb:.2f} MB): {output_excel}")


def main():
    parser = argparse.ArgumentParser(description='SIRTOD INEI Regional Panel Extractor (Excel Long Only)')
    parser.add_argument('--output-dir', type=str, default='.', help='Directorio de salida para el archivo Excel')
    parser.add_argument('--level', type=int, default=1, choices=[1, 2, 3], help='Nivel geográfico: 1=Regional/Dep (default)')
    parser.add_argument('--batch-size', type=int, default=50, help='Tamaño de lote HTTP (default: 50)')
    parser.add_argument('--limit', type=int, default=0, help='Límite opcional de indicadores (0 = todos los ~4,200)')
    
    args = parser.parse_args()

    extractor = SIRTODExtractor(output_dir=args.output_dir, batch_size=args.batch_size)

    # 1. Fetch tree and build catalog
    tree = extractor.fetch_arbol_tematico()
    df_cat = extractor.build_indicator_catalog(tree, level=args.level)

    if args.limit > 0:
        logger.info(f"Aplicando límite opcional: {args.limit} indicadores.")
        df_cat = df_cat.iloc[:args.limit].copy()

    # 2. Extract raw data
    indicator_ids = df_cat['id_indicador'].tolist()
    df_raw = extractor.fetch_panel_data(indicator_ids, level=args.level)

    if df_raw.empty:
        logger.error("No se obtuvieron datos.")
        sys.exit(1)

    # 3. Format to Panel Long
    df_long = extractor.clean_and_format_panel(df_raw, df_cat)

    # 4. Save exclusively to Excel (.xlsx)
    extractor.save_excel_output(df_long)

    logger.info("¡PROCESO COMPLETADO! Archivo Excel Panel Long listo para análisis.")


if __name__ == '__main__':
    main()
