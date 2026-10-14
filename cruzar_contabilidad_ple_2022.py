"""
================================================================================
CRUCE DE CONTABILIDAD vs PLE DIARIO (SUNAT) - EJERCICIO 2022
================================================================================
Empresa: CORPORACIÓN COMATPE SAC
RUC: 20516403650
Archivos de Entrada:
  - Reporte Contabilidad 1-01-2022  a 31-12-2022 del 12.10.2026.xlsx
  - Mensual ple diario 2022 decl ----- copia.xlsx
Archivo de Salida:
  - Cruce_Contabilidad_vs_PLE_Diario_2022.xlsx

Características Especiales:
  - Full Outer Join 1 a 1 por Voucher, Fecha, Cuenta Contable (sin puntos), Debe y Haber.
  - Asignación de índices de ocurrencia enteros (#0, #1...) para evitar duplicaciones cartesianas.
  - Integración del Asiento de Apertura (Mes 00) dentro de Enero (Mes 01).
  - Hoja 1: RESUMEN_GENERAL (Panel ejecutivo anual y tabla comparativa mes a mes).
  - Hoja 2: RESUMEN_POR_CUENTA (Consolidado anual por cada cuenta contable con diferencias y estado).
  - Hoja 3: DIFERENCIAS_TOTALES (Las 21,589 diferencias anuales para auditoría rápida).
  - Hoja 4: CRUCE_ANUAL (Los 387,638 registros completos del cruce con todas las columnas de ambos sistemas).
  - Hojas 5 a 16: Hojas mensuales (01_ENERO a 12_DICIEMBRE) con mini panel resumen y diferencias al inicio.
================================================================================
"""

import os
import sys
import time
import glob
import pandas as pd
import numpy as np
import xlsxwriter

MONTH_NAMES = {
    '01': '01_ENERO',
    '02': '02_FEBRERO',
    '03': '03_MARZO',
    '04': '04_ABRIL',
    '05': '05_MAYO',
    '06': '06_JUNIO',
    '07': '07_JULIO',
    '08': '08_AGOSTO',
    '09': '09_SETIEMBRE',
    '10': '10_OCTUBRE',
    '11': '11_NOVIEMBRE',
    '12': '12_DICIEMBRE'
}

CLASS_9_DESC = {
    '91': 'COSTO DE VENTAS / OPERACIÓN (CLASE 9)',
    '92': 'COSTOS DE PRODUCCIÓN (CLASE 9)',
    '94': 'GASTOS DE ADMINISTRACIÓN (CLASE 9)',
    '95': 'GASTOS DE VENTAS (CLASE 9)',
    '97': 'GASTOS FINANCIEROS (CLASE 9)'
}

def find_file(directory, patterns):
    for pattern in patterns:
        matches = glob.glob(os.path.join(directory, pattern))
        if matches:
            matches.sort(key=os.path.getsize, reverse=True)
            return matches[0]
    return None

def main():
    start_time = time.time()
    downloads_dir = r"c:\Users\Sistemas\Downloads"
    year = 2022

    print("=" * 80)
    print(f" INICIANDO CRUCE CONTABILIDAD vs PLE DIARIO - EJERCICIO {year}")
    print("=" * 80)

    # 1. Localizar archivos de entrada
    rep_patterns = [
        f"Reporte Contabilidad 1-01-{year}*del 12.10.2026.xlsx",
        f"Reporte Contabilidad*{year}*.xlsx",
        f"Reporte de Contabilidad*{year}*.xlsx"
    ]
    ple_patterns = [
        f"Mensual ple diario {year} decl ----- copia.xlsx",
        f"Mensual ple diario {year} decl*.xlsx",
        f"*ple*diario*{year}*.xlsx"
    ]

    file_reporte = find_file(downloads_dir, rep_patterns)
    file_ple = find_file(downloads_dir, ple_patterns)

    if not file_reporte:
        raise FileNotFoundError(f"No se encontró el Reporte de Contabilidad {year} en '{downloads_dir}'.")
    if not file_ple:
        raise FileNotFoundError(f"No se encontró el PLE Diario {year} en '{downloads_dir}'.")

    output_excel = os.path.join(downloads_dir, f"Cruce_Contabilidad_vs_PLE_Diario_{year}.xlsx")

    print(f"Archivo Reporte Contabilidad : {os.path.basename(file_reporte)} ({os.path.getsize(file_reporte)/(1024*1024):.2f} MB)")
    print(f"Archivo PLE Diario Declarado : {os.path.basename(file_ple)} ({os.path.getsize(file_ple)/(1024*1024):.2f} MB)")
    print(f"Archivo Destino Generado     : {os.path.basename(output_excel)}")

    # 2. Cargar y procesar Reporte de Contabilidad
    print(f"\n[1/5] Cargando '{os.path.basename(file_reporte)}' con motor calamine...")
    t0 = time.time()
    cols_rep = [
        'Ejer.', 'Prdo.', 'Reg. Ctb.', 'Fuente', 'Fec. Mov.', 'Nro. Cta.',
        'Nombre Cuenta', 'Mto. Debe', 'Mto. Haber', 'Glosa', 'Cod. Tip. Doc.', 'Nro. Doc.'
    ]
    df_rep = pd.read_excel(
        file_reporte,
        sheet_name='Contabilidad',
        header=4,
        engine='calamine',
        usecols=cols_rep
    )
    print(f"      -> {len(df_rep):,} filas leídas en {time.time()-t0:.2f}s")

    # Limpieza de registros vacíos
    df_rep = df_rep[df_rep['Reg. Ctb.'].notna() & df_rep['Nro. Cta.'].notna()].copy()

    # Normalización contable
    df_rep['REP_Nro_Cta_Original'] = df_rep['Nro. Cta.'].astype(str).str.strip()
    df_rep['REP_Nro_Cta_Limpia'] = df_rep['Nro. Cta.'].astype(str).str.replace('.', '', regex=False).str.strip()
    df_rep['REP_Reg_Ctb'] = df_rep['Reg. Ctb.'].astype(str).str.strip().str.upper()
    df_rep['REP_Fec_Mov'] = pd.to_datetime(df_rep['Fec. Mov.'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
    df_rep['REP_Prdo'] = df_rep['Prdo.'].fillna(0).astype(int).astype(str)
    df_rep['REP_Nombre_Cuenta'] = df_rep['Nombre Cuenta'].fillna('').astype(str).str.strip()
    df_rep['REP_Glosa'] = df_rep['Glosa'].fillna('').astype(str).str.strip()
    df_rep['REP_Tip_Doc'] = df_rep['Cod. Tip. Doc.'].fillna('').astype(str).str.strip()
    df_rep['REP_Nro_Doc'] = df_rep['Nro. Doc.'].fillna('').astype(str).str.strip()

    df_rep['REP_Debe_Num'] = df_rep['Mto. Debe'].fillna(0).astype(float).round(2)
    df_rep['REP_Haber_Num'] = df_rep['Mto. Haber'].fillna(0).astype(float).round(2)
    df_rep['REP_Debe_Str'] = df_rep['REP_Debe_Num'].apply(lambda x: f"{x:.2f}")
    df_rep['REP_Haber_Str'] = df_rep['REP_Haber_Num'].apply(lambda x: f"{x:.2f}")

    # Clave de cruce: Reg. Ctb. + Fec. Mov. + Nro. Cta. + Mto. Debe + Mto. Haber
    df_rep['CLAVE_CONCAT'] = (
        df_rep['REP_Reg_Ctb'] + "|" +
        df_rep['REP_Fec_Mov'] + "|" +
        df_rep['REP_Nro_Cta_Limpia'] + "|" +
        df_rep['REP_Debe_Str'] + "|" +
        df_rep['REP_Haber_Str']
    )
    df_rep['OCURRENCIA'] = df_rep.groupby('CLAVE_CONCAT').cumcount().fillna(0).astype(int).astype(str)
    df_rep['CLAVE_MATCH'] = df_rep['CLAVE_CONCAT'] + "#" + df_rep['OCURRENCIA']
    print(f"      -> {len(df_rep):,} movimientos contables normalizados.")

    # Diccionario de Cuentas Contables (Código -> Nombre)
    cta_map = df_rep.drop_duplicates('REP_Nro_Cta_Limpia').set_index('REP_Nro_Cta_Limpia')['REP_Nombre_Cuenta'].to_dict()
    print(f"      -> {len(cta_map):,} cuentas contables indexadas.")

    # 3. Cargar y procesar PLE Diario Mensual Declarado
    print(f"\n[2/5] Cargando '{os.path.basename(file_ple)}' con motor calamine...")
    t1 = time.time()
    cols_ple = [
        'MES', 'VOUCHER', 'SECCION', 'CUENTA', 'MONEDA', 'TD', 'SERIE', 'NUMERO',
        'FECHA 1', 'GLOSA', 'DEBE', 'HABER'
    ]
    df_ple = pd.read_excel(
        file_ple,
        sheet_name='2022 ANUAL (3)',
        header=0,
        engine='calamine',
        usecols=cols_ple
    )
    print(f"      -> {len(df_ple):,} filas leídas en {time.time()-t1:.2f}s")

    # Limpieza PLE
    df_ple = df_ple[df_ple['VOUCHER'].notna() & df_ple['CUENTA'].notna()].copy()

    df_ple['PLE_Cuenta'] = df_ple['CUENTA'].astype(str).str.replace('.', '', regex=False).str.split('.').str[0].str.strip()
    df_ple['PLE_Voucher'] = df_ple['VOUCHER'].astype(str).str.strip().str.upper()
    df_ple['PLE_Fecha'] = pd.to_datetime(df_ple['FECHA 1'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
    df_ple['PLE_Mes'] = df_ple['MES'].fillna('').astype(str).str.strip()
    df_ple['PLE_Seccion'] = df_ple['SECCION'].fillna('').astype(str).str.strip()
    df_ple['PLE_Glosa'] = df_ple['GLOSA'].fillna('').astype(str).str.strip()
    df_ple['PLE_TD'] = df_ple['TD'].fillna('').astype(str).str.strip()
    df_ple['PLE_Serie'] = df_ple['SERIE'].fillna('').astype(str).str.strip()
    df_ple['PLE_Numero'] = df_ple['NUMERO'].fillna('').astype(str).str.strip()

    df_ple['PLE_Debe_Num'] = df_ple['DEBE'].fillna(0).astype(float).round(2)
    df_ple['PLE_Haber_Num'] = df_ple['HABER'].fillna(0).astype(float).round(2)
    df_ple['PLE_Debe_Str'] = df_ple['PLE_Debe_Num'].apply(lambda x: f"{x:.2f}")
    df_ple['PLE_Haber_Str'] = df_ple['PLE_Haber_Num'].apply(lambda x: f"{x:.2f}")

    # Clave de cruce en PLE
    df_ple['CLAVE_CONCAT'] = (
        df_ple['PLE_Voucher'] + "|" +
        df_ple['PLE_Fecha'] + "|" +
        df_ple['PLE_Cuenta'] + "|" +
        df_ple['PLE_Debe_Str'] + "|" +
        df_ple['PLE_Haber_Str']
    )
    df_ple['OCURRENCIA'] = df_ple.groupby('CLAVE_CONCAT').cumcount().fillna(0).astype(int).astype(str)
    df_ple['CLAVE_MATCH'] = df_ple['CLAVE_CONCAT'] + "#" + df_ple['OCURRENCIA']
    print(f"      -> {len(df_ple):,} movimientos PLE normalizados.")

    # 4. Full Outer Join
    print(f"\n[3/5] Ejecutando Full Outer Join 1 a 1...")
    t2 = time.time()
    rep_merge_cols = [
        'CLAVE_MATCH', 'CLAVE_CONCAT', 'REP_Reg_Ctb', 'REP_Prdo', 'REP_Fec_Mov',
        'REP_Nro_Cta_Original', 'REP_Nro_Cta_Limpia', 'REP_Nombre_Cuenta',
        'REP_Debe_Num', 'REP_Haber_Num', 'REP_Glosa', 'REP_Tip_Doc', 'REP_Nro_Doc'
    ]
    ple_merge_cols = [
        'CLAVE_MATCH', 'CLAVE_CONCAT', 'PLE_Mes', 'PLE_Voucher', 'PLE_Seccion',
        'PLE_Cuenta', 'PLE_Fecha', 'PLE_Debe_Num', 'PLE_Haber_Num', 'PLE_Glosa',
        'PLE_TD', 'PLE_Serie', 'PLE_Numero'
    ]

    merged = pd.merge(
        df_rep[rep_merge_cols],
        df_ple[ple_merge_cols],
        on='CLAVE_MATCH',
        how='outer',
        suffixes=('_REP', '_PLE'),
        indicator=True
    )
    print(f"      -> Cruce finalizado en {time.time()-t2:.2f}s. Total universo: {len(merged):,} registros.")

    status_map = {
        'both': 'COINCIDE',
        'left_only': 'SOLO EN REPORTE',
        'right_only': 'SOLO EN PLE'
    }
    merged['ESTADO_CRUCE'] = merged['_merge'].map(status_map)

    # Clave concatenada unificada
    merged['CLAVE_CONCATENADA'] = np.where(
        merged['CLAVE_CONCAT_REP'].notna(),
        merged['CLAVE_CONCAT_REP'],
        merged['CLAVE_CONCAT_PLE']
    )

    # 5. Normalización del Mes (Integrando Mes 00 en Mes 01)
    def determine_month(row):
        p = row['REP_Prdo']
        if pd.notna(p) and str(p).strip() != '':
            try:
                val = int(float(p))
                if val <= 1:
                    return '01'
                return f"{val:02d}"
            except Exception:
                pass
        m = row['PLE_Mes']
        if pd.notna(m) and len(str(m).strip()) >= 6:
            try:
                s = str(m).strip()
                val = int(s[4:6])
                if val <= 1:
                    return '01'
                return f"{val:02d}"
            except Exception:
                pass
        return '01'

    merged['MES_COD'] = merged.apply(determine_month, axis=1)

    # Estadísticas globales
    conteo_estado = merged['ESTADO_CRUCE'].value_counts()
    n_coincide = conteo_estado.get('COINCIDE', 0)
    n_solo_rep = conteo_estado.get('SOLO EN REPORTE', 0)
    n_solo_ple = conteo_estado.get('SOLO EN PLE', 0)

    print(f"\n=== RESULTADOS DE AUDITORÍA EJERCICIO {year} ===")
    print(f"  Coincidencias Exactas (Ambos) : {n_coincide:,} ({(n_coincide/len(df_ple))*100:.2f}% de PLE)")
    print(f"  Solo en Reporte Contable      : {n_solo_rep:,}")
    print(f"  Solo en PLE Diario Declarado  : {n_solo_ple:,}")
    print(f"  Total Universo Conciliación   : {len(merged):,}")

    # 6. Preparar campos unificados de visualización
    is_rep = (merged['ESTADO_CRUCE'] == 'SOLO EN REPORTE').to_numpy()
    is_match = (merged['ESTADO_CRUCE'] == 'COINCIDE').to_numpy()

    # Voucher
    v_rep = merged['REP_Reg_Ctb'].fillna('').astype(str).to_numpy()
    v_ple = merged['PLE_Voucher'].fillna('').astype(str).to_numpy()
    u_voucher = np.where(is_rep, v_rep, v_ple)

    # Fecha
    f_rep = merged['REP_Fec_Mov'].fillna('').astype(str).to_numpy()
    f_ple = merged['PLE_Fecha'].fillna('').astype(str).to_numpy()
    u_fecha = np.where(is_rep, f_rep, f_ple)

    # Cuenta
    c_rep = merged['REP_Nro_Cta_Limpia'].fillna('').astype(str).to_numpy()
    c_ple = merged['PLE_Cuenta'].fillna('').astype(str).to_numpy()
    u_cuenta = np.where(is_rep, c_rep, c_ple)

    # Nombre Cuenta
    nom_rep = merged['REP_Nombre_Cuenta'].fillna('').astype(str).to_numpy()
    def get_account_name(c, r_nom):
        if r_nom and r_nom.strip():
            return r_nom.strip()
        if c in cta_map and cta_map[c].strip():
            return cta_map[c].strip()
        prefix2 = c[:2]
        if prefix2 in CLASS_9_DESC:
            return CLASS_9_DESC[prefix2]
        return ""

    u_nom_cuenta = [get_account_name(c, n) for c, n in zip(u_cuenta, nom_rep)]
    u_nom_cuenta = np.array(u_nom_cuenta, dtype=object)

    # Debe y Haber
    d_rep = merged['REP_Debe_Num'].fillna(0.0).to_numpy()
    d_ple = merged['PLE_Debe_Num'].fillna(0.0).to_numpy()
    u_debe = np.where(is_rep, d_rep, d_ple)

    h_rep = merged['REP_Haber_Num'].fillna(0.0).to_numpy()
    h_ple = merged['PLE_Haber_Num'].fillna(0.0).to_numpy()
    u_haber = np.where(is_rep, h_rep, h_ple)

    # Glosa
    g_rep = merged['REP_Glosa'].fillna('').astype(str).to_numpy()
    g_ple = merged['PLE_Glosa'].fillna('').astype(str).to_numpy()
    u_glosa = np.where(is_rep, g_rep, g_ple)

    # Documentos
    td_rep = merged['REP_Tip_Doc'].fillna('').astype(str).to_numpy()
    num_rep = merged['REP_Nro_Doc'].fillna('').astype(str).to_numpy()
    u_doc_rep = [f"{td}-{nr}".strip('-') if td or nr else "" for td, nr in zip(td_rep, num_rep)]
    u_doc_rep = np.array(u_doc_rep, dtype=object)

    td_ple = merged['PLE_TD'].fillna('').astype(str).to_numpy()
    s_ple = merged['PLE_Serie'].fillna('').astype(str).to_numpy()
    n_ple = merged['PLE_Numero'].fillna('').astype(str).to_numpy()
    def format_doc_ple(td, s, n):
        parts = [p for p in [s, n] if p and p != '-']
        core = "-".join(parts)
        if td and core:
            return f"TD:{td} {core}"
        return core or (f"TD:{td}" if td else "")
    u_doc_ple = [format_doc_ple(td, s, n) for td, s, n in zip(td_ple, s_ple, n_ple)]
    u_doc_ple = np.array(u_doc_ple, dtype=object)

    # Origen
    u_origen = np.where(
        is_match,
        "AMBOS (Reporte y PLE)",
        np.where(is_rep, f"Reporte Contabilidad {year}", f"PLE Diario {year}")
    )

    merged['U_VOUCHER'] = u_voucher
    merged['U_FECHA'] = u_fecha
    merged['U_CUENTA'] = u_cuenta
    merged['U_NOM_CUENTA'] = u_nom_cuenta
    merged['U_DEBE'] = u_debe
    merged['U_HABER'] = u_haber
    merged['U_GLOSA'] = u_glosa
    merged['U_DOC_REP'] = u_doc_rep
    merged['U_DOC_PLE'] = u_doc_ple
    merged['U_ORIGEN'] = u_origen

    # Orden jerárquico para mostrar diferencias primero en las pestañas mensuales:
    # 0 = SOLO EN REPORTE, 1 = SOLO EN PLE, 2 = COINCIDE
    prio_map = {'SOLO EN REPORTE': 0, 'SOLO EN PLE': 1, 'COINCIDE': 2}
    merged['SORT_PRIO'] = merged['ESTADO_CRUCE'].map(prio_map)

    # 7. GENERACIÓN DEL ARCHIVO EXCEL MULTI-HOJA
    print(f"\n[4/5] Escribiendo libro Excel multi-pestaña en '{output_excel}'...")
    t3 = time.time()
    workbook = xlsxwriter.Workbook(output_excel, {'constant_memory': True})

    # Paleta de Formatos Premium
    fmt_header = workbook.add_format({
        'bold': True, 'bg_color': '#1F497D', 'font_color': '#FFFFFF',
        'font_name': 'Calibri', 'font_size': 11, 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    fmt_title = workbook.add_format({
        'bold': True, 'font_size': 14, 'font_color': '#1F497D', 'font_name': 'Calibri'
    })
    fmt_subtitle = workbook.add_format({
        'font_size': 10, 'font_color': '#595959', 'font_name': 'Calibri', 'italic': True
    })
    fmt_card_title = workbook.add_format({
        'bold': True, 'font_size': 10, 'font_color': '#1F497D', 'font_name': 'Calibri', 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'
    })
    fmt_card_val = workbook.add_format({
        'bold': True, 'font_size': 12, 'font_name': 'Calibri', 'border': 1, 'align': 'center', 'num_format': '#,##0'
    })
    fmt_card_pct = workbook.add_format({
        'bold': True, 'font_size': 12, 'font_name': 'Calibri', 'border': 1, 'align': 'center', 'num_format': '0.0%'
    })
    fmt_bold = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 11})
    fmt_number = workbook.add_format({'num_format': '#,##0.00', 'font_name': 'Calibri', 'font_size': 10})
    fmt_integer = workbook.add_format({'num_format': '#,##0', 'font_name': 'Calibri', 'font_size': 10})
    fmt_pct = workbook.add_format({'num_format': '0.0%', 'font_name': 'Calibri', 'font_size': 10})
    fmt_text = workbook.add_format({'font_name': 'Calibri', 'font_size': 10})
    fmt_center = workbook.add_format({'font_name': 'Calibri', 'font_size': 10, 'align': 'center'})

    fmt_total_bold = workbook.add_format({
        'bold': True, 'bg_color': '#D9E1F2', 'font_name': 'Calibri', 'font_size': 11, 'border': 1
    })
    fmt_total_num = workbook.add_format({
        'bold': True, 'bg_color': '#D9E1F2', 'font_name': 'Calibri', 'font_size': 11, 'border': 1, 'num_format': '#,##0.00'
    })
    fmt_total_int = workbook.add_format({
        'bold': True, 'bg_color': '#D9E1F2', 'font_name': 'Calibri', 'font_size': 11, 'border': 1, 'num_format': '#,##0'
    })

    # Formatos de Estado
    fmt_match = workbook.add_format({
        'bg_color': '#C6EFCE', 'font_color': '#006100', 'font_name': 'Calibri', 'font_size': 10, 'bold': True, 'align': 'center'
    })
    fmt_solo_rep = workbook.add_format({
        'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'font_name': 'Calibri', 'font_size': 10, 'bold': True, 'align': 'center'
    })
    fmt_solo_ple = workbook.add_format({
        'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'font_name': 'Calibri', 'font_size': 10, 'bold': True, 'align': 'center'
    })
    fmt_cuadrada = workbook.add_format({
        'bg_color': '#C6EFCE', 'font_color': '#006100', 'font_name': 'Calibri', 'font_size': 10, 'bold': True, 'align': 'center'
    })
    fmt_con_dif = workbook.add_format({
        'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'font_name': 'Calibri', 'font_size': 10, 'bold': True, 'align': 'center'
    })

    # ==========================================================================
    # HOJA 1: RESUMEN_GENERAL
    # ==========================================================================
    print("      -> Generando Hoja 1: 'RESUMEN_GENERAL'...")
    ws_resumen = workbook.add_worksheet('RESUMEN_GENERAL')
    ws_resumen.set_column('A:A', 38)
    ws_resumen.set_column('B:E', 20)
    ws_resumen.set_column('F:H', 18)

    ws_resumen.write('A1', f'REPORTE EJECUTIVO DE CONCILIACIÓN CONTABLE vs PLE DIARIO {year}', fmt_title)
    ws_resumen.write('A2', f'Empresa: CORPORACIÓN COMATPE SAC | RUC: 20516403650 | Asiento Apertura Integrado en Enero', fmt_subtitle)

    ws_resumen.write_row(4, 0, ['MÉTRICA / CONCEPTO', 'CANTIDAD REGISTROS', 'TOTAL DEBE (S/.)', 'TOTAL HABER (S/.)', '% DEL TOTAL'], fmt_header)

    sum_rep_debe = df_rep['REP_Debe_Num'].sum()
    sum_rep_haber = df_rep['REP_Haber_Num'].sum()
    sum_ple_debe = df_ple['PLE_Debe_Num'].sum()
    sum_ple_haber = df_ple['PLE_Haber_Num'].sum()

    coinciden_rows = merged[merged['ESTADO_CRUCE'] == 'COINCIDE']
    sum_coincide_debe = coinciden_rows['REP_Debe_Num'].sum()
    sum_coincide_haber = coinciden_rows['REP_Haber_Num'].sum()

    solo_rep_rows = merged[merged['ESTADO_CRUCE'] == 'SOLO EN REPORTE']
    sum_solo_rep_debe = solo_rep_rows['REP_Debe_Num'].sum()
    sum_solo_rep_haber = solo_rep_rows['REP_Haber_Num'].sum()

    solo_ple_rows = merged[merged['ESTADO_CRUCE'] == 'SOLO EN PLE']
    sum_solo_ple_debe = solo_ple_rows['PLE_Debe_Num'].sum()
    sum_solo_ple_haber = solo_ple_rows['PLE_Haber_Num'].sum()

    tot_universo_debe = sum_coincide_debe + sum_solo_rep_debe + sum_solo_ple_debe
    tot_universo_haber = sum_coincide_haber + sum_solo_rep_haber + sum_solo_ple_haber

    filas_resumen = [
        (f'1. Total Registros Reporte Contabilidad {year}', len(df_rep), sum_rep_debe, sum_rep_haber, '-'),
        (f'2. Total Registros PLE Diario Declarado {year}', len(df_ple), sum_ple_debe, sum_ple_haber, '-'),
        ('3. Coincidencias Exactas (Ambos)', n_coincide, sum_coincide_debe, sum_coincide_haber, n_coincide / len(merged)),
        ('4. Solo en Reporte Contable (Faltantes en PLE)', n_solo_rep, sum_solo_rep_debe, sum_solo_rep_haber, n_solo_rep / len(merged)),
        ('5. Solo en PLE Diario (No están en Reporte)', n_solo_ple, sum_solo_ple_debe, sum_solo_ple_haber, n_solo_ple / len(merged)),
        ('TOTAL UNIVERSO CRUCE FULL OUTER JOIN', len(merged), tot_universo_debe, tot_universo_haber, 1.0)
    ]

    for idx, (concepto, cant, d, h, pct) in enumerate(filas_resumen, start=5):
        is_total = (idx == 10)
        ws_resumen.write(idx, 0, concepto, fmt_total_bold if is_total else (fmt_text if idx < 10 else fmt_bold))
        ws_resumen.write(idx, 1, cant, fmt_total_int if is_total else fmt_integer)
        ws_resumen.write(idx, 2, d, fmt_total_num if is_total else fmt_number)
        ws_resumen.write(idx, 3, h, fmt_total_num if is_total else fmt_number)
        if isinstance(pct, (int, float)):
            ws_resumen.write(idx, 4, pct, fmt_total_bold if is_total else fmt_pct)
        else:
            ws_resumen.write(idx, 4, pct, fmt_text)

    # Cuadro comparativo mes a mes
    ws_resumen.write('A13', 'DESGLOSE COMPARATIVO MENSUAL (ENERO A DICIEMBRE)', fmt_title)
    ws_resumen.write_row(14, 0, [
        'PERÍODO / MES', 'REG. REPORTE', 'REG. PLE', 'COINCIDEN', 'SOLO REPORTE', 'SOLO PLE', '% COINCIDENCIA', 'ESTADO MES'
    ], fmt_header)

    curr_row = 15
    tot_m_rep = 0
    tot_m_ple = 0
    tot_m_both = 0
    tot_m_solo_rep = 0
    tot_m_solo_ple = 0

    for m_code in sorted(MONTH_NAMES.keys()):
        m_name = MONTH_NAMES[m_code]
        df_m = merged[merged['MES_COD'] == m_code]
        c_both = int((df_m['ESTADO_CRUCE'] == 'COINCIDE').sum())
        c_only_rep = int((df_m['ESTADO_CRUCE'] == 'SOLO EN REPORTE').sum())
        c_only_ple = int((df_m['ESTADO_CRUCE'] == 'SOLO EN PLE').sum())
        c_rep = c_both + c_only_rep
        c_ple = c_both + c_only_ple
        pct_match = (c_both / c_ple) if c_ple > 0 else 1.0

        tot_m_rep += c_rep
        tot_m_ple += c_ple
        tot_m_both += c_both
        tot_m_solo_rep += c_only_rep
        tot_m_solo_ple += c_only_ple

        st_mes = "CUADRADO 100%" if (c_only_rep == 0 and c_only_ple == 0) else "CON DIFERENCIAS"
        fmt_st_mes = fmt_cuadrada if (c_only_rep == 0 and c_only_ple == 0) else fmt_con_dif

        ws_resumen.write(curr_row, 0, f"Mes {m_code} - {m_name.split('_')[1]}", fmt_text)
        ws_resumen.write(curr_row, 1, c_rep, fmt_integer)
        ws_resumen.write(curr_row, 2, c_ple, fmt_integer)
        ws_resumen.write(curr_row, 3, c_both, fmt_integer)
        ws_resumen.write(curr_row, 4, c_only_rep, fmt_integer)
        ws_resumen.write(curr_row, 5, c_only_ple, fmt_integer)
        ws_resumen.write(curr_row, 6, pct_match, fmt_pct)
        ws_resumen.write(curr_row, 7, st_mes, fmt_st_mes)
        curr_row += 1

    # Total General Meses
    ws_resumen.write(curr_row, 0, 'TOTAL GENERAL', fmt_total_bold)
    ws_resumen.write(curr_row, 1, tot_m_rep, fmt_total_int)
    ws_resumen.write(curr_row, 2, tot_m_ple, fmt_total_int)
    ws_resumen.write(curr_row, 3, tot_m_both, fmt_total_int)
    ws_resumen.write(curr_row, 4, tot_m_solo_rep, fmt_total_int)
    ws_resumen.write(curr_row, 5, tot_m_solo_ple, fmt_total_int)
    ws_resumen.write(curr_row, 6, tot_m_both / tot_m_ple if tot_m_ple > 0 else 1.0, fmt_total_bold)
    ws_resumen.write(curr_row, 7, 'AUDITORÍA CONSOLIDADA', fmt_total_bold)

    # ==========================================================================
    # HOJA 2: RESUMEN_POR_CUENTA
    # ==========================================================================
    print("      -> Generando Hoja 2: 'RESUMEN_POR_CUENTA'...")
    ws_cta = workbook.add_worksheet('RESUMEN_POR_CUENTA')
    ws_cta.set_column('A:A', 16)
    ws_cta.set_column('B:B', 38)
    ws_cta.set_column('C:G', 15)
    ws_cta.set_column('H:K', 18)
    ws_cta.set_column('L:N', 18)
    ws_cta.set_column('O:O', 18)

    ws_cta.write('A1', f'CONSOLIDADO DE CONCILIACIÓN POR CUENTA CONTABLE - AÑO {year}', fmt_title)
    ws_cta.write('A2', 'Resumen anual por cuenta: movimientos en Reporte, movimientos en PLE, diferencias y saldos netos', fmt_subtitle)

    headers_cta = [
        'CUENTA_CONTABLE', 'NOMBRE_CUENTA', 'REG_REPORTE', 'REG_PLE', 'COINCIDEN',
        'SOLO_REPORTE', 'SOLO_PLE', 'DEBE_REPORTE', 'HABER_REPORTE', 'DEBE_PLE',
        'HABER_PLE', 'DIF_DEBE', 'DIF_HABER', 'DIF_SALDO_NETO', 'ESTADO_CUENTA'
    ]
    ws_cta.write_row(3, 0, headers_cta, fmt_header)

    # Agrupar por cuenta contable
    all_accounts = sorted(set(merged['U_CUENTA'].unique()))
    cta_curr_row = 4

    tot_c_rep = 0
    tot_c_ple = 0
    tot_c_both = 0
    tot_c_solo_rep = 0
    tot_c_solo_ple = 0
    tot_c_d_rep = 0.0
    tot_c_h_rep = 0.0
    tot_c_d_ple = 0.0
    tot_c_h_ple = 0.0

    for cta in all_accounts:
        sub = merged[merged['U_CUENTA'] == cta]
        n_both = int((sub['ESTADO_CRUCE'] == 'COINCIDE').sum())
        n_s_rep = int((sub['ESTADO_CRUCE'] == 'SOLO EN REPORTE').sum())
        n_s_ple = int((sub['ESTADO_CRUCE'] == 'SOLO EN PLE').sum())
        n_rep = n_both + n_s_rep
        n_ple = n_both + n_s_ple

        # Montos Reporte
        d_rep_sum = float(sub['REP_Debe_Num'].fillna(0.0).sum())
        h_rep_sum = float(sub['REP_Haber_Num'].fillna(0.0).sum())

        # Montos PLE
        d_ple_sum = float(sub['PLE_Debe_Num'].fillna(0.0).sum())
        h_ple_sum = float(sub['PLE_Haber_Num'].fillna(0.0).sum())

        dif_debe = round(d_rep_sum - d_ple_sum, 2)
        dif_haber = round(h_rep_sum - h_ple_sum, 2)
        dif_saldo = round((d_rep_sum - h_rep_sum) - (d_ple_sum - h_ple_sum), 2)

        nom = cta_map.get(cta, '')
        if not nom:
            prefix2 = cta[:2]
            nom = CLASS_9_DESC.get(prefix2, '')

        is_cuadrada = (n_s_rep == 0 and n_s_ple == 0 and abs(dif_debe) < 0.01 and abs(dif_haber) < 0.01)
        st_cta = "CUADRADA" if is_cuadrada else "CON DIFERENCIAS"
        fmt_st_cta = fmt_cuadrada if is_cuadrada else fmt_con_dif

        tot_c_rep += n_rep
        tot_c_ple += n_ple
        tot_c_both += n_both
        tot_c_solo_rep += n_s_rep
        tot_c_solo_ple += n_s_ple
        tot_c_d_rep += d_rep_sum
        tot_c_h_rep += h_rep_sum
        tot_c_d_ple += d_ple_sum
        tot_c_h_ple += h_ple_sum

        ws_cta.write(cta_curr_row, 0, cta, fmt_text)
        ws_cta.write(cta_curr_row, 1, nom, fmt_text)
        ws_cta.write(cta_curr_row, 2, n_rep, fmt_integer)
        ws_cta.write(cta_curr_row, 3, n_ple, fmt_integer)
        ws_cta.write(cta_curr_row, 4, n_both, fmt_integer)
        ws_cta.write(cta_curr_row, 5, n_s_rep, fmt_integer)
        ws_cta.write(cta_curr_row, 6, n_s_ple, fmt_integer)
        ws_cta.write(cta_curr_row, 7, d_rep_sum, fmt_number)
        ws_cta.write(cta_curr_row, 8, h_rep_sum, fmt_number)
        ws_cta.write(cta_curr_row, 9, d_ple_sum, fmt_number)
        ws_cta.write(cta_curr_row, 10, h_ple_sum, fmt_number)
        ws_cta.write(cta_curr_row, 11, dif_debe, fmt_number)
        ws_cta.write(cta_curr_row, 12, dif_haber, fmt_number)
        ws_cta.write(cta_curr_row, 13, dif_saldo, fmt_number)
        ws_cta.write(cta_curr_row, 14, st_cta, fmt_st_cta)
        cta_curr_row += 1

    # Fila Total General Cuentas
    ws_cta.write(cta_curr_row, 0, 'TOTAL GENERAL', fmt_total_bold)
    ws_cta.write(cta_curr_row, 1, f'{len(all_accounts)} CUENTAS CONTABLES', fmt_total_bold)
    ws_cta.write(cta_curr_row, 2, tot_c_rep, fmt_total_int)
    ws_cta.write(cta_curr_row, 3, tot_c_ple, fmt_total_int)
    ws_cta.write(cta_curr_row, 4, tot_c_both, fmt_total_int)
    ws_cta.write(cta_curr_row, 5, tot_c_solo_rep, fmt_total_int)
    ws_cta.write(cta_curr_row, 6, tot_c_solo_ple, fmt_total_int)
    ws_cta.write(cta_curr_row, 7, tot_c_d_rep, fmt_total_num)
    ws_cta.write(cta_curr_row, 8, tot_c_h_rep, fmt_total_num)
    ws_cta.write(cta_curr_row, 9, tot_c_d_ple, fmt_total_num)
    ws_cta.write(cta_curr_row, 10, tot_c_h_ple, fmt_total_num)
    ws_cta.write(cta_curr_row, 11, tot_c_d_rep - tot_c_d_ple, fmt_total_num)
    ws_cta.write(cta_curr_row, 12, tot_c_h_rep - tot_c_h_ple, fmt_total_num)
    ws_cta.write(cta_curr_row, 13, (tot_c_d_rep - tot_c_h_rep) - (tot_c_d_ple - tot_c_h_ple), fmt_total_num)
    ws_cta.write(cta_curr_row, 14, 'TOTAL CONSOLIDADO', fmt_total_bold)

    # ==========================================================================
    # HOJA 3: DIFERENCIAS_TOTALES (Vectorial de alta velocidad)
    # ==========================================================================
    print("      -> Generando Hoja 3: 'DIFERENCIAS_TOTALES'...")
    ws_dif = workbook.add_worksheet('DIFERENCIAS_TOTALES')
    ws_dif.set_column('A:A', 18)
    ws_dif.set_column('B:B', 15)
    ws_dif.set_column('C:C', 18)
    ws_dif.set_column('D:D', 12)
    ws_dif.set_column('E:E', 15)
    ws_dif.set_column('F:F', 32)
    ws_dif.set_column('G:H', 15)
    ws_dif.set_column('I:I', 35)
    ws_dif.set_column('J:K', 16)
    ws_dif.set_column('L:L', 25)
    ws_dif.set_column('M:M', 38)

    headers_dif = [
        'ESTADO_CRUCE', 'MES', 'VOUCHER / REG_CTB', 'FECHA_MOV', 'CUENTA_CONTABLE',
        'NOMBRE_CUENTA', 'DEBE', 'HABER', 'GLOSA', 'DOC_REPORTE', 'DOC_PLE',
        'ORIGEN_REGISTRO', 'CLAVE_CONCATENADA'
    ]
    ws_dif.write_row(0, 0, headers_dif, fmt_header)

    dif_df = merged[merged['ESTADO_CRUCE'] != 'COINCIDE'].sort_values(['MES_COD', 'SORT_PRIO', 'U_VOUCHER', 'U_CUENTA'])

    d_est = dif_df['ESTADO_CRUCE'].to_numpy()
    d_mes = dif_df['MES_COD'].map(MONTH_NAMES).to_numpy()
    d_vouc = dif_df['U_VOUCHER'].to_numpy()
    d_fec = dif_df['U_FECHA'].to_numpy()
    d_cta = dif_df['U_CUENTA'].to_numpy()
    d_nom = dif_df['U_NOM_CUENTA'].to_numpy()
    d_debe = dif_df['U_DEBE'].to_numpy()
    d_haber = dif_df['U_HABER'].to_numpy()
    d_glo = dif_df['U_GLOSA'].to_numpy()
    d_doc_r = dif_df['U_DOC_REP'].to_numpy()
    d_doc_p = dif_df['U_DOC_PLE'].to_numpy()
    d_orig = dif_df['U_ORIGEN'].to_numpy()
    d_clav = dif_df['CLAVE_CONCATENADA'].fillna('').astype(str).to_numpy()

    for i in range(len(dif_df)):
        r = i + 1
        st = d_est[i]
        fmt_st = fmt_solo_rep if st == 'SOLO EN REPORTE' else fmt_solo_ple

        ws_dif.write(r, 0, st, fmt_st)
        ws_dif.write(r, 1, str(d_mes[i]), fmt_text)
        ws_dif.write(r, 2, str(d_vouc[i]), fmt_text)
        ws_dif.write(r, 3, str(d_fec[i]), fmt_center)
        ws_dif.write(r, 4, str(d_cta[i]), fmt_text)
        ws_dif.write(r, 5, str(d_nom[i]), fmt_text)
        ws_dif.write_number(r, 6, float(d_debe[i]), fmt_number)
        ws_dif.write_number(r, 7, float(d_haber[i]), fmt_number)
        ws_dif.write(r, 8, str(d_glo[i]), fmt_text)
        ws_dif.write(r, 9, str(d_doc_r[i]), fmt_text)
        ws_dif.write(r, 10, str(d_doc_p[i]), fmt_text)
        ws_dif.write(r, 11, str(d_orig[i]), fmt_text)
        ws_dif.write(r, 12, str(d_clav[i]), fmt_text)

    # ==========================================================================
    # HOJA 4: CRUCE_ANUAL (Full Outer Join Completo de 387,638 filas)
    # ==========================================================================
    print("      -> Generando Hoja 4: 'CRUCE_ANUAL' (Full Outer Join total)...")
    ws_full = workbook.add_worksheet('CRUCE_ANUAL')
    ws_full.set_column('A:A', 18)
    ws_full.set_column('B:B', 38)
    ws_full.set_column('C:L', 14)
    ws_full.set_column('M:V', 14)

    full_headers = [
        'ESTADO_CRUCE', 'CLAVE_CONCATENADA',
        # Columnas Reporte
        'REP_Reg_Ctb', 'REP_Prdo', 'REP_Fec_Mov', 'REP_Nro_Cta_Original',
        'REP_Nro_Cta_Limpia', 'REP_Nombre_Cuenta', 'REP_Debe', 'REP_Haber', 'REP_Glosa', 'REP_Doc',
        # Columnas PLE
        'PLE_Mes', 'PLE_Voucher', 'PLE_Seccion', 'PLE_Cuenta',
        'PLE_Fecha', 'PLE_Debe', 'PLE_Haber', 'PLE_Glosa', 'PLE_TD', 'PLE_Doc'
    ]
    ws_full.write_row(0, 0, full_headers, fmt_header)

    f_estados = merged['ESTADO_CRUCE'].to_numpy()
    f_claves = merged['CLAVE_CONCATENADA'].fillna('').astype(str).to_numpy()

    # Columnas Reporte
    r_vouchers = merged['REP_Reg_Ctb'].fillna('').astype(str).to_numpy()
    r_prdos = merged['REP_Prdo'].fillna('').astype(str).to_numpy()
    r_fechas = merged['REP_Fec_Mov'].fillna('').astype(str).to_numpy()
    r_ctas_orig = merged['REP_Nro_Cta_Original'].fillna('').astype(str).to_numpy()
    r_ctas_clean = merged['REP_Nro_Cta_Limpia'].fillna('').astype(str).to_numpy()
    r_nom_ctas = merged['REP_Nombre_Cuenta'].fillna('').astype(str).to_numpy()
    r_debes = merged['REP_Debe_Num'].to_numpy()
    r_haberes = merged['REP_Haber_Num'].to_numpy()
    r_glosas = merged['REP_Glosa'].fillna('').astype(str).to_numpy()
    r_docs = u_doc_rep

    # Columnas PLE
    p_meses = merged['PLE_Mes'].fillna('').astype(str).to_numpy()
    p_vouchers = merged['PLE_Voucher'].fillna('').astype(str).to_numpy()
    p_secs = merged['PLE_Seccion'].fillna('').astype(str).to_numpy()
    p_ctas = merged['PLE_Cuenta'].fillna('').astype(str).to_numpy()
    p_fechas = merged['PLE_Fecha'].fillna('').astype(str).to_numpy()
    p_debes = merged['PLE_Debe_Num'].to_numpy()
    p_haberes = merged['PLE_Haber_Num'].to_numpy()
    p_glosas = merged['PLE_Glosa'].fillna('').astype(str).to_numpy()
    p_tds = merged['PLE_TD'].fillna('').astype(str).to_numpy()
    p_docs = u_doc_ple

    n_total = len(merged)
    for i in range(n_total):
        r = i + 1
        st = f_estados[i]
        fmt_st = fmt_match if st == 'COINCIDE' else (fmt_solo_rep if st == 'SOLO EN REPORTE' else fmt_solo_ple)

        ws_full.write(r, 0, st, fmt_st)
        ws_full.write(r, 1, f_claves[i], fmt_text)

        # Rep
        ws_full.write(r, 2, r_vouchers[i], fmt_text)
        ws_full.write(r, 3, r_prdos[i], fmt_center)
        ws_full.write(r, 4, r_fechas[i], fmt_center)
        ws_full.write(r, 5, r_ctas_orig[i], fmt_text)
        ws_full.write(r, 6, r_ctas_clean[i], fmt_text)
        ws_full.write(r, 7, r_nom_ctas[i], fmt_text)
        if not np.isnan(r_debes[i]):
            ws_full.write_number(r, 8, float(r_debes[i]), fmt_number)
        else:
            ws_full.write_blank(r, 8, None)
        if not np.isnan(r_haberes[i]):
            ws_full.write_number(r, 9, float(r_haberes[i]), fmt_number)
        else:
            ws_full.write_blank(r, 9, None)
        ws_full.write(r, 10, r_glosas[i], fmt_text)
        ws_full.write(r, 11, r_docs[i], fmt_text)

        # PLE
        ws_full.write(r, 12, p_meses[i], fmt_text)
        ws_full.write(r, 13, p_vouchers[i], fmt_text)
        ws_full.write(r, 14, p_secs[i], fmt_text)
        ws_full.write(r, 15, p_ctas[i], fmt_text)
        ws_full.write(r, 16, p_fechas[i], fmt_center)
        if not np.isnan(p_debes[i]):
            ws_full.write_number(r, 17, float(p_debes[i]), fmt_number)
        else:
            ws_full.write_blank(r, 17, None)
        if not np.isnan(p_haberes[i]):
            ws_full.write_number(r, 18, float(p_haberes[i]), fmt_number)
        else:
            ws_full.write_blank(r, 18, None)
        ws_full.write(r, 19, p_glosas[i], fmt_text)
        ws_full.write(r, 20, p_tds[i], fmt_text)
        ws_full.write(r, 21, p_docs[i], fmt_text)

    # ==========================================================================
    # HOJAS 5 A 16: HOJAS MENSUALES (01_ENERO A 12_DICIEMBRE)
    # Con Diferencias primero (Rojo / Amarillo) y luego Coincidencias (Verde)
    # ==========================================================================
    print("\n      -> Generando Hojas Mensuales (01_ENERO a 12_DICIEMBRE)...")
    headers_mes = [
        'ESTADO_CRUCE', 'VOUCHER / REG_CTB', 'FECHA_MOV', 'CUENTA_CONTABLE',
        'NOMBRE_CUENTA', 'DEBE', 'HABER', 'GLOSA', 'DOC_REPORTE', 'DOC_PLE',
        'ORIGEN_REGISTRO', 'CLAVE_CONCATENADA'
    ]

    for m_code in sorted(MONTH_NAMES.keys()):
        sheet_title = MONTH_NAMES[m_code]
        print(f"         * Escribiendo pestaña '{sheet_title}'...")
        ws_m = workbook.add_worksheet(sheet_title)

        ws_m.set_column('A:A', 18)
        ws_m.set_column('B:B', 18)
        ws_m.set_column('C:C', 12)
        ws_m.set_column('D:D', 15)
        ws_m.set_column('E:E', 32)
        ws_m.set_column('F:G', 15)
        ws_m.set_column('H:H', 35)
        ws_m.set_column('I:J', 16)
        ws_m.set_column('K:K', 25)
        ws_m.set_column('L:L', 38)

        # Filtrar y ordenar mes: SORT_PRIO (0: SOLO REP, 1: SOLO PLE, 2: COINCIDE)
        df_sub_m = merged[merged['MES_COD'] == m_code].sort_values(['SORT_PRIO', 'U_VOUCHER', 'U_CUENTA'])

        m_both = int((df_sub_m['ESTADO_CRUCE'] == 'COINCIDE').sum())
        m_solo_r = int((df_sub_m['ESTADO_CRUCE'] == 'SOLO EN REPORTE').sum())
        m_solo_p = int((df_sub_m['ESTADO_CRUCE'] == 'SOLO EN PLE').sum())
        m_total_reg = len(df_sub_m)
        m_pct_cuad = (m_both / (m_both + m_solo_p)) if (m_both + m_solo_p) > 0 else 1.0

        # Mini panel de tarjetas en encabezado
        mes_nombre_txt = sheet_title.split('_')[1]
        ws_m.write('A1', f'CONCILIACIÓN CONTABLE vs PLE - MES {mes_nombre_txt} {year}', fmt_title)
        if m_code == '01':
            ws_m.write('A2', 'Incluye movimientos de Asiento de Apertura (Mes 00) integrados en Enero', fmt_subtitle)
        else:
            ws_m.write('A2', f'Auditoría mensual: Diferencias al inicio para revisión inmediata | {year}', fmt_subtitle)

        # Tarjetas de resumen
        ws_m.write('B4', 'TOTAL MOVIMIENTOS', fmt_card_title)
        ws_m.write('B5', m_total_reg, fmt_card_val)

        ws_m.write('C4', 'COINCIDENCIAS EXACTAS', fmt_card_title)
        ws_m.write('C5', m_both, fmt_card_val)

        ws_m.write('D4', '% CUADRATURA PLE', fmt_card_title)
        ws_m.write('D5', m_pct_cuad, fmt_card_pct)

        ws_m.write('E4', 'DIF. SOLO EN REPORTE', fmt_card_title)
        ws_m.write('E5', m_solo_r, fmt_card_val)

        ws_m.write('F4', 'DIF. SOLO EN PLE', fmt_card_title)
        ws_m.write('F5', m_solo_p, fmt_card_val)

        # Fila de encabezados de la tabla de datos
        ws_m.write_row(6, 0, headers_mes, fmt_header)

        # Vectores de datos para escritura ultrarrápida
        sm_est = df_sub_m['ESTADO_CRUCE'].to_numpy()
        sm_vouc = df_sub_m['U_VOUCHER'].to_numpy()
        sm_fec = df_sub_m['U_FECHA'].to_numpy()
        sm_cta = df_sub_m['U_CUENTA'].to_numpy()
        sm_nom = df_sub_m['U_NOM_CUENTA'].to_numpy()
        sm_debe = df_sub_m['U_DEBE'].to_numpy()
        sm_haber = df_sub_m['U_HABER'].to_numpy()
        sm_glo = df_sub_m['U_GLOSA'].to_numpy()
        sm_doc_r = df_sub_m['U_DOC_REP'].to_numpy()
        sm_doc_p = df_sub_m['U_DOC_PLE'].to_numpy()
        sm_orig = df_sub_m['U_ORIGEN'].to_numpy()
        sm_clav = df_sub_m['CLAVE_CONCATENADA'].fillna('').astype(str).to_numpy()

        for j in range(m_total_reg):
            row_idx = j + 7
            st = sm_est[j]
            fmt_st = fmt_match if st == 'COINCIDE' else (fmt_solo_rep if st == 'SOLO EN REPORTE' else fmt_solo_ple)

            ws_m.write(row_idx, 0, st, fmt_st)
            ws_m.write(row_idx, 1, str(sm_vouc[j]), fmt_text)
            ws_m.write(row_idx, 2, str(sm_fec[j]), fmt_center)
            ws_m.write(row_idx, 3, str(sm_cta[j]), fmt_text)
            ws_m.write(row_idx, 4, str(sm_nom[j]), fmt_text)
            ws_m.write_number(row_idx, 5, float(sm_debe[j]), fmt_number)
            ws_m.write_number(row_idx, 6, float(sm_haber[j]), fmt_number)
            ws_m.write(row_idx, 7, str(sm_glo[j]), fmt_text)
            ws_m.write(row_idx, 8, str(sm_doc_r[j]), fmt_text)
            ws_m.write(row_idx, 9, str(sm_doc_p[j]), fmt_text)
            ws_m.write(row_idx, 10, str(sm_orig[j]), fmt_text)
            ws_m.write(row_idx, 11, str(sm_clav[j]), fmt_text)

    # 8. Cierre y guardado final
    print("\n[5/5] Finalizando escritura en disco y cerrando libro Excel...")
    t_close = time.time()
    workbook.close()
    print(f"      -> Archivo generado exitosamente en {time.time()-t_close:.2f}s")
    file_size_mb = os.path.getsize(output_excel) / (1024 * 1024)
    print(f"      -> Ruta final: {output_excel}")
    print(f"      -> Tamaño del archivo: {file_size_mb:.2f} MB")
    print(f"\nPROCESO TOTAL COMPLETADO EN {time.time()-start_time:.2f} SEGUNDOS.")
    print("=" * 80)

if __name__ == '__main__':
    main()
