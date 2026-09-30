"""
Script para Cruzar el Reporte de Contabilidad 2023 con el PLE Mensual Diario 2023
Generado para: Corporación COMATPE SAC
Ruta de salida: Cruce_Contabilidad_vs_PLE_Diario_2023.xlsx
"""

import os
import time
import pandas as pd
import numpy as np
import xlsxwriter

def main():
    start_time = time.time()
    downloads_dir = r"c:\Users\Sistemas\Downloads"
    
    file_reporte = os.path.join(downloads_dir, "Reporte de Contabilidad Fecha 01-01-2023 al 31-12-2023.xlsx")
    file_ple = os.path.join(downloads_dir, "Mensual ple diario 2023 decl.xlsx")
    output_excel = os.path.join(downloads_dir, "Cruce_Contabilidad_vs_PLE_Diario_2023.xlsx")

    print("================================================================================")
    print(" INICIANDO CRUCE CONTABILIDAD vs PLE DIARIO 2023")
    print("================================================================================")

    # 1. CARGA DE REPORTE DE CONTABILIDAD
    print(f"\n[1/4] Leyendo '{os.path.basename(file_reporte)}'...")
    cols_rep = [
        'Ejer.', 'Prdo.', 'Reg. Ctb.', 'Fuente', 'Fec. Mov.', 'Nro. Cta.',
        'Nombre Cuenta', 'Mto. Debe', 'Mto. Haber', 'Glosa', 'Cod. Tip. Doc.', 'Nro. Doc.'
    ]
    t0 = time.time()
    df_rep = pd.read_excel(
        file_reporte,
        sheet_name='Contabilidad',
        header=5,
        engine='calamine',
        usecols=cols_rep
    )
    print(f"      -> {len(df_rep):,} filas leídas en {time.time()-t0:.2f}s")

    # Limpieza Reporte
    df_rep = df_rep[df_rep['Reg. Ctb.'].notna() & df_rep['Nro. Cta.'].notna()].copy()
    
    # "quítale los puntos a la columna nro. cta. de el reporte de contabilidad"
    df_rep['REP_Nro_Cta_Original'] = df_rep['Nro. Cta.'].astype(str).str.strip()
    df_rep['REP_Nro_Cta_Limpia'] = df_rep['Nro. Cta.'].astype(str).str.replace('.', '', regex=False).str.strip()
    df_rep['REP_Reg_Ctb'] = df_rep['Reg. Ctb.'].astype(str).str.strip().str.upper()
    df_rep['REP_Fec_Mov'] = pd.to_datetime(df_rep['Fec. Mov.'], errors='coerce').dt.strftime('%d/%m/%Y')
    df_rep['REP_Prdo'] = df_rep['Prdo.'].fillna(0).astype(int).astype(str)
    df_rep['REP_Nombre_Cuenta'] = df_rep['Nombre Cuenta'].fillna('').astype(str).str.strip()
    df_rep['REP_Glosa'] = df_rep['Glosa'].fillna('').astype(str).str.strip()
    df_rep['REP_Tip_Doc'] = df_rep['Cod. Tip. Doc.'].fillna('').astype(str).str.strip()
    df_rep['REP_Nro_Doc'] = df_rep['Nro. Doc.'].fillna('').astype(str).str.strip()

    df_rep['REP_Debe_Num'] = df_rep['Mto. Debe'].fillna(0).astype(float).round(2)
    df_rep['REP_Haber_Num'] = df_rep['Mto. Haber'].fillna(0).astype(float).round(2)
    df_rep['REP_Debe_Str'] = df_rep['REP_Debe_Num'].apply(lambda x: f"{x:.2f}")
    df_rep['REP_Haber_Str'] = df_rep['REP_Haber_Num'].apply(lambda x: f"{x:.2f}")

    # Clave de concatenación solicitada: Reg. Ctb. + Fec. Mov. + Nro. Cta. + Mto. Debe + Mto. Haber
    df_rep['CLAVE_CONCAT'] = (
        df_rep['REP_Reg_Ctb'] + "|" +
        df_rep['REP_Fec_Mov'] + "|" +
        df_rep['REP_Nro_Cta_Limpia'] + "|" +
        df_rep['REP_Debe_Str'] + "|" +
        df_rep['REP_Haber_Str']
    )
    
    # Manejo de asientos repetidos dentro del mismo voucher para cruce 1 a 1 exacto
    df_rep['OCURRENCIA'] = df_rep.groupby('CLAVE_CONCAT').cumcount()
    df_rep['CLAVE_MATCH'] = df_rep['CLAVE_CONCAT'] + "#" + df_rep['OCURRENCIA'].astype(str)
    print(f"      -> {len(df_rep):,} registros contables válidos normalizados.")

    # 2. CARGA DE PLE DIARIO
    print(f"\n[2/4] Leyendo '{os.path.basename(file_ple)}'...")
    cols_ple = [
        'MES', 'VOUCHER', 'SECCION', 'CUENTA', 'MONEDA', 'TD', 'SERIE', 'NUMERO',
        'FECHA 1', 'GLOSA', 'DEBE', 'HABER'
    ]
    t1 = time.time()
    df_ple = pd.read_excel(
        file_ple,
        sheet_name=0,
        header=0,
        engine='calamine',
        usecols=cols_ple
    )
    print(f"      -> {len(df_ple):,} filas leídas en {time.time()-t1:.2f}s")

    # Limpieza PLE
    df_ple = df_ple[df_ple['VOUCHER'].notna() & df_ple['CUENTA'].notna()].copy()
    
    df_ple['PLE_Cuenta'] = df_ple['CUENTA'].astype(str).str.replace('.', '', regex=False).str.split('.').str[0].str.strip()
    df_ple['PLE_Voucher'] = df_ple['VOUCHER'].astype(str).str.strip().str.upper()
    df_ple['PLE_Fecha'] = pd.to_datetime(df_ple['FECHA 1'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
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

    # Clave de concatenación en PLE
    df_ple['CLAVE_CONCAT'] = (
        df_ple['PLE_Voucher'] + "|" +
        df_ple['PLE_Fecha'] + "|" +
        df_ple['PLE_Cuenta'] + "|" +
        df_ple['PLE_Debe_Str'] + "|" +
        df_ple['PLE_Haber_Str']
    )
    df_ple['OCURRENCIA'] = df_ple.groupby('CLAVE_CONCAT').cumcount()
    df_ple['CLAVE_MATCH'] = df_ple['CLAVE_CONCAT'] + "#" + df_ple['OCURRENCIA'].astype(str)
    print(f"      -> {len(df_ple):,} registros PLE válidos normalizados.")

    # 3. FULL OUTER JOIN
    print("\n[3/4] Ejecutando Full Outer Join...")
    t2 = time.time()
    
    # Seleccionar solo las columnas necesarias para el merge
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
    print(f"      -> Cruce finalizado en {time.time()-t2:.2f}s. Total registros resultantes: {len(merged):,}")

    # Asignar estado del cruce
    status_map = {
        'both': 'COINCIDE',
        'left_only': 'SOLO EN REPORTE',
        'right_only': 'SOLO EN PLE'
    }
    merged['ESTADO_CRUCE'] = merged['_merge'].map(status_map)

    # Clave consolidada
    merged['CLAVE_CONCATENADA'] = np.where(
        merged['CLAVE_CONCAT_REP'].notna(),
        merged['CLAVE_CONCAT_REP'],
        merged['CLAVE_CONCAT_PLE']
    )

    # Métricas del cruce
    conteo_estado = merged['ESTADO_CRUCE'].value_counts()
    n_coincide = conteo_estado.get('COINCIDE', 0)
    n_solo_rep = conteo_estado.get('SOLO EN REPORTE', 0)
    n_solo_ple = conteo_estado.get('SOLO EN PLE', 0)

    print("\n=== RESUMEN DE COINCIDENCIAS ===")
    print(f"  Total Coincidencias Exactas : {n_coincide:,} ({(n_coincide/len(df_ple))*100:.1f}% de PLE)")
    print(f"  Solo en Reporte Contable    : {n_solo_rep:,}")
    print(f"  Solo en PLE Diario          : {n_solo_ple:,}")

    # 4. EXPORTACIÓN A EXCEL OPTIMIZADA (xlsxwriter streaming)
    print(f"\n[4/4] Generando archivo Excel en '{output_excel}'...")
    t3 = time.time()

    workbook = xlsxwriter.Workbook(output_excel, {'constant_memory': True})

    # Formatos
    fmt_header = workbook.add_format({
        'bold': True,
        'bg_color': '#1F497D',
        'font_color': '#FFFFFF',
        'font_name': 'Calibri',
        'font_size': 11,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    fmt_title = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_color': '#1F497D',
        'font_name': 'Calibri'
    })
    fmt_bold = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 11})
    fmt_number = workbook.add_format({'num_format': '#,##0.00', 'font_name': 'Calibri', 'font_size': 10})
    fmt_integer = workbook.add_format({'num_format': '#,##0', 'font_name': 'Calibri', 'font_size': 10})
    fmt_pct = workbook.add_format({'num_format': '0.0%', 'font_name': 'Calibri', 'font_size': 10})
    fmt_text = workbook.add_format({'font_name': 'Calibri', 'font_size': 10})

    # Formatos de estado
    fmt_match = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'font_name': 'Calibri', 'font_size': 10, 'bold': True})
    fmt_solo_rep = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'font_name': 'Calibri', 'font_size': 10, 'bold': True})
    fmt_solo_ple = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'font_name': 'Calibri', 'font_size': 10, 'bold': True})

    # --------------------------------------------------------------------------
    # HOJA 1: RESUMEN
    # --------------------------------------------------------------------------
    ws_resumen = workbook.add_worksheet('RESUMEN')
    ws_resumen.set_column('A:A', 36)
    ws_resumen.set_column('B:E', 20)

    ws_resumen.write('A1', 'REPORTE EJECUTIVO DE CONCILIACIÓN CONTABLE vs PLE DIARIO 2023', fmt_title)
    ws_resumen.write('A2', f'Empresa: CORPORACIÓN COMATPE SAC | RUC: 20516403650 | Fecha Cruce: {time.strftime("%d/%m/%Y %H:%M")}', fmt_bold)

    ws_resumen.write_row(4, 0, ['MÉTRICA / CONCEPTO', 'CANTIDAD REGISTROS', 'TOTAL DEBE (S/.)', 'TOTAL HABER (S/.)', '% DEL TOTAL'], fmt_header)
    
    # Cálculos de montos
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

    filas_resumen = [
        ('1. Total Registros Reporte Contabilidad', len(df_rep), sum_rep_debe, sum_rep_haber, '-'),
        ('2. Total Registros PLE Diario Declarado', len(df_ple), sum_ple_debe, sum_ple_haber, '-'),
        ('3. Coincidencias Exactas (Ambos)', n_coincide, sum_coincide_debe, sum_coincide_haber, n_coincide / len(merged)),
        ('4. Solo en Reporte Contable (Faltantes en PLE)', n_solo_rep, sum_solo_rep_debe, sum_solo_rep_haber, n_solo_rep / len(merged)),
        ('5. Solo en PLE Diario (No están en Reporte)', n_solo_ple, sum_solo_ple_debe, sum_solo_ple_haber, n_solo_ple / len(merged)),
        ('TOTAL UNIVERSO CRUCE FULL OUTER JOIN', len(merged), sum_coincide_debe + sum_solo_rep_debe + sum_solo_ple_debe, sum_coincide_haber + sum_solo_rep_haber + sum_solo_ple_haber, 1.0)
    ]

    for idx, (concepto, cant, d, h, pct) in enumerate(filas_resumen, start=5):
        ws_resumen.write(idx, 0, concepto, fmt_text if idx < 10 else fmt_bold)
        ws_resumen.write(idx, 1, cant, fmt_integer)
        ws_resumen.write(idx, 2, d, fmt_number)
        ws_resumen.write(idx, 3, h, fmt_number)
        if isinstance(pct, (int, float)):
            ws_resumen.write(idx, 4, pct, fmt_pct)
        else:
            ws_resumen.write(idx, 4, pct, fmt_text)

    # Desglose Mensual
    ws_resumen.write('A13', 'DESGLOSE COMPARATIVO POR MES / PERÍODO', fmt_title)
    ws_resumen.write_row(14, 0, ['PERÍODO / MES', 'REG. REPORTE', 'REG. PLE', 'COINCIDEN', 'SOLO REPORTE', 'SOLO PLE'], fmt_header)
    
    # Extraer mes
    merged['MES_CRUCE'] = np.where(
        merged['REP_Prdo'].notna() & (merged['REP_Prdo'] != ''),
        merged['REP_Prdo'].apply(lambda x: f"Mes {int(x):02d}" if str(x).isdigit() else str(x)),
        merged['PLE_Mes'].astype(str).str[-4:-2].apply(lambda x: f"Mes {int(x):02d}" if x.isdigit() else str(x))
    )

    meses_orden = sorted(merged['MES_CRUCE'].unique())
    curr_row = 15
    for mes in meses_orden:
        df_m = merged[merged['MES_CRUCE'] == mes]
        m_both = (df_m['ESTADO_CRUCE'] == 'COINCIDE').sum()
        m_only_rep = (df_m['ESTADO_CRUCE'] == 'SOLO EN REPORTE').sum()
        m_only_ple = (df_m['ESTADO_CRUCE'] == 'SOLO EN PLE').sum()
        m_rep = m_both + m_only_rep
        m_ple = m_both + m_only_ple

        ws_resumen.write(curr_row, 0, mes, fmt_text)
        ws_resumen.write(curr_row, 1, m_rep, fmt_integer)
        ws_resumen.write(curr_row, 2, m_ple, fmt_integer)
        ws_resumen.write(curr_row, 3, m_both, fmt_integer)
        ws_resumen.write(curr_row, 4, m_only_rep, fmt_integer)
        ws_resumen.write(curr_row, 5, m_only_ple, fmt_integer)
        curr_row += 1

    # --------------------------------------------------------------------------
    # HOJA 2: DIFERENCIAS (Vectorial / Alta Velocidad)
    # --------------------------------------------------------------------------
    print("      -> Preparando y escribiendo Hoja 2: 'DIFERENCIAS'...")
    ws_dif = workbook.add_worksheet('DIFERENCIAS')
    ws_dif.set_column('A:A', 18) # Estado
    ws_dif.set_column('B:B', 40) # Clave
    ws_dif.set_column('C:C', 18) # Voucher
    ws_dif.set_column('D:D', 12) # Fecha
    ws_dif.set_column('E:E', 12) # Cuenta
    ws_dif.set_column('F:F', 30) # Nombre Cuenta
    ws_dif.set_column('G:H', 15) # Debe, Haber
    ws_dif.set_column('I:I', 35) # Glosa
    ws_dif.set_column('J:L', 12) # TD, Serie, Nro

    dif_headers = [
        'ESTADO_CRUCE', 'CLAVE_CONCATENADA', 'VOUCHER / REG_CTB', 'FECHA_MOV',
        'CUENTA_CONTABLE', 'NOMBRE_CUENTA', 'DEBE', 'HABER', 'GLOSA',
        'TIPO_DOC', 'SERIE_DOC', 'NUMERO_DOC', 'ORIGEN_REGISTRO'
    ]
    ws_dif.write_row(0, 0, dif_headers, fmt_header)

    dif_df = merged[merged['ESTADO_CRUCE'] != 'COINCIDE'].copy()

    # Pre-calcular columnas unificadas con NumPy
    is_rep = (dif_df['ESTADO_CRUCE'] == 'SOLO EN REPORTE').to_numpy()
    
    dif_estados = dif_df['ESTADO_CRUCE'].to_numpy()
    dif_claves = dif_df['CLAVE_CONCATENADA'].fillna('').astype(str).to_numpy()
    dif_vouchers = np.where(is_rep, dif_df['REP_Reg_Ctb'].fillna(''), dif_df['PLE_Voucher'].fillna(''))
    dif_fechas = np.where(is_rep, dif_df['REP_Fec_Mov'].fillna(''), dif_df['PLE_Fecha'].fillna(''))
    dif_cuentas = np.where(is_rep, dif_df['REP_Nro_Cta_Limpia'].fillna(''), dif_df['PLE_Cuenta'].fillna(''))
    dif_nom_ctas = np.where(is_rep, dif_df['REP_Nombre_Cuenta'].fillna(''), '')
    dif_debes = np.where(is_rep, dif_df['REP_Debe_Num'].fillna(0.0), dif_df['PLE_Debe_Num'].fillna(0.0))
    dif_haberes = np.where(is_rep, dif_df['REP_Haber_Num'].fillna(0.0), dif_df['PLE_Haber_Num'].fillna(0.0))
    dif_glosas = np.where(is_rep, dif_df['REP_Glosa'].fillna(''), dif_df['PLE_Glosa'].fillna(''))
    dif_tds = np.where(is_rep, dif_df['REP_Tip_Doc'].fillna(''), dif_df['PLE_TD'].fillna(''))
    dif_series = np.where(is_rep, '', dif_df['PLE_Serie'].fillna(''))
    dif_nros = np.where(is_rep, dif_df['REP_Nro_Doc'].fillna(''), dif_df['PLE_Numero'].fillna(''))
    dif_origenes = np.where(is_rep, 'Reporte de Contabilidad', 'PLE Mensual Diario 2023')

    n_dif = len(dif_df)
    for i in range(n_dif):
        row_idx = i + 1
        st = dif_estados[i]
        fmt_st = fmt_solo_rep if st == 'SOLO EN REPORTE' else fmt_solo_ple

        ws_dif.write(row_idx, 0, st, fmt_st)
        ws_dif.write(row_idx, 1, str(dif_claves[i]), fmt_text)
        ws_dif.write(row_idx, 2, str(dif_vouchers[i]), fmt_text)
        ws_dif.write(row_idx, 3, str(dif_fechas[i]), fmt_text)
        ws_dif.write(row_idx, 4, str(dif_cuentas[i]), fmt_text)
        ws_dif.write(row_idx, 5, str(dif_nom_ctas[i]), fmt_text)
        ws_dif.write_number(row_idx, 6, float(dif_debes[i]), fmt_number)
        ws_dif.write_number(row_idx, 7, float(dif_haberes[i]), fmt_number)
        ws_dif.write(row_idx, 8, str(dif_glosas[i]), fmt_text)
        ws_dif.write(row_idx, 9, str(dif_tds[i]), fmt_text)
        ws_dif.write(row_idx, 10, str(dif_series[i]), fmt_text)
        ws_dif.write(row_idx, 11, str(dif_nros[i]), fmt_text)
        ws_dif.write(row_idx, 12, str(dif_origenes[i]), fmt_text)

    # --------------------------------------------------------------------------
    # HOJA 3: CRUCE_COMPLETO (Vectorial / Alta Velocidad)
    # --------------------------------------------------------------------------
    print("      -> Preparando y escribiendo Hoja 3: 'CRUCE_COMPLETO'...")
    ws_full = workbook.add_worksheet('CRUCE_COMPLETO')
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

    # Reporte
    r_vouchers = merged['REP_Reg_Ctb'].fillna('').astype(str).to_numpy()
    r_prdos = merged['REP_Prdo'].fillna('').astype(str).to_numpy()
    r_fechas = merged['REP_Fec_Mov'].fillna('').astype(str).to_numpy()
    r_ctas_orig = merged['REP_Nro_Cta_Original'].fillna('').astype(str).to_numpy()
    r_ctas_clean = merged['REP_Nro_Cta_Limpia'].fillna('').astype(str).to_numpy()
    r_nom_ctas = merged['REP_Nombre_Cuenta'].fillna('').astype(str).to_numpy()
    r_debes = merged['REP_Debe_Num'].to_numpy()
    r_haberes = merged['REP_Haber_Num'].to_numpy()
    r_glosas = merged['REP_Glosa'].fillna('').astype(str).to_numpy()
    r_docs = merged['REP_Nro_Doc'].fillna('').astype(str).to_numpy()

    # PLE
    p_meses = merged['PLE_Mes'].fillna('').astype(str).to_numpy()
    p_vouchers = merged['PLE_Voucher'].fillna('').astype(str).to_numpy()
    p_secs = merged['PLE_Seccion'].fillna('').astype(str).to_numpy()
    p_ctas = merged['PLE_Cuenta'].fillna('').astype(str).to_numpy()
    p_fechas = merged['PLE_Fecha'].fillna('').astype(str).to_numpy()
    p_debes = merged['PLE_Debe_Num'].to_numpy()
    p_haberes = merged['PLE_Haber_Num'].to_numpy()
    p_glosas = merged['PLE_Glosa'].fillna('').astype(str).to_numpy()
    p_tds = merged['PLE_TD'].fillna('').astype(str).to_numpy()
    p_series = merged['PLE_Serie'].fillna('').astype(str).to_numpy()
    p_nums = merged['PLE_Numero'].fillna('').astype(str).to_numpy()

    n_total = len(merged)
    for i in range(n_total):
        row_idx = i + 1
        st = f_estados[i]
        if st == 'COINCIDE':
            fmt_st = fmt_match
        elif st == 'SOLO EN REPORTE':
            fmt_st = fmt_solo_rep
        else:
            fmt_st = fmt_solo_ple

        ws_full.write(row_idx, 0, st, fmt_st)
        ws_full.write(row_idx, 1, f_claves[i], fmt_text)

        # Rep
        ws_full.write(row_idx, 2, r_vouchers[i], fmt_text)
        ws_full.write(row_idx, 3, r_prdos[i], fmt_text)
        ws_full.write(row_idx, 4, r_fechas[i], fmt_text)
        ws_full.write(row_idx, 5, r_ctas_orig[i], fmt_text)
        ws_full.write(row_idx, 6, r_ctas_clean[i], fmt_text)
        ws_full.write(row_idx, 7, r_nom_ctas[i], fmt_text)
        if not np.isnan(r_debes[i]):
            ws_full.write_number(row_idx, 8, float(r_debes[i]), fmt_number)
        else:
            ws_full.write_blank(row_idx, 8, None)
        if not np.isnan(r_haberes[i]):
            ws_full.write_number(row_idx, 9, float(r_haberes[i]), fmt_number)
        else:
            ws_full.write_blank(row_idx, 9, None)
        ws_full.write(row_idx, 10, r_glosas[i], fmt_text)
        ws_full.write(row_idx, 11, r_docs[i], fmt_text)

        # PLE
        ws_full.write(row_idx, 12, p_meses[i], fmt_text)
        ws_full.write(row_idx, 13, p_vouchers[i], fmt_text)
        ws_full.write(row_idx, 14, p_secs[i], fmt_text)
        ws_full.write(row_idx, 15, p_ctas[i], fmt_text)
        ws_full.write(row_idx, 16, p_fechas[i], fmt_text)
        if not np.isnan(p_debes[i]):
            ws_full.write_number(row_idx, 17, float(p_debes[i]), fmt_number)
        else:
            ws_full.write_blank(row_idx, 17, None)
        if not np.isnan(p_haberes[i]):
            ws_full.write_number(row_idx, 18, float(p_haberes[i]), fmt_number)
        else:
            ws_full.write_blank(row_idx, 18, None)
        ws_full.write(row_idx, 19, p_glosas[i], fmt_text)
        ws_full.write(row_idx, 20, p_tds[i], fmt_text)
        
        # doc ple
        p_serie_val = p_series[i]
        p_num_val = p_nums[i]
        if p_serie_val and p_num_val:
            ws_full.write(row_idx, 21, f"{p_serie_val}-{p_num_val}", fmt_text)
        elif p_num_val:
            ws_full.write(row_idx, 21, p_num_val, fmt_text)
        else:
            ws_full.write_blank(row_idx, 21, None)

    workbook.close()
    print(f"      -> Archivo Excel creado exitosamente en {time.time()-t3:.2f}s")
    print(f"      -> Tamaño del archivo: {os.path.getsize(output_excel) / (1024*1024):.2f} MB")
    print(f"\nPROCESO COMPLETADO EN {time.time()-start_time:.2f} SEGUNDOS.")
    print("================================================================================")

if __name__ == '__main__':
    main()
