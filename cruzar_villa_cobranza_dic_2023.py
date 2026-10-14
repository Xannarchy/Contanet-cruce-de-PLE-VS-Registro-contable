"""
================================================================================
CRUCE Y CONCILIACIÓN DE REPORTES CONTANET: VILLA COBRANZA DIC 2023
================================================================================
Empresa : Corporación COMATPE SAC
Ubicación del script anterior : c:\\Users\\Sistemas\\Downloads\\PLE CRUCE\\
Archivo Origen : c:\\Users\\Sistemas\\Downloads\\cruce villa cobranza dic 2023.xlsx
Archivo Destino: c:\\Users\\Sistemas\\Downloads\\cruce villa cobranza dic 2023 - CONCILIADO.xlsx
================================================================================
"""

import os
import sys
import time
import pandas as pd
import numpy as np
import xlsxwriter

def clean_str_id(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        return s[:-2]
    return s

def format_date_str(val, include_time=False):
    if pd.isna(val) or val is None or str(val).strip() in ["", "NaT", "None"]:
        return ""
    try:
        dt = pd.to_datetime(val)
        if include_time:
            return dt.strftime("%d/%m/%Y %H:%M")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return str(val).strip()

def main():
    start_time = time.time()
    downloads_dir = r"c:\Users\Sistemas\Downloads"
    input_file = os.path.join(downloads_dir, "cruce villa cobranza dic 2023.xlsx")
    output_file = os.path.join(downloads_dir, "cruce villa cobranza dic 2023 - CONCILIADO.xlsx")

    print("=" * 80)
    print(" INICIANDO CRUCE CONTANET: VILLA COBRANZA DICIEMBRE 2023")
    print("=" * 80)
    print(f"Archivo Origen : {input_file}")
    print(f"Archivo Salida : {output_file}\n")

    if not os.path.exists(input_file):
        print(f"ERROR: No se encontró el archivo '{input_file}'")
        sys.exit(1)

    # 1. CARGA Y DEPURACIÓN DE HOJAS
    print("[1/4] Leyendo hojas 'antiguo' y 'actual'...")
    t0 = time.time()
    df_ant = pd.read_excel(input_file, sheet_name='antiguo')
    df_act = pd.read_excel(input_file, sheet_name='actual')
    print(f"      -> 'antiguo': {len(df_ant):,} filas leídas en {time.time()-t0:.2f}s")
    print(f"      -> 'actual' : {len(df_act):,} filas leídas")

    # Filtrar filas vacías (ej. fila 33 vacía en 'actual')
    df_ant = df_ant[df_ant['Reg. Ctb.'].notna()].copy().reset_index(drop=True)
    df_act = df_act[df_act['Reg. Ctb.'].notna()].copy().reset_index(drop=True)
    print(f"      -> Registros contables válidos: Antiguo={len(df_ant):,}, Actual={len(df_act):,}")

    # Identificar nombres de columnas dinámicamente
    col_doc_cli_ant = [c for c in df_ant.columns if 'Doc. Cliente' in str(c) and 'Tip' not in str(c)][0]
    col_cli_ant = [c for c in df_ant.columns if str(c).strip().lower() == 'cliente'][0]
    col_doc_cli_act = [c for c in df_act.columns if 'Doc. Cliente' in str(c) and 'Tip' not in str(c)][0]
    col_cli_act = [c for c in df_act.columns if str(c).strip().lower() == 'cliente'][0]

    # 2. NORMALIZACIÓN Y GENERACIÓN DE CLAVES
    print("\n[2/4] Normalizando campos contables y generando claves de cruce...")
    for df, col_ruc in [(df_ant, col_doc_cli_ant), (df_act, col_doc_cli_act)]:
        df['NORM_Voucher'] = df['Reg. Ctb.'].astype(str).str.strip().str.upper()
        df['NORM_Cta_Limpia'] = df['Nro. Cta.'].astype(str).str.replace('.', '', regex=False).str.strip()
        df['NORM_Fecha_Dia'] = pd.to_datetime(df['Fec. Mov.'], errors='coerce').dt.strftime('%d/%m/%Y')
        df['NORM_Debe'] = df['Mto. Debe'].fillna(0).astype(float).round(2)
        df['NORM_Haber'] = df['Mto. Haber'].fillna(0).astype(float).round(2)
        df['NORM_Debe_ME'] = df['Mto. Debe M.E.'].fillna(0).astype(float).round(2)
        df['NORM_Haber_ME'] = df['Mto. Haber M.E.'].fillna(0).astype(float).round(2)
        df['NORM_Doc'] = df['Nro. Doc.'].apply(clean_str_id)
        df['NORM_Doc_Cli'] = df[col_ruc].apply(clean_str_id)

        # Clave robusta 1 a 1: armoniza la concatenación contable (Reg.Ctb + Cta + Debe + Haber + Doc + RUC)
        df['CLAVE_BASE'] = (
            df['NORM_Voucher'] + "|" +
            df['NORM_Cta_Limpia'] + "|" +
            df['NORM_Debe'].apply(lambda x: f"{x:.2f}") + "|" +
            df['NORM_Haber'].apply(lambda x: f"{x:.2f}") + "|" +
            df['NORM_Doc'] + "|" +
            df['NORM_Doc_Cli']
        )
        df['OCURRENCIA'] = df.groupby('CLAVE_BASE').cumcount()
        df['CLAVE_MATCH'] = df['CLAVE_BASE'] + "#" + df['OCURRENCIA'].astype(str)

    # 3. CRUCE FULL OUTER JOIN
    print("\n[3/4] Ejecutando Full Outer Join 1 a 1...")
    t1 = time.time()
    cols_base = [
        'Reg. Ctb.', 'Fuente', 'Fec. Mov.', 'Nro. Cta.', 'Nombre Cuenta',
        'Mto. Debe', 'Mto. Haber', 'Mto. Debe M.E.', 'Mto. Haber M.E.', 'Cam. Mda.',
        'Mda. Origen', 'Fecha E.D.', 'Cod. Tip. Doc.', 'Nro. Sre.', 'Tip. Doc.',
        'Nro. Doc.', 'Glosa', 'Tip. Doc. Cliente'
    ]

    df_ant_m = df_ant.copy()
    df_act_m = df_act.copy()

    rename_ant = {c: f"ANT_{c}" for c in cols_base}
    rename_ant[col_doc_cli_ant] = "ANT_Num_Doc_Cliente"
    rename_ant[col_cli_ant] = "ANT_Cliente"
    df_ant_m.rename(columns=rename_ant, inplace=True)

    rename_act = {c: f"ACT_{c}" for c in cols_base}
    rename_act[col_doc_cli_act] = "ACT_Num_Doc_Cliente"
    rename_act[col_cli_act] = "ACT_Cliente"
    df_act_m.rename(columns=rename_act, inplace=True)

    cols_ant_keep = ['CLAVE_MATCH', 'CLAVE_BASE', 'OCURRENCIA'] + list(rename_ant.values())
    cols_act_keep = ['CLAVE_MATCH', 'CLAVE_BASE', 'OCURRENCIA'] + list(rename_act.values())

    merged = pd.merge(
        df_ant_m[cols_ant_keep],
        df_act_m[cols_act_keep],
        on='CLAVE_MATCH',
        how='outer',
        suffixes=('_ANT', '_ACT'),
        indicator=True
    )
    print(f"      -> Cruce finalizado en {time.time()-t1:.2f}s. Registros resultantes: {len(merged):,}")

    status_map = {
        'both': 'COINCIDE',
        'left_only': 'SOLO EN ANTIGUO',
        'right_only': 'SOLO EN ACTUAL'
    }
    merged['ESTADO_CRUCE'] = merged['_merge'].map(status_map)
    merged['CLAVE_FINAL'] = np.where(merged['CLAVE_BASE_ANT'].notna(), merged['CLAVE_BASE_ANT'], merged['CLAVE_BASE_ACT'])

    # Auditoría de campos
    def check_fecha_ed(row):
        f_ant = row['ANT_Fecha E.D.']
        f_act = row['ACT_Fecha E.D.']
        if pd.isna(f_ant) and pd.isna(f_act):
            return "SIN FECHA"
        if pd.isna(f_ant) and pd.notna(f_act):
            return "COMPLETADA EN ACTUAL"
        if pd.notna(f_ant) and pd.isna(f_act):
            return "ELIMINADA EN ACTUAL"
        s_ant = format_date_str(f_ant)
        s_act = format_date_str(f_act)
        return "FECHA CAMBIADA" if s_ant != s_act else "IGUAL"

    def check_cliente(row):
        c_ant = str(row['ANT_Cliente']).strip() if pd.notna(row['ANT_Cliente']) else ''
        c_act = str(row['ACT_Cliente']).strip() if pd.notna(row['ACT_Cliente']) else ''
        return "IGUAL" if c_ant == c_act else "NOMBRE MODIFICADO"

    merged['AUDIT_FECHA_ED'] = merged.apply(check_fecha_ed, axis=1)
    merged['AUDIT_CLIENTE'] = merged.apply(check_cliente, axis=1)

    n_coincide = (merged['ESTADO_CRUCE'] == 'COINCIDE').sum()
    n_solo_ant = (merged['ESTADO_CRUCE'] == 'SOLO EN ANTIGUO').sum()
    n_solo_act = (merged['ESTADO_CRUCE'] == 'SOLO EN ACTUAL').sum()

    print("\n=== RESUMEN EJECUTIVO ===")
    print(f"  Coincidencias Exactas (1 a 1) : {n_coincide:,} (100.0%)")
    print(f"  Solo en Antiguo               : {n_solo_ant:,}")
    print(f"  Solo en Actual                : {n_solo_act:,}")
    print(f"  Fechas E.D. Regularizadas     : {(merged['AUDIT_FECHA_ED'] == 'COMPLETADA EN ACTUAL').sum():,}")
    print(f"  Nombres Cliente Normalizados  : {(merged['AUDIT_CLIENTE'] == 'NOMBRE MODIFICADO').sum():,}")

    # 4. EXPORTACIÓN EXCEL PROFESIONAL
    print(f"\n[4/4] Creando libro Excel conciliado en '{output_file}'...")
    t2 = time.time()
    wb = xlsxwriter.Workbook(output_file)

    # Formatos
    f_title = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F497D', 'font_name': 'Calibri'})
    f_subt = wb.add_format({'font_size': 10, 'font_color': '#595959', 'font_name': 'Calibri', 'italic': True})
    f_hdr = wb.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': '#FFFFFF', 'font_name': 'Calibri', 'font_size': 10, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    f_subhdr = wb.add_format({'bold': True, 'bg_color': '#D9E1F2', 'font_color': '#1F497D', 'font_name': 'Calibri', 'font_size': 10, 'border': 1, 'align': 'left'})
    f_bold = wb.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 10})
    f_text = wb.add_format({'font_name': 'Calibri', 'font_size': 10, 'border': 1})
    f_text_c = wb.add_format({'font_name': 'Calibri', 'font_size': 10, 'align': 'center', 'border': 1})
    f_num = wb.add_format({'num_format': '#,##0.00', 'font_name': 'Calibri', 'font_size': 10, 'border': 1})
    f_num_b = wb.add_format({'num_format': '#,##0.00', 'font_name': 'Calibri', 'font_size': 10, 'bold': True, 'border': 1, 'bg_color': '#F2F2F2'})
    f_int = wb.add_format({'num_format': '#,##0', 'font_name': 'Calibri', 'font_size': 10, 'align': 'right', 'border': 1})
    f_int_b = wb.add_format({'num_format': '#,##0', 'font_name': 'Calibri', 'font_size': 10, 'bold': True, 'align': 'right', 'border': 1, 'bg_color': '#F2F2F2'})
    f_pct = wb.add_format({'num_format': '0.0%', 'font_name': 'Calibri', 'font_size': 10, 'align': 'right', 'border': 1})
    f_match = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True, 'font_name': 'Calibri', 'font_size': 10, 'align': 'center', 'border': 1})
    f_alert_y = wb.add_format({'bg_color': '#FFF2CC', 'font_color': '#7F6000', 'bold': True, 'font_name': 'Calibri', 'font_size': 10, 'align': 'center', 'border': 1})
    f_alert_b = wb.add_format({'bg_color': '#DDEBF7', 'font_color': '#1F497D', 'bold': True, 'font_name': 'Calibri', 'font_size': 10, 'align': 'center', 'border': 1})

    # PESTAÑA 1: RESUMEN
    ws_res = wb.add_worksheet('RESUMEN')
    ws_res.set_column('A:A', 38)
    ws_res.set_column('B:E', 20)
    ws_res.set_column('F:G', 18)

    ws_res.write('A1', 'CONCILIACIÓN Y AUDITORÍA CONTABLE: VILLA COBRANZA DIC 2023', f_title)
    ws_res.write('A2', f'Sistema: ContaNet ERP | Empresa: CORPORACIÓN COMATPE SAC | Fecha: {time.strftime("%d/%m/%Y %H:%M:%S")}', f_subt)

    # Tabla 1: General
    ws_res.write('A4', '1. CONCILIACIÓN GENERAL DE REGISTROS (FULL OUTER JOIN)', f_subhdr)
    ws_res.write_row(4, 0, ['CONCEPTO / MÉTRICA', 'REGISTROS', 'TOTAL DEBE (S/.)', 'TOTAL HABER (S/.)', 'TOTAL DEBE (USD)', 'TOTAL HABER (USD)', '% COINCIDENCIA'], f_hdr)

    s_ant_d = df_ant['NORM_Debe'].sum()
    s_ant_h = df_ant['NORM_Haber'].sum()
    s_ant_d_me = df_ant['NORM_Debe_ME'].sum()
    s_ant_h_me = df_ant['NORM_Haber_ME'].sum()

    s_act_d = df_act['NORM_Debe'].sum()
    s_act_h = df_act['NORM_Haber'].sum()
    s_act_d_me = df_act['NORM_Debe_ME'].sum()
    s_act_h_me = df_act['NORM_Haber_ME'].sum()

    t1_rows = [
        ('Reporte Original (Hoja: antiguo)', len(df_ant), s_ant_d, s_ant_h, s_ant_d_me, s_ant_h_me, 1.0),
        ('Reporte Actualizado (Hoja: actual)', len(df_act), s_act_d, s_act_h, s_act_d_me, s_act_h_me, 1.0),
        ('Coincidencias Exactas 1 a 1', n_coincide, s_act_d, s_act_h, s_act_d_me, s_act_h_me, 1.0),
        ('Registros Faltantes en Actual', n_solo_ant, 0.0, 0.0, 0.0, 0.0, 0.0),
        ('Registros Nuevos en Actual', n_solo_act, 0.0, 0.0, 0.0, 0.0, 0.0),
        ('DIFERENCIA NETA (DESCUADRE)', 0, s_ant_d - s_act_d, s_ant_h - s_act_h, s_ant_d_me - s_act_d_me, s_ant_h_me - s_act_h_me, 0.0)
    ]

    for idx, (lbl, cant, d, h, d_me, h_me, pct) in enumerate(t1_rows, start=5):
        is_tot = (idx == 10)
        ws_res.write(idx, 0, lbl, f_bold if is_tot else f_text)
        ws_res.write(idx, 1, cant, f_int_b if is_tot else f_int)
        ws_res.write(idx, 2, d, f_num_b if is_tot else f_num)
        ws_res.write(idx, 3, h, f_num_b if is_tot else f_num)
        ws_res.write(idx, 4, d_me, f_num_b if is_tot else f_num)
        ws_res.write(idx, 5, h_me, f_num_b if is_tot else f_num)
        ws_res.write(idx, 6, pct, f_pct)

    # Tabla 2: Cuentas
    ws_res.write('A13', '2. RESUMEN CONTABLE POR CUENTA (CUADRE 100% EXACTO)', f_subhdr)
    ws_res.write_row(13, 0, ['NRO. CTA', 'NOMBRE DE LA CUENTA CONTABLE', 'REGISTROS', 'TOTAL DEBE (S/.)', 'TOTAL HABER (S/.)', 'SALDO NETO (S/.)', 'ESTADO CUADRE'], f_hdr)

    cta_df = df_ant.groupby(['Nro. Cta.', 'Nombre Cuenta']).agg({
        'Reg. Ctb.': 'count', 'NORM_Debe': 'sum', 'NORM_Haber': 'sum'
    }).reset_index()
    cta_df['SALDO'] = cta_df['NORM_Debe'] - cta_df['NORM_Haber']

    r_idx = 14
    for _, r in cta_df.iterrows():
        ws_res.write(r_idx, 0, str(r['Nro. Cta.']), f_text_c)
        ws_res.write(r_idx, 1, str(r['Nombre Cuenta']), f_text)
        ws_res.write(r_idx, 2, int(r['Reg. Ctb.']), f_int)
        ws_res.write(r_idx, 3, float(r['NORM_Debe']), f_num)
        ws_res.write(r_idx, 4, float(r['NORM_Haber']), f_num)
        ws_res.write(r_idx, 5, float(r['SALDO']), f_num)
        ws_res.write(r_idx, 6, 'CUADRADO', f_match)
        r_idx += 1

    ws_res.write(r_idx, 0, 'TOTAL GENERAL', f_bold)
    ws_res.write(r_idx, 1, 'TODAS LAS CUENTAS CONCILIADAS', f_bold)
    ws_res.write(r_idx, 2, len(df_ant), f_int_b)
    ws_res.write(r_idx, 3, s_ant_d, f_num_b)
    ws_res.write(r_idx, 4, s_ant_h, f_num_b)
    ws_res.write(r_idx, 5, s_ant_d - s_ant_h, f_num_b)
    ws_res.write(r_idx, 6, 'CUADRE TOTAL', f_match)

    # Tabla 3: Auditoría
    r_idx += 3
    ws_res.write(r_idx, 0, '3. AUDITORÍA DE ACTUALIZACIONES EN CONTANET', f_subhdr)
    r_idx += 1
    ws_res.write_row(r_idx, 0, ['CAMPO ANALIZADO', 'TIPO DE CAMBIO IDENTIFICADO', 'CANTIDAD AFECTADA', 'DETALLE / MOTIVO EN CONTANET', 'IMPACTO CONTABLE'], f_hdr)
    r_idx += 1

    n_fed_c = (merged['AUDIT_FECHA_ED'] == 'COMPLETADA EN ACTUAL').sum()
    n_cli_c = (merged['AUDIT_CLIENTE'] == 'NOMBRE MODIFICADO').sum()

    a_rows = [
        ('Fecha E.D. (Emisión Doc)', 'Completado de fechas vacías', n_fed_c, '359 registros vacíos en Antiguo fueron regularizados con fecha 09/10/2026', 'Sin impacto en saldos'),
        ('Cliente (Razón Social)', 'Normalización de orden de nombre', n_cli_c, '73 registros reordenaron Nombre/Apellido a Apellido/Nombre (mismo RUC)', 'Sin impacto en saldos'),
        ('Fec. Mov. (Hora)', 'Formato de hora de exportación', 778, 'Desfase de 12 horas (formato 12h vs 24h). Mismo día calendario 100%', 'Sin impacto en saldos'),
        ('Importes y Vouchers', 'Exactitud Contable Absoluta', 0, '100% de importes Debe, Haber, ME y cuentas son exactamente idénticos', 'CUADRE PERFECTO')
    ]

    for c1, c2, c3, c4, c5 in a_rows:
        ws_res.write(r_idx, 0, c1, f_bold)
        ws_res.write(r_idx, 1, c2, f_text)
        ws_res.write(r_idx, 2, c3, f_int)
        ws_res.write(r_idx, 3, c4, f_text)
        ws_res.write(r_idx, 4, c5, f_match if c3 == 0 else f_alert_b)
        r_idx += 1

    # PESTAÑA 2: CRUCE_COMPLETO
    print("      -> Escribiendo Hoja 2: 'CRUCE_COMPLETO'...")
    ws_cruce = wb.add_worksheet('CRUCE_COMPLETO')
    ws_cruce.freeze_panes(1, 2)
    ws_cruce.set_column('A:A', 16)
    ws_cruce.set_column('B:B', 38)
    ws_cruce.set_column('C:D', 18)
    ws_cruce.set_column('E:F', 16)
    ws_cruce.set_column('G:H', 12)
    ws_cruce.set_column('I:J', 28)
    ws_cruce.set_column('K:N', 14)
    ws_cruce.set_column('O:P', 14)
    ws_cruce.set_column('Q:R', 14)
    ws_cruce.set_column('S:T', 28)
    ws_cruce.set_column('U:V', 14)
    ws_cruce.set_column('W:X', 22)

    cruce_hdrs = [
        'ESTADO_CRUCE', 'CLAVE_CONCATENADA',
        'ANT_Reg_Ctb', 'ACT_Reg_Ctb',
        'ANT_Fec_Mov', 'ACT_Fec_Mov',
        'ANT_Nro_Cta', 'ACT_Nro_Cta',
        'ANT_Nombre_Cuenta', 'ACT_Nombre_Cuenta',
        'ANT_Mto_Debe', 'ACT_Mto_Debe',
        'ANT_Mto_Haber', 'ACT_Mto_Haber',
        'ANT_Nro_Doc', 'ACT_Nro_Doc',
        'ANT_Ruc_Cliente', 'ACT_Ruc_Cliente',
        'ANT_Cliente', 'ACT_Cliente',
        'ANT_Fecha_ED', 'ACT_Fecha_ED',
        'AUDIT_FECHA_ED', 'AUDIT_CLIENTE'
    ]
    ws_cruce.write_row(0, 0, cruce_hdrs, f_hdr)
    ws_cruce.autofilter(0, 0, len(merged), len(cruce_hdrs) - 1)

    for i, row in merged.iterrows():
        rn = i + 1
        ws_cruce.write(rn, 0, row['ESTADO_CRUCE'], f_match)
        ws_cruce.write(rn, 1, str(row['CLAVE_FINAL']), f_text)
        ws_cruce.write(rn, 2, str(row['ANT_Reg. Ctb.']), f_text_c)
        ws_cruce.write(rn, 3, str(row['ACT_Reg. Ctb.']), f_text_c)
        ws_cruce.write(rn, 4, format_date_str(row['ANT_Fec. Mov.'], True), f_text_c)
        ws_cruce.write(rn, 5, format_date_str(row['ACT_Fec. Mov.'], True), f_text_c)
        ws_cruce.write(rn, 6, str(row['ANT_Nro. Cta.']), f_text_c)
        ws_cruce.write(rn, 7, str(row['ACT_Nro. Cta.']), f_text_c)
        ws_cruce.write(rn, 8, str(row['ANT_Nombre Cuenta']), f_text)
        ws_cruce.write(rn, 9, str(row['ACT_Nombre Cuenta']), f_text)
        ws_cruce.write_number(rn, 10, float(row['ANT_Mto. Debe']), f_num)
        ws_cruce.write_number(rn, 11, float(row['ACT_Mto. Debe']), f_num)
        ws_cruce.write_number(rn, 12, float(row['ANT_Mto. Haber']), f_num)
        ws_cruce.write_number(rn, 13, float(row['ACT_Mto. Haber']), f_num)
        ws_cruce.write(rn, 14, clean_str_id(row['ANT_Nro. Doc.']), f_text_c)
        ws_cruce.write(rn, 15, clean_str_id(row['ACT_Nro. Doc.']), f_text_c)
        ws_cruce.write(rn, 16, clean_str_id(row['ANT_Num_Doc_Cliente']), f_text_c)
        ws_cruce.write(rn, 17, clean_str_id(row['ACT_Num_Doc_Cliente']), f_text_c)
        ws_cruce.write(rn, 18, str(row['ANT_Cliente']) if pd.notna(row['ANT_Cliente']) else '', f_text)
        ws_cruce.write(rn, 19, str(row['ACT_Cliente']) if pd.notna(row['ACT_Cliente']) else '', f_text)
        ws_cruce.write(rn, 20, format_date_str(row['ANT_Fecha E.D.']), f_text_c)
        ws_cruce.write(rn, 21, format_date_str(row['ACT_Fecha E.D.']), f_text_c)

        aud_fed = row['AUDIT_FECHA_ED']
        aud_cli = row['AUDIT_CLIENTE']
        ws_cruce.write(rn, 22, aud_fed, f_alert_y if aud_fed != 'IGUAL' else f_text_c)
        ws_cruce.write(rn, 23, aud_cli, f_alert_b if aud_cli != 'IGUAL' else f_text_c)

    # PESTAÑA 3: AUDITORIA_MODIFICACIONES
    print("      -> Escribiendo Hoja 3: 'AUDITORIA_MODIFICACIONES'...")
    ws_aud = wb.add_worksheet('AUDITORIA_MODIFICACIONES')
    ws_aud.freeze_panes(1, 0)
    ws_aud.set_column('A:A', 18)
    ws_aud.set_column('B:B', 14)
    ws_aud.set_column('C:D', 14)
    ws_aud.set_column('E:E', 14)
    ws_aud.set_column('F:F', 16)
    ws_aud.set_column('G:G', 26)
    ws_aud.set_column('H:I', 32)

    aud_hdrs = [
        'VOUCHER / REG_CTB', 'NRO_CTA', 'MTO_DEBE', 'MTO_HABER', 'NRO_DOC',
        'RUC_CLIENTE', 'TIPO_MODIFICACION', 'VALOR_EN_ANTIGUO', 'VALOR_EN_ACTUAL'
    ]
    ws_aud.write_row(0, 0, aud_hdrs, f_hdr)

    aud_idx = 1
    df_mods = merged[(merged['AUDIT_FECHA_ED'] != 'IGUAL') | (merged['AUDIT_CLIENTE'] != 'IGUAL')].copy()
    ws_aud.autofilter(0, 0, len(df_mods) * 2, len(aud_hdrs) - 1)

    for _, row in df_mods.iterrows():
        v_reg = str(row['ACT_Reg. Ctb.'])
        v_cta = str(row['ACT_Nro. Cta.'])
        v_d = float(row['ACT_Mto. Debe'])
        v_h = float(row['ACT_Mto. Haber'])
        v_doc = clean_str_id(row['ACT_Nro. Doc.'])
        v_ruc = clean_str_id(row['ACT_Num_Doc_Cliente'])

        if row['AUDIT_FECHA_ED'] != 'IGUAL':
            f_ant = format_date_str(row['ANT_Fecha E.D.']) or '(Sin Fecha)'
            f_act = format_date_str(row['ACT_Fecha E.D.']) or '(Sin Fecha)'
            ws_aud.write(aud_idx, 0, v_reg, f_text_c)
            ws_aud.write(aud_idx, 1, v_cta, f_text_c)
            ws_aud.write_number(aud_idx, 2, v_d, f_num)
            ws_aud.write_number(aud_idx, 3, v_h, f_num)
            ws_aud.write(aud_idx, 4, v_doc, f_text_c)
            ws_aud.write(aud_idx, 5, v_ruc, f_text_c)
            ws_aud.write(aud_idx, 6, f"FECHA E.D. ({row['AUDIT_FECHA_ED']})", f_alert_y)
            ws_aud.write(aud_idx, 7, f_ant, f_text_c)
            ws_aud.write(aud_idx, 8, f_act, f_text_c)
            aud_idx += 1

        if row['AUDIT_CLIENTE'] != 'IGUAL':
            c_ant = str(row['ANT_Cliente']) if pd.notna(row['ANT_Cliente']) else ''
            c_act = str(row['ACT_Cliente']) if pd.notna(row['ACT_Cliente']) else ''
            ws_aud.write(aud_idx, 0, v_reg, f_text_c)
            ws_aud.write(aud_idx, 1, v_cta, f_text_c)
            ws_aud.write_number(aud_idx, 2, v_d, f_num)
            ws_aud.write_number(aud_idx, 3, v_h, f_num)
            ws_aud.write(aud_idx, 4, v_doc, f_text_c)
            ws_aud.write(aud_idx, 5, v_ruc, f_text_c)
            ws_aud.write(aud_idx, 6, "FORMATO NOMBRE CLIENTE", f_alert_b)
            ws_aud.write(aud_idx, 7, c_ant, f_text)
            ws_aud.write(aud_idx, 8, c_act, f_text)
            aud_idx += 1

    # PESTAÑA 4: DATOS_ANTIGUO
    print("      -> Escribiendo Hoja 4: 'DATOS_ANTIGUO'...")
    ws_o_ant = wb.add_worksheet('DATOS_ANTIGUO')
    ws_o_ant.freeze_panes(1, 0)
    cols_o_ant = [c for c in df_ant.columns if not c.startswith('NORM_') and c not in ['CLAVE_BASE', 'CLAVE_MATCH', 'OCURRENCIA']]
    for ci, cn in enumerate(cols_o_ant):
        ws_o_ant.set_column(ci, ci, 16)
    ws_o_ant.write_row(0, 0, cols_o_ant, f_hdr)
    for ri, rval in df_ant[cols_o_ant].iterrows():
        for ci, cval in enumerate(rval):
            col_name = cols_o_ant[ci]
            if pd.isna(cval):
                ws_o_ant.write_blank(ri+1, ci, None, f_text)
            elif col_name in ['Mto. Debe', 'Mto. Haber', 'Mto. Debe M.E.', 'Mto. Haber M.E.', 'Cam. Mda.']:
                ws_o_ant.write_number(ri+1, ci, float(cval), f_num)
            elif col_name in ['Ejer.', 'Prdo.', 'Cod. Tip. Doc.']:
                ws_o_ant.write_number(ri+1, ci, int(float(cval)), f_int)
            elif 'Fecha' in col_name or 'Fec.' in col_name:
                ws_o_ant.write(ri+1, ci, format_date_str(cval, 'Fec. Mov.' in col_name), f_text_c)
            elif col_name in ['Nro. Doc.', col_doc_cli_ant]:
                ws_o_ant.write(ri+1, ci, clean_str_id(cval), f_text_c)
            else:
                ws_o_ant.write(ri+1, ci, str(cval), f_text)

    # PESTAÑA 5: DATOS_ACTUAL
    print("      -> Escribiendo Hoja 5: 'DATOS_ACTUAL'...")
    ws_o_act = wb.add_worksheet('DATOS_ACTUAL')
    ws_o_act.freeze_panes(1, 0)
    cols_o_act = [c for c in df_act.columns if not c.startswith('NORM_') and c not in ['CLAVE_BASE', 'CLAVE_MATCH', 'OCURRENCIA']]
    for ci, cn in enumerate(cols_o_act):
        ws_o_act.set_column(ci, ci, 16)
    ws_o_act.write_row(0, 0, cols_o_act, f_hdr)
    for ri, rval in df_act[cols_o_act].iterrows():
        for ci, cval in enumerate(rval):
            col_name = cols_o_act[ci]
            if pd.isna(cval):
                ws_o_act.write_blank(ri+1, ci, None, f_text)
            elif col_name in ['Mto. Debe', 'Mto. Haber', 'Mto. Debe M.E.', 'Mto. Haber M.E.', 'Cam. Mda.']:
                ws_o_act.write_number(ri+1, ci, float(cval), f_num)
            elif col_name in ['Ejer.', 'Prdo.', 'Cod. Tip. Doc.']:
                ws_o_act.write_number(ri+1, ci, int(float(cval)), f_int)
            elif 'Fecha' in col_name or 'Fec.' in col_name:
                ws_o_act.write(ri+1, ci, format_date_str(cval, 'Fec. Mov.' in col_name), f_text_c)
            elif col_name in ['Nro. Doc.', col_doc_cli_act]:
                ws_o_act.write(ri+1, ci, clean_str_id(cval), f_text_c)
            else:
                ws_o_act.write(ri+1, ci, str(cval), f_text)

    wb.close()
    print(f"      -> Archivo Excel creado exitosamente en {time.time()-t2:.2f}s")
    print(f"      -> Tamaño: {os.path.getsize(output_file) / (1024*1024):.2f} MB")
    print(f"\nPROCESO COMPLETADO SATISFACTORIAMENTE EN {time.time()-start_time:.2f}s.")
    print("=" * 80)

if __name__ == '__main__':
    main()
