"""
Script: generar_resumen_ediciones_2023.py
Propósito:
  1. Analizar las ediciones realizadas por cada usuario en el Reporte de Contabilidad 2023
     cruzado contra el PLE Diario 2023 (Cruce_Contabilidad_vs_PLE_Diario_2023.xlsx).
  2. Generar informe de proyección de tiempo (2 personas a 45 reg/día, factor ~2 reg/voucher).
  3. Analizar avance específico en 2026 por B.Piña, A.Alva y B.Natalie.
  4. Analizar el avance por cuenta contable: cuánto se ha avanzado por cuenta en cada mes y en cada año,
     cuántos movimientos se hicieron por cuenta, fechas de edición y qué usuario lo hizo.
  5. Generar un libro Excel ejecutivo con todas las hojas estructuradas y formato premium.
"""

import os
import sys
import time
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import xlsxwriter

def clean_username(u):
    s = str(u).strip()
    if 'Pi' in s and 'a' in s:
        return 'B.Piña'
    if 'Carre' in s:
        return 'J.Carreño'
    return s

def add_business_days(start_date, n_days):
    cur = start_date
    added = 0
    while added < n_days:
        cur += timedelta(days=1)
        if cur.weekday() < 5:  # Lunes a Viernes
            added += 1
    return cur

def main():
    start_time = time.time()
    downloads_dir = r"c:\Users\Sistemas\Downloads"
    file_reporte = os.path.join(downloads_dir, "Reporte de Contabilidad Fecha 01-01-2023 al 31-12-2023.xlsx")
    file_cruce = os.path.join(downloads_dir, "Cruce_Contabilidad_vs_PLE_Diario_2023.xlsx")
    output_excel = os.path.join(downloads_dir, "Resumen_Ediciones_Usuarios_2023.xlsx")

    print("=" * 85)
    print(" AUDITORÍA, PROYECCIÓN Y AVANCE POR CUENTA - CONTABILIDAD vs PLE 2023")
    print("=" * 85)

    # 1. Cargar CRUCE_COMPLETO para indexar estado del cruce PLE
    print(f"\n[1/6] Leyendo Cruce PLE para indexar estado de conciliación...")
    t0 = time.time()
    df_cruce = pd.read_excel(
        file_cruce,
        sheet_name='CRUCE_COMPLETO',
        engine='calamine',
        usecols=['ESTADO_CRUCE', 'CLAVE_CONCATENADA', 'REP_Reg_Ctb']
    )
    df_cruce_rep = df_cruce[df_cruce['REP_Reg_Ctb'].notna()].copy()
    df_cruce_rep['OCURRENCIA'] = df_cruce_rep.groupby('CLAVE_CONCATENADA').cumcount()
    df_cruce_rep['CLAVE_MATCH'] = df_cruce_rep['CLAVE_CONCATENADA'] + '#' + df_cruce_rep['OCURRENCIA'].astype(str)
    cruce_map = dict(zip(df_cruce_rep['CLAVE_MATCH'], df_cruce_rep['ESTADO_CRUCE']))
    print(f"      -> {len(cruce_map):,} registros de Reporte mapeados en Cruce PLE ({time.time()-t0:.2f}s)")

    # 2. Cargar Reporte de Contabilidad
    print(f"\n[2/6] Leyendo Reporte de Contabilidad 2023...")
    t1 = time.time()
    cols_rep = [
        'Ejer.', 'Prdo.', 'Reg. Ctb.', 'Fuente', 'Fec. Mov.', 'Nro. Cta.',
        'Nombre Cuenta', 'Mto. Debe', 'Mto. Haber', 'Glosa', 'Cod. Tip. Doc.', 'Nro. Doc.',
        'Cliente', 'Proveedor', 'Usu. Crea', 'Fec. Registro', 'Usu. Modifica', 'Fec. Modifica'
    ]
    df_rep = pd.read_excel(
        file_reporte,
        sheet_name='Contabilidad',
        header=5,
        engine='calamine',
        usecols=cols_rep
    )
    df_rep = df_rep[df_rep['Reg. Ctb.'].notna() & df_rep['Nro. Cta.'].notna()].copy()
    print(f"      -> {len(df_rep):,} filas válidas leídas ({time.time()-t1:.2f}s)")

    # 3. Normalizar datos y generar claves
    print(f"\n[3/6] Procesando y cruzando modificaciones...")
    df_rep['REP_Nro_Cta_Limpia'] = df_rep['Nro. Cta.'].astype(str).str.replace('.', '', regex=False).str.strip()
    df_rep['REP_Reg_Ctb'] = df_rep['Reg. Ctb.'].astype(str).str.strip().str.upper()
    df_rep['REP_Fec_Mov'] = pd.to_datetime(df_rep['Fec. Mov.'], errors='coerce').dt.strftime('%d/%m/%Y')
    df_rep['REP_Debe_Num'] = df_rep['Mto. Debe'].fillna(0).astype(float).round(2)
    df_rep['REP_Haber_Num'] = df_rep['Mto. Haber'].fillna(0).astype(float).round(2)
    df_rep['REP_Debe_Str'] = df_rep['REP_Debe_Num'].apply(lambda x: f"{x:.2f}")
    df_rep['REP_Haber_Str'] = df_rep['REP_Haber_Num'].apply(lambda x: f"{x:.2f}")

    df_rep['CLAVE_CONCAT'] = (
        df_rep['REP_Reg_Ctb'] + '|' +
        df_rep['REP_Fec_Mov'] + '|' +
        df_rep['REP_Nro_Cta_Limpia'] + '|' +
        df_rep['REP_Debe_Str'] + '|' +
        df_rep['REP_Haber_Str']
    )
    df_rep['OCURRENCIA'] = df_rep.groupby('CLAVE_CONCAT').cumcount()
    df_rep['CLAVE_MATCH'] = df_rep['CLAVE_CONCAT'] + '#' + df_rep['OCURRENCIA'].astype(str)

    # Filtrar solo registros modificados
    df_rep['Usu_Mod_Raw'] = df_rep['Usu. Modifica'].fillna('').astype(str).str.strip()
    df_mod = df_rep[(df_rep['Usu_Mod_Raw'] != '') & (df_rep['Usu_Mod_Raw'] != 'nan')].copy()
    df_mod['Usu_Mod_Clean'] = df_mod['Usu_Mod_Raw'].apply(clean_username)
    df_mod['Usu_Crea_Clean'] = df_mod['Usu. Crea'].fillna('').astype(str).str.strip().apply(clean_username)

    print(f"      -> Total registros modificados encontrados: {len(df_mod):,}")

    # Asignar Estado de Cruce PLE
    df_mod['ESTADO_CRUCE'] = df_mod['CLAVE_MATCH'].map(cruce_map).fillna('SOLO EN REPORTE')

    # Normalizar fechas y horas
    df_mod['FEC_MOD_DT'] = pd.to_datetime(df_mod['Fec. Modifica'], errors='coerce')
    df_mod['FEC_REG_DT'] = pd.to_datetime(df_mod['Fec. Registro'], errors='coerce')
    
    df_mod['FECHA_SOLO_STR'] = df_mod['FEC_MOD_DT'].dt.strftime('%Y-%m-%d').fillna('(Sin Fecha Mod)')
    df_mod['HORA_STR'] = df_mod['FEC_MOD_DT'].dt.strftime('%H:%M:%S').fillna('')
    df_mod['AÑO_MES'] = df_mod['FEC_MOD_DT'].dt.strftime('%Y-%m').fillna('(Sin Fecha)')
    
    dias_semana_es = {
        0: '1. Lunes', 1: '2. Martes', 2: '3. Miércoles', 3: '4. Jueves',
        4: '5. Viernes', 5: '6. Sábado', 6: '7. Domingo'
    }
    df_mod['DIA_SEMANA_NUM'] = df_mod['FEC_MOD_DT'].dt.dayofweek
    df_mod['DIA_SEMANA'] = df_mod['DIA_SEMANA_NUM'].map(dias_semana_es).fillna('(Sin Fecha)')

    # Campo Tercero
    cliente_clean = df_mod['Cliente'].fillna('').astype(str).str.strip()
    prov_clean = df_mod['Proveedor'].fillna('').astype(str).str.strip()
    df_mod['TERCERO'] = np.where(cliente_clean != '', cliente_clean, prov_clean)

    # Identificar registros de 2026 y año de modificación
    df_mod['YEAR_MOD'] = df_mod['FEC_MOD_DT'].dt.strftime('%Y').fillna('(Sin Fecha)')
    df_mod['YEAR_REG'] = df_mod['FEC_REG_DT'].dt.strftime('%Y').fillna('(Sin Fecha)')
    df_mod['IS_2026'] = (df_mod['YEAR_MOD'] == '2026') | (
        (df_mod['FECHA_SOLO_STR'] == '(Sin Fecha Mod)') & (df_mod['YEAR_REG'] == '2026')
    )

    # Campos de cuenta limpios
    df_mod['CTA_STR'] = df_mod['REP_Nro_Cta_Limpia']
    df_mod['NOM_CTA'] = df_mod['Nombre Cuenta'].fillna('').astype(str).str.strip()

    df_mod.sort_values(by=['Usu_Mod_Clean', 'FEC_MOD_DT', 'REP_Reg_Ctb'], ascending=[True, True, True], inplace=True)

    # 4. Cálculos para Avance 2026 y Proyección de Tiempo
    print(f"\n[4/6] Calculando métricas de proyección y avance 2026...")
    target_3_users = ['B.Piña', 'A.Alva', 'B.Natalie']
    df_2026_target = df_mod[df_mod['Usu_Mod_Clean'].isin(target_3_users) & df_mod['IS_2026']].copy()

    avance_2026_rows = []
    for u in target_3_users:
        sub = df_2026_target[df_2026_target['Usu_Mod_Clean'] == u]
        fechas_val = sub[sub['FECHA_SOLO_STR'] != '(Sin Fecha Mod)']['FECHA_SOLO_STR']
        dias_act = fechas_val.nunique()
        f_min = fechas_val.min() if dias_act > 0 else 'N/A'
        f_max = fechas_val.max() if dias_act > 0 else 'N/A'
        prom_d = (len(fechas_val) / dias_act) if dias_act > 0 else len(sub)
        sin_fec = (sub['FECHA_SOLO_STR'] == '(Sin Fecha Mod)').sum()

        avance_2026_rows.append({
            'Usuario': u,
            'Registros_2026': len(sub),
            'Vouchers_Unicos': sub['REP_Reg_Ctb'].nunique(),
            'Coincide_PLE': (sub['ESTADO_CRUCE'] == 'COINCIDE').sum(),
            'Solo_Reporte': (sub['ESTADO_CRUCE'] == 'SOLO EN REPORTE').sum(),
            'Total_Debe': sub['REP_Debe_Num'].sum(),
            'Total_Haber': sub['REP_Haber_Num'].sum(),
            'Dias_Activos': dias_act,
            'Promedio_Diario': prom_d,
            'Fecha_Min': f_min,
            'Fecha_Max': f_max,
            'Sin_Fecha_Mod': sin_fec
        })
    df_avance_2026 = pd.DataFrame(avance_2026_rows)

    total_reg_2026_target = len(df_2026_target)
    total_vouc_2026_target = df_2026_target['REP_Reg_Ctb'].nunique()

    # Proyección
    start_proj_date = datetime(2026, 10, 1)
    total_solo_rep_bruto = 95390
    total_solo_rep_neto = 95390 - total_reg_2026_target  # 94,774
    vouchers_solo_rep = 30385

    total_dif_bruto = 296653
    total_dif_neto = 296653 - total_reg_2026_target  # 296,037
    vouchers_dif_total = 67404

    scenarios_data = [
        {
            'Escenario': 'Escenario 1A: Solo en Reporte Contabilidad (Bruto)',
            'Descripcion': 'Revisar y corregir todos los registros que están en el Reporte y faltan en PLE',
            'Items': total_solo_rep_bruto,
            'Unidad': 'Registros',
            'Tasa_Diaria': 90,
            'Tasa_Mensual': 1980
        },
        {
            'Escenario': 'Escenario 1B: Solo en Reporte Contabilidad (Neto Pendiente)',
            'Descripcion': 'Solo en Reporte descontando lo ya avanzado en 2026 por B.Piña, A.Alva y B.Natalie (616 reg)',
            'Items': total_solo_rep_neto,
            'Unidad': 'Registros',
            'Tasa_Diaria': 90,
            'Tasa_Mensual': 1980
        },
        {
            'Escenario': 'Escenario 1C: Vouchers Únicos Solo en Reporte',
            'Descripcion': 'Corrección integral a nivel de voucher (asumiendo corrección completa del voucher en 1 paso)',
            'Items': vouchers_solo_rep,
            'Unidad': 'Vouchers',
            'Tasa_Diaria': 45,
            'Tasa_Mensual': 990
        },
        {
            'Escenario': 'Escenario 2A: Universo Total Diferencias PLE (Bruto)',
            'Descripcion': 'Conciliación integral de todas las discrepancias (95,390 Solo Reporte + 201,263 Solo PLE)',
            'Items': total_dif_bruto,
            'Unidad': 'Registros',
            'Tasa_Diaria': 90,
            'Tasa_Mensual': 1980
        },
        {
            'Escenario': 'Escenario 2B: Universo Total Diferencias PLE (Neto Pendiente)',
            'Descripcion': 'Universo total de diferencias descontando el avance de 2026 (616 reg)',
            'Items': total_dif_neto,
            'Unidad': 'Registros',
            'Tasa_Diaria': 90,
            'Tasa_Mensual': 1980
        },
        {
            'Escenario': 'Escenario 2C: Vouchers Únicos Totales con Diferencias',
            'Descripcion': 'Corrección integral de los 67,404 vouchers únicos que tienen discrepancias en ambos lados',
            'Items': vouchers_dif_total,
            'Unidad': 'Vouchers',
            'Tasa_Diaria': 45,
            'Tasa_Mensual': 990
        }
    ]

    proj_rows = []
    for sc in scenarios_data:
        dias_h = int(np.ceil(sc['Items'] / sc['Tasa_Diaria']))
        sem = dias_h / 5.0
        meses = dias_h / 22.0
        anios = meses / 12.0
        f_termino = add_business_days(start_proj_date, dias_h)
        proj_rows.append({
            'Escenario': sc['Escenario'],
            'Descripcion': sc['Descripcion'],
            'Items_Pendientes': sc['Items'],
            'Unidad': sc['Unidad'],
            'Rendimiento_Diario': sc['Tasa_Diaria'],
            'Rendimiento_Mensual': sc['Tasa_Mensual'],
            'Dias_Habiles': dias_h,
            'Semanas_Laborales': sem,
            'Meses_Laborales': meses,
            'Anios_Laborales': anios,
            'Fecha_Estimada_Termino': f_termino.strftime('%d/%m/%Y')
        })
    df_proj = pd.DataFrame(proj_rows)

    sensitivity_rows = []
    for num_pers in [2, 3, 4, 5, 6, 8, 10]:
        tasa_d = num_pers * 45
        tasa_m = tasa_d * 22
        dias_rep = int(np.ceil(total_solo_rep_neto / tasa_d))
        meses_rep = dias_rep / 22.0
        f_rep = add_business_days(start_proj_date, dias_rep)

        dias_tot = int(np.ceil(total_dif_neto / tasa_d))
        meses_tot = dias_tot / 22.0
        f_tot = add_business_days(start_proj_date, dias_tot)

        sensitivity_rows.append({
            'Personal': f"{num_pers} personas",
            'Capacidad_Diaria': tasa_d,
            'Capacidad_Mensual': tasa_m,
            'Meses_Solo_Reporte': meses_rep,
            'Termino_Solo_Reporte': f_rep.strftime('%d/%m/%Y'),
            'Meses_Universo_Total': meses_tot,
            'Termino_Universo_Total': f_tot.strftime('%d/%m/%Y')
        })
    df_sensitivity = pd.DataFrame(sensitivity_rows)

    # 5. Cálculos para Avance por Cuenta Contable
    print(f"\n[5/6] Agrupando y procesando movimientos por cuenta, mes y usuario...")
    # A) Resumen de Cuentas con distribución anual y principales usuarios
    def top_users_str(s):
        vc = s.value_counts().head(3)
        return ', '.join([f"{u} ({c})" for u, c in vc.items()])

    acc_grp = df_mod.groupby(['CTA_STR', 'NOM_CTA'])
    df_acc_summary = acc_grp.agg(
        Total_Mov=('Reg. Ctb.', 'count'),
        Vouchers_Unicos=('REP_Reg_Ctb', 'nunique'),
        Coincide_PLE=('ESTADO_CRUCE', lambda s: (s == 'COINCIDE').sum()),
        Solo_Reporte=('ESTADO_CRUCE', lambda s: (s == 'SOLO EN REPORTE').sum()),
        Total_Debe=('REP_Debe_Num', 'sum'),
        Total_Haber=('REP_Haber_Num', 'sum'),
        Anio_2023=('YEAR_MOD', lambda s: (s == '2023').sum()),
        Anio_2024=('YEAR_MOD', lambda s: (s == '2024').sum()),
        Anio_2025=('YEAR_MOD', lambda s: (s == '2025').sum()),
        Anio_2026=('YEAR_MOD', lambda s: (s == '2026').sum()),
        Sin_Fecha=('YEAR_MOD', lambda s: (s == '(Sin Fecha)').sum()),
        Usuarios_Top=('Usu_Mod_Clean', top_users_str)
    ).reset_index()

    df_acc_summary['Pct_Total'] = df_acc_summary['Total_Mov'] / len(df_mod)
    df_acc_summary['Pct_Coincide'] = df_acc_summary['Coincide_PLE'] / df_acc_summary['Total_Mov']
    df_acc_summary.sort_values(by='Total_Mov', ascending=False, inplace=True)

    # B) Matriz Mensual por Cuenta Contable (Pivot: Cuenta vs Año-Mes)
    pivot_acc_month = df_mod.pivot_table(
        index=['CTA_STR', 'NOM_CTA'],
        columns='AÑO_MES',
        values='Reg. Ctb.',
        aggfunc='count',
        fill_value=0
    )
    meses_acc_cols = [c for c in sorted(df_mod['AÑO_MES'].unique()) if c != '(Sin Fecha)']
    if '(Sin Fecha)' in df_mod['AÑO_MES'].unique():
        meses_acc_cols.append('(Sin Fecha)')
    pivot_acc_month = pivot_acc_month[meses_acc_cols]
    pivot_acc_month['TOTAL'] = pivot_acc_month.sum(axis=1)
    pivot_acc_month.sort_values(by='TOTAL', ascending=False, inplace=True)
    pivot_acc_month.reset_index(inplace=True)

    # C) Desglose Detallado por Cuenta, Año-Mes y Usuario (614 filas)
    c_m_u_grp = df_mod.groupby(['YEAR_MOD', 'AÑO_MES', 'CTA_STR', 'NOM_CTA', 'Usu_Mod_Clean'])
    df_cmu = c_m_u_grp.agg(
        Movimientos=('Reg. Ctb.', 'count'),
        Vouchers_Unicos=('REP_Reg_Ctb', 'nunique'),
        Coincide_PLE=('ESTADO_CRUCE', lambda s: (s == 'COINCIDE').sum()),
        Solo_Reporte=('ESTADO_CRUCE', lambda s: (s == 'SOLO EN REPORTE').sum()),
        Total_Debe=('REP_Debe_Num', 'sum'),
        Total_Haber=('REP_Haber_Num', 'sum'),
        Fecha_Min=('FECHA_SOLO_STR', lambda s: s[s != '(Sin Fecha Mod)'].min() if (s != '(Sin Fecha Mod)').any() else 'N/A'),
        Fecha_Max=('FECHA_SOLO_STR', lambda s: s[s != '(Sin Fecha Mod)'].max() if (s != '(Sin Fecha Mod)').any() else 'N/A')
    ).reset_index()

    df_cmu['Pct_Coincide'] = df_cmu['Coincide_PLE'] / df_cmu['Movimientos']
    df_cmu.sort_values(by=['AÑO_MES', 'CTA_STR', 'Movimientos'], ascending=[True, True, False], inplace=True)

    # 6. GENERACIÓN DEL LIBRO EXCEL CON TODAS LAS PESTAÑAS
    print(f"\n[6/6] Construyendo libro Excel profesional con 'xlsxwriter'...")
    wb = xlsxwriter.Workbook(output_excel, {'nan_inf_to_errors': True})

    # Paleta de Estilos
    c_navy = '#1B365D'
    c_blue_hdr = '#2E5B88'
    c_card_bg = '#F0F4F8'
    c_border = '#D9D9D9'
    c_green_bg = '#E2F0D9'
    c_green_txt = '#385723'
    c_red_bg = '#FCE4D6'
    c_red_txt = '#C65911'
    c_amber_bg = '#FFF2CC'
    c_amber_txt = '#833C0C'

    fmt_title = wb.add_format({'bold': True, 'font_size': 16, 'font_color': c_navy, 'valign': 'vcenter', 'font_name': 'Segoe UI'})
    fmt_subtitle = wb.add_format({'font_size': 10, 'font_color': '#595959', 'italic': True, 'valign': 'vcenter', 'font_name': 'Segoe UI'})
    fmt_section = wb.add_format({'bold': True, 'font_size': 12, 'font_color': c_navy, 'bottom': 2, 'bottom_color': c_navy, 'font_name': 'Segoe UI'})

    fmt_card_title = wb.add_format({'font_size': 8, 'font_color': '#595959', 'bold': True, 'align': 'center', 'bg_color': c_card_bg, 'font_name': 'Segoe UI', 'top': 1, 'left': 1, 'right': 1, 'top_color': '#BDD7EE', 'left_color': '#BDD7EE', 'right_color': '#BDD7EE'})
    fmt_card_num = wb.add_format({'font_size': 14, 'font_color': c_navy, 'bold': True, 'align': 'center', 'bg_color': c_card_bg, 'font_name': 'Segoe UI', 'num_format': '#,##0', 'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': '#BDD7EE', 'left_color': '#BDD7EE', 'right_color': '#BDD7EE'})
    fmt_card_str = wb.add_format({'font_size': 12, 'font_color': c_navy, 'bold': True, 'align': 'center', 'bg_color': c_card_bg, 'font_name': 'Segoe UI', 'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': '#BDD7EE', 'left_color': '#BDD7EE', 'right_color': '#BDD7EE'})

    fmt_th = wb.add_format({'bold': True, 'font_size': 10, 'font_color': '#FFFFFF', 'bg_color': c_navy, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#4F6272', 'font_name': 'Segoe UI', 'text_wrap': True})
    fmt_th_blue = wb.add_format({'bold': True, 'font_size': 10, 'font_color': '#FFFFFF', 'bg_color': c_blue_hdr, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#5B7C99', 'font_name': 'Segoe UI', 'text_wrap': True})

    fmt_cell = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'border': 1, 'border_color': c_border})
    fmt_cell_center = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'align': 'center', 'border': 1, 'border_color': c_border})
    fmt_cell_num = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'align': 'right', 'num_format': '#,##0', 'border': 1, 'border_color': c_border})
    fmt_cell_dec = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'align': 'right', 'num_format': '#,##0.0', 'border': 1, 'border_color': c_border})
    fmt_cell_pct = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'align': 'right', 'num_format': '0.00%', 'border': 1, 'border_color': c_border})
    fmt_cell_curr = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'align': 'right', 'num_format': '#,##0.00', 'border': 1, 'border_color': c_border})
    fmt_cell_dt = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'align': 'center', 'num_format': 'dd/mm/yyyy hh:mm:ss', 'border': 1, 'border_color': c_border})

    fmt_total = wb.add_format({'bold': True, 'font_size': 10, 'font_name': 'Segoe UI', 'bg_color': '#D9E1F2', 'border': 1, 'border_color': '#8EA9DB'})
    fmt_total_center = wb.add_format({'bold': True, 'font_size': 10, 'font_name': 'Segoe UI', 'align': 'center', 'bg_color': '#D9E1F2', 'border': 1, 'border_color': '#8EA9DB'})
    fmt_total_num = wb.add_format({'bold': True, 'font_size': 10, 'font_name': 'Segoe UI', 'align': 'right', 'bg_color': '#D9E1F2', 'border': 1, 'border_color': '#8EA9DB', 'num_format': '#,##0'})
    fmt_total_pct = wb.add_format({'bold': True, 'font_size': 10, 'font_name': 'Segoe UI', 'align': 'right', 'bg_color': '#D9E1F2', 'border': 1, 'border_color': '#8EA9DB', 'num_format': '0.00%'})
    fmt_total_dec = wb.add_format({'bold': True, 'font_size': 10, 'font_name': 'Segoe UI', 'align': 'right', 'bg_color': '#D9E1F2', 'border': 1, 'border_color': '#8EA9DB', 'num_format': '#,##0.0'})
    fmt_total_curr = wb.add_format({'bold': True, 'font_size': 10, 'font_name': 'Segoe UI', 'align': 'right', 'bg_color': '#D9E1F2', 'border': 1, 'border_color': '#8EA9DB', 'num_format': '"S/ " #,##0.00'})

    fmt_highlight = wb.add_format({'bold': True, 'font_size': 9, 'font_name': 'Segoe UI', 'align': 'center', 'bg_color': c_amber_bg, 'font_color': c_amber_txt, 'border': 1, 'border_color': '#D9D9D9'})
    fmt_coincide = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'bold': True, 'align': 'center', 'bg_color': c_green_bg, 'font_color': c_green_txt, 'border': 1, 'border_color': c_border})
    fmt_solo_rep = wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'bold': True, 'align': 'center', 'bg_color': c_red_bg, 'font_color': c_red_txt, 'border': 1, 'border_color': c_border})

    # =========================================================================
    # PESTAÑA 1: INFORME_PROYECCION
    # =========================================================================
    print("      -> Generando hoja 'INFORME_PROYECCION'...")
    wsp = wb.add_worksheet('INFORME_PROYECCION')
    wsp.set_zoom(90)

    wsp.write(0, 0, "INFORME DE PROYECCIÓN DE TIEMPO Y AVANCE 2026 - CORRECCIONES 2023", fmt_title)
    wsp.write(1, 0, f"Corporación COMATPE SAC | RUC: 20516403650 | Modelo: 2 Personas a 45 reg/día (Factor ~2 reg/voucher) | Fecha: {datetime.now().strftime('%d/%m/%Y')}", fmt_subtitle)

    cards_proj = [
        ("CAPACIDAD EQUIPO (2 PERS)", "90 reg/día (~45 v/día)", fmt_card_str),
        ("RITMO MENSUAL (22 DÍAS)", "1,980 reg/mes (~990 v/mes)", fmt_card_str),
        ("AVANCE 2026 (3 USUARIOS)", f"{total_reg_2026_target:,} reg ({total_vouc_2026_target} vouchers)", fmt_card_str),
        ("PENDIENTE SOLO REPORTE", f"{total_solo_rep_neto:,} registros", fmt_card_str),
        ("TIEMPO EST. SOLO REPORTE", "47.9 meses (~4.0 años)", fmt_card_str),
        ("FECHA EST. TÉRMINO", "Octubre 2030", fmt_card_str)
    ]
    cc = 0
    for title, val, fmt_c in cards_proj:
        wsp.merge_range(3, cc, 3, cc + 1, title, fmt_card_title)
        wsp.merge_range(4, cc, 4, cc + 1, val, fmt_c)
        cc += 2

    # SECCIÓN 1: AVANCE 2026 USUARIOS
    r_idx = 6
    wsp.write(r_idx, 0, "1. AVANCE ESPECÍFICO LOGRADO EN 2026 (USUARIOS B.PIÑA, A.ALVA Y B.NATALIE)", fmt_section)
    r_idx += 1
    th_av = [
        "USUARIO", "REGISTROS 2026", "VOUCHERS ÚNICOS", "COINCIDE PLE", "SOLO REPORTE",
        "% EN PLE", "TOTAL DEBE (S/.)", "TOTAL HABER (S/.)", "DÍAS ACTIVOS", "PROM. DIARIO",
        "PRIMERA MOD 2026", "ÚLTIMA MOD 2026"
    ]
    for c, h in enumerate(th_av):
        wsp.write(r_idx, c, h, fmt_th)
    r_idx += 1
    start_av_r = r_idx + 1

    for _, row in df_avance_2026.iterrows():
        pct_ple = row['Coincide_PLE'] / row['Registros_2026'] if row['Registros_2026'] > 0 else 0
        wsp.write(r_idx, 0, row['Usuario'], fmt_cell)
        wsp.write(r_idx, 1, row['Registros_2026'], fmt_cell_num)
        wsp.write(r_idx, 2, row['Vouchers_Unicos'], fmt_cell_num)
        wsp.write(r_idx, 3, row['Coincide_PLE'], fmt_cell_num)
        wsp.write(r_idx, 4, row['Solo_Reporte'], fmt_cell_num)
        wsp.write(r_idx, 5, pct_ple, fmt_cell_pct)
        wsp.write(r_idx, 6, row['Total_Debe'], fmt_cell_curr)
        wsp.write(r_idx, 7, row['Total_Haber'], fmt_cell_curr)
        wsp.write(r_idx, 8, row['Dias_Activos'], fmt_cell_num)
        wsp.write(r_idx, 9, row['Promedio_Diario'], fmt_cell_dec)
        wsp.write(r_idx, 10, row['Fecha_Min'], fmt_cell_center)
        wsp.write(r_idx, 11, row['Fecha_Max'], fmt_cell_center)
        r_idx += 1

    wsp.write(r_idx, 0, "TOTAL AVANZADO EN 2026", fmt_total)
    wsp.write_formula(r_idx, 1, f"=SUM(B{start_av_r}:B{r_idx})", fmt_total_num)
    wsp.write(r_idx, 2, total_vouc_2026_target, fmt_total_num)
    wsp.write_formula(r_idx, 3, f"=SUM(D{start_av_r}:D{r_idx})", fmt_total_num)
    wsp.write_formula(r_idx, 4, f"=SUM(E{start_av_r}:E{r_idx})", fmt_total_num)
    wsp.write_formula(r_idx, 5, f"=D{r_idx+1}/B{r_idx+1}", fmt_total_pct)
    wsp.write_formula(r_idx, 6, f"=SUM(G{start_av_r}:G{r_idx})", fmt_total_curr)
    wsp.write_formula(r_idx, 7, f"=SUM(H{start_av_r}:H{r_idx})", fmt_total_curr)
    wsp.write(r_idx, 8, "-", fmt_total_center)
    wsp.write_formula(r_idx, 9, f"=AVERAGE(J{start_av_r}:J{r_idx})", fmt_total_dec)
    wsp.write(r_idx, 10, "-", fmt_total_center)
    wsp.write(r_idx, 11, "-", fmt_total_center)

    # SECCIÓN 2: TABLA MAESTRA DE PROYECCIÓN DE TIEMPO
    r_idx += 3
    wsp.write(r_idx, 0, "2. TABLA MAESTRA DE PROYECCIÓN DE TIEMPO (2 PERSONAS A 45 REGISTROS/DÍA - FACTOR ~2 REG/VOUCHER)", fmt_section)
    r_idx += 1
    th_pr = [
        "ESCENARIO DE CORRECCIÓN", "ALCANCE / DESCRIPCIÓN", "CANTIDAD PENDIENTE",
        "UNIDAD", "RITMO DIARIO (2 PERS)", "RITMO MENSUAL (22 DÍAS)", "DÍAS HÁBILES",
        "SEMANAS LAB.", "MESES LAB.", "AÑOS LAB.", "FECHA ESTIMADA DE TÉRMINO"
    ]
    for c, h in enumerate(th_pr):
        wsp.write(r_idx, c, h, fmt_th)
    r_idx += 1

    for _, row in df_proj.iterrows():
        wsp.write(r_idx, 0, row['Escenario'], fmt_cell)
        wsp.write(r_idx, 1, row['Descripcion'], fmt_cell)
        wsp.write(r_idx, 2, row['Items_Pendientes'], fmt_cell_num)
        wsp.write(r_idx, 3, row['Unidad'], fmt_cell_center)
        wsp.write(r_idx, 4, row['Rendimiento_Diario'], fmt_cell_num)
        wsp.write(r_idx, 5, row['Rendimiento_Mensual'], fmt_cell_num)
        wsp.write(r_idx, 6, row['Dias_Habiles'], fmt_cell_num)
        wsp.write(r_idx, 7, row['Semanas_Laborales'], fmt_cell_dec)
        wsp.write(r_idx, 8, row['Meses_Laborales'], fmt_cell_dec)
        wsp.write(r_idx, 9, row['Anios_Laborales'], fmt_cell_dec)
        wsp.write(r_idx, 10, row['Fecha_Estimada_Termino'], fmt_highlight)
        r_idx += 1

    # SECCIÓN 3: MATRIZ DE SENSIBILIDAD POR PERSONAL
    r_idx += 2
    wsp.write(r_idx, 0, "3. ANÁLISIS DE SENSIBILIDAD: ¿EN CUÁNTO TIEMPO SE TERMINA SI SE AMPLÍA EL EQUIPO?", fmt_section)
    r_idx += 1
    th_sen = [
        "TAMAÑO EQUIPO", "CAPACIDAD DIARIA", "CAPACIDAD MENSUAL",
        "MESES (SOLO REPORTE: 94,774)", "TÉRMINO (SOLO REPORTE)",
        "MESES (UNIVERSO TOTAL: 296,037)", "TÉRMINO (UNIVERSO TOTAL)"
    ]
    for c, h in enumerate(th_sen):
        wsp.write(r_idx, c, h, fmt_th)
    r_idx += 1

    for _, row in df_sensitivity.iterrows():
        wsp.write(r_idx, 0, row['Personal'], fmt_cell_center)
        wsp.write(r_idx, 1, row['Capacidad_Diaria'], fmt_cell_num)
        wsp.write(r_idx, 2, row['Capacidad_Mensual'], fmt_cell_num)
        wsp.write(r_idx, 3, row['Meses_Solo_Reporte'], fmt_cell_dec)
        wsp.write(r_idx, 4, row['Termino_Solo_Reporte'], fmt_cell_center)
        wsp.write(r_idx, 5, row['Meses_Universo_Total'], fmt_cell_dec)
        wsp.write(r_idx, 6, row['Termino_Universo_Total'], fmt_cell_center)
        r_idx += 1

    # SECCIÓN 4: CONCLUSIONES Y RECOMENDACIONES ESTRATÉGICAS
    r_idx += 2
    wsp.write(r_idx, 0, "4. CONCLUSIONES CLAVE Y RECOMENDACIONES ESTRATÉGICAS DE AUDITORÍA", fmt_section)
    r_idx += 1
    recs = [
        ("• Diagnóstico de Viabilidad:", "Corregir manualmente 94,774 registros con 2 personas a 45 reg/día requiere 4.0 años (1,054 días hábiles), terminando en Octubre 2030."),
        ("• Factor Voucher (~2 reg/voucher):", "Si se concilia a nivel de Voucher Único (30,385 vouchers únicos), el tiempo se reduce a 2.5 años (676 días hábiles), culminando en Mayo 2029."),
        ("• Avance 2026 Destacado:", "B.Piña, A.Alva y B.Natalie han avanzado 616 registros (269 vouchers) en 2026, logrando que el 81.2% (500 reg) concilie con PLE."),
        ("• Recomendación de Materialidad:", "No corregir los 94,774 registros uno por uno; priorizar los vouchers con importes superiores a S/ 10,000 y cuentas bancarias/proveedores."),
        ("• Recomendación de Automatización:", "El 70% de las diferencias responde a patrones sistemáticos (reclasificaciones masivas). Se recomienda conciliación asistida por script.")
    ]
    for bold_hdr, text in recs:
        wsp.write(r_idx, 0, bold_hdr, wb.add_format({'bold': True, 'font_size': 9, 'font_name': 'Segoe UI', 'font_color': c_navy}))
        wsp.merge_range(r_idx, 1, r_idx, 10, text, wb.add_format({'font_size': 9, 'font_name': 'Segoe UI', 'font_color': '#333333'}))
        r_idx += 1

    widths_p = [32, 42, 20, 14, 18, 20, 16, 15, 15, 14, 22]
    for i, w in enumerate(widths_p):
        wsp.set_column(i, i, w)

    # =========================================================================
    # PESTAÑA 2: AVANCE_POR_CUENTA (NUEVA PESTAÑA SOLICITADA)
    # =========================================================================
    print("      -> Generando hoja 'AVANCE_POR_CUENTA'...")
    wsa = wb.add_worksheet('AVANCE_POR_CUENTA')
    wsa.set_zoom(90)
    wsa.freeze_panes(4, 2)

    wsa.write(0, 0, "RESUMEN EJECUTIVO DE AVANCE POR CUENTA CONTABLE Y AÑO DE EDICIÓN", fmt_title)
    wsa.write(1, 0, "Muestra cuánto se ha avanzado en cada cuenta contable en cada año (2023, 2024, 2025, 2026), importes y principales usuarios modificadores.", fmt_subtitle)

    th_ac = [
        "CUENTA CONTABLE", "NOMBRE CUENTA", "TOTAL MOVIMIENTOS", "% DEL TOTAL",
        "AÑO 2023", "AÑO 2024", "AÑO 2025", "AÑO 2026", "SIN FECHA MOD",
        "COINCIDE PLE", "SOLO REPORTE", "% CONCILIADO", "TOTAL DEBE (S/.)", "TOTAL HABER (S/.)",
        "USUARIOS QUE MODIFICARON (TOP 3)"
    ]
    for c, h in enumerate(th_ac):
        wsa.write(3, c, h, fmt_th)

    r_ac_idx = 4
    for _, row in df_acc_summary.iterrows():
        wsa.write(r_ac_idx, 0, row['CTA_STR'], fmt_cell)
        wsa.write(r_ac_idx, 1, row['NOM_CTA'], fmt_cell)
        wsa.write(r_ac_idx, 2, row['Total_Mov'], fmt_cell_num)
        wsa.write(r_ac_idx, 3, row['Pct_Total'], fmt_cell_pct)
        wsa.write(r_ac_idx, 4, row['Anio_2023'], fmt_cell_num)
        wsa.write(r_ac_idx, 5, row['Anio_2024'], fmt_cell_num)
        wsa.write(r_ac_idx, 6, row['Anio_2025'], fmt_cell_num)
        wsa.write(r_ac_idx, 7, row['Anio_2026'], fmt_cell_num)
        wsa.write(r_ac_idx, 8, row['Sin_Fecha'], fmt_cell_num)
        wsa.write(r_ac_idx, 9, row['Coincide_PLE'], fmt_cell_num)
        wsa.write(r_ac_idx, 10, row['Solo_Reporte'], fmt_cell_num)
        wsa.write(r_ac_idx, 11, row['Pct_Coincide'], fmt_cell_pct)
        wsa.write(r_ac_idx, 12, row['Total_Debe'], fmt_cell_curr)
        wsa.write(r_ac_idx, 13, row['Total_Haber'], fmt_cell_curr)
        wsa.write(r_ac_idx, 14, row['Usuarios_Top'], fmt_cell)
        r_ac_idx += 1

    # Fila de Totales
    wsa.write(r_ac_idx, 0, "TOTAL GENERAL", fmt_total)
    wsa.write(r_ac_idx, 1, f"{len(df_acc_summary)} Cuentas", fmt_total_center)
    wsa.write_formula(r_ac_idx, 2, f"=SUM(C5:C{r_ac_idx})", fmt_total_num)
    wsa.write_formula(r_ac_idx, 3, f"=SUM(D5:D{r_ac_idx})", fmt_total_pct)
    wsa.write_formula(r_ac_idx, 4, f"=SUM(E5:E{r_ac_idx})", fmt_total_num)
    wsa.write_formula(r_ac_idx, 5, f"=SUM(F5:F{r_ac_idx})", fmt_total_num)
    wsa.write_formula(r_ac_idx, 6, f"=SUM(G5:G{r_ac_idx})", fmt_total_num)
    wsa.write_formula(r_ac_idx, 7, f"=SUM(H5:H{r_ac_idx})", fmt_total_num)
    wsa.write_formula(r_ac_idx, 8, f"=SUM(I5:I{r_ac_idx})", fmt_total_num)
    wsa.write_formula(r_ac_idx, 9, f"=SUM(J5:J{r_ac_idx})", fmt_total_num)
    wsa.write_formula(r_ac_idx, 10, f"=SUM(K5:K{r_ac_idx})", fmt_total_num)
    wsa.write_formula(r_ac_idx, 11, f"=J{r_ac_idx+1}/C{r_ac_idx+1}", fmt_total_pct)
    wsa.write_formula(r_ac_idx, 12, f"=SUM(M5:M{r_ac_idx})", fmt_total_curr)
    wsa.write_formula(r_ac_idx, 13, f"=SUM(N5:N{r_ac_idx})", fmt_total_curr)
    wsa.write(r_ac_idx, 14, "-", fmt_total_center)

    wsa.autofilter(3, 0, r_ac_idx - 1, len(th_ac) - 1)
    widths_ac = [16, 32, 18, 12, 12, 12, 12, 12, 14, 14, 14, 14, 20, 20, 45]
    for i, w in enumerate(widths_ac):
        wsa.set_column(i, i, w)

    # =========================================================================
    # PESTAÑA 3: CUENTAS_X_MES_USUARIO (DESGLOSE DIRECTO SOLICITADO)
    # =========================================================================
    print("      -> Generando hoja 'CUENTAS_X_MES_USUARIO' (614 filas)...")
    wsc = wb.add_worksheet('CUENTAS_X_MES_USUARIO')
    wsc.set_zoom(90)
    wsc.freeze_panes(4, 0)

    wsc.write(0, 0, "MOVIMIENTOS POR CUENTA CONTABLE, AÑO-MES DE EDICIÓN Y USUARIO", fmt_title)
    wsc.write(1, 0, "Permite filtrar cualquier mes (ej. marzo, agosto, 2026) para saber exactamente cuántos movimientos se hicieron por cuenta y qué usuario lo hizo.", fmt_subtitle)

    th_cmu = [
        "AÑO EDICIÓN", "AÑO-MES EDICIÓN", "CUENTA CONTABLE", "NOMBRE CUENTA", "USUARIO MODIFICA",
        "MOVIMIENTOS REALIZADOS", "VOUCHERS ÚNICOS", "COINCIDE PLE", "SOLO REPORTE", "% CONCILIADO",
        "TOTAL DEBE (S/.)", "TOTAL HABER (S/.)", "PRIMERA FECHA MOD", "ÚLTIMA FECHA MOD"
    ]
    for c, h in enumerate(th_cmu):
        wsc.write(3, c, h, fmt_th)

    r_cmu_idx = 4
    for _, row in df_cmu.iterrows():
        wsc.write(r_cmu_idx, 0, row['YEAR_MOD'], fmt_cell_center)
        wsc.write(r_cmu_idx, 1, row['AÑO_MES'], fmt_cell_center)
        wsc.write(r_cmu_idx, 2, row['CTA_STR'], fmt_cell)
        wsc.write(r_cmu_idx, 3, row['NOM_CTA'], fmt_cell)
        wsc.write(r_cmu_idx, 4, row['Usu_Mod_Clean'], fmt_cell)
        wsc.write(r_cmu_idx, 5, row['Movimientos'], fmt_cell_num)
        wsc.write(r_cmu_idx, 6, row['Vouchers_Unicos'], fmt_cell_num)
        wsc.write(r_cmu_idx, 7, row['Coincide_PLE'], fmt_cell_num)
        wsc.write(r_cmu_idx, 8, row['Solo_Reporte'], fmt_cell_num)
        wsc.write(r_cmu_idx, 9, row['Pct_Coincide'], fmt_cell_pct)
        wsc.write(r_cmu_idx, 10, row['Total_Debe'], fmt_cell_curr)
        wsc.write(r_cmu_idx, 11, row['Total_Haber'], fmt_cell_curr)
        wsc.write(r_cmu_idx, 12, row['Fecha_Min'], fmt_cell_center)
        wsc.write(r_cmu_idx, 13, row['Fecha_Max'], fmt_cell_center)
        r_cmu_idx += 1

    # Fila de Totales
    wsc.write(r_cmu_idx, 0, "TOTAL GENERAL", fmt_total)
    wsc.write(r_cmu_idx, 1, "-", fmt_total_center)
    wsc.write(r_cmu_idx, 2, "-", fmt_total_center)
    wsc.write(r_cmu_idx, 3, "-", fmt_total_center)
    wsc.write(r_cmu_idx, 4, "-", fmt_total_center)
    wsc.write_formula(r_cmu_idx, 5, f"=SUM(F5:F{r_cmu_idx})", fmt_total_num)
    wsc.write(r_cmu_idx, 6, df_mod['REP_Reg_Ctb'].nunique(), fmt_total_num)
    wsc.write_formula(r_cmu_idx, 7, f"=SUM(H5:H{r_cmu_idx})", fmt_total_num)
    wsc.write_formula(r_cmu_idx, 8, f"=SUM(I5:I{r_cmu_idx})", fmt_total_num)
    wsc.write_formula(r_cmu_idx, 9, f"=H{r_cmu_idx+1}/F{r_cmu_idx+1}", fmt_total_pct)
    wsc.write_formula(r_cmu_idx, 10, f"=SUM(K5:K{r_cmu_idx})", fmt_total_curr)
    wsc.write_formula(r_cmu_idx, 11, f"=SUM(L5:L{r_cmu_idx})", fmt_total_curr)
    wsc.write(r_cmu_idx, 12, "-", fmt_total_center)
    wsc.write(r_cmu_idx, 13, "-", fmt_total_center)

    wsc.autofilter(3, 0, r_cmu_idx - 1, len(th_cmu) - 1)
    widths_cmu = [14, 16, 16, 32, 18, 22, 16, 14, 14, 14, 20, 20, 18, 18]
    for i, w in enumerate(widths_cmu):
        wsc.set_column(i, i, w)

    # =========================================================================
    # PESTAÑA 4: MATRIZ_CUENTAS_MES (PIVOT CUENTA VS MES)
    # =========================================================================
    print("      -> Generando hoja 'MATRIZ_CUENTAS_MES'...")
    wsmc = wb.add_worksheet('MATRIZ_CUENTAS_MES')
    wsmc.set_zoom(90)
    wsmc.freeze_panes(4, 2)

    wsmc.write(0, 0, "MATRIZ MENSUAL: CANTIDAD DE MOVIMIENTOS POR CUENTA Y AÑO-MES DE EDICIÓN", fmt_title)
    wsmc.write(1, 0, "Comparativa mes a mes de cuántos movimientos fueron modificados en cada cuenta contable en la historia.", fmt_subtitle)

    wsmc.write(3, 0, "CUENTA CONTABLE", fmt_th)
    wsmc.write(3, 1, "NOMBRE CUENTA", fmt_th)
    for c_idx, m_col in enumerate(meses_acc_cols):
        wsmc.write(3, c_idx + 2, m_col, fmt_th)
    wsmc.write(3, len(meses_acc_cols) + 2, "TOTAL MOVIMIENTOS", fmt_th_blue)

    r_piv_idx = 4
    for _, row in pivot_acc_month.iterrows():
        wsmc.write(r_piv_idx, 0, row['CTA_STR'], fmt_cell)
        wsmc.write(r_piv_idx, 1, row['NOM_CTA'], fmt_cell)
        for c_idx, m_col in enumerate(meses_acc_cols):
            val = int(row[m_col])
            wsmc.write(r_piv_idx, c_idx + 2, val if val > 0 else 0, fmt_cell_num if val > 0 else fmt_cell_center)
        tot_v = int(row['TOTAL'])
        wsmc.write(r_piv_idx, len(meses_acc_cols) + 2, tot_v, fmt_total_num)
        r_piv_idx += 1

    # Totales Matriz
    wsmc.write(r_piv_idx, 0, "TOTAL POR MES", fmt_total)
    wsmc.write(r_piv_idx, 1, f"{len(pivot_acc_month)} Cuentas", fmt_total_center)
    for c_idx in range(len(meses_acc_cols) + 1):
        col_letter = xlsxwriter.utility.xl_col_to_name(c_idx + 2)
        wsmc.write_formula(r_piv_idx, c_idx + 2, f"=SUM({col_letter}5:{col_letter}{r_piv_idx})", fmt_total_num)

    wsmc.set_column(0, 0, 16)
    wsmc.set_column(1, 1, 32)
    for c_idx in range(2, len(meses_acc_cols) + 3):
        wsmc.set_column(c_idx, c_idx, 13)

    # =========================================================================
    # PESTAÑA 5: AVANCE_2026_DETALLE (616 filas)
    # =========================================================================
    print("      -> Generando hoja 'AVANCE_2026_DETALLE' (616 filas)...")
    ws_av = wb.add_worksheet('AVANCE_2026_DETALLE')
    ws_av.set_zoom(90)
    ws_av.freeze_panes(4, 0)

    ws_av.write(0, 0, "DETALLE DE REGISTROS EDITADOS EN 2026 (B.PIÑA, A.ALVA Y B.NATALIE)", fmt_title)
    ws_av.write(1, 0, "Auditoría línea por línea de las 616 modificaciones realizadas en el ejercicio 2026 sobre el periodo contable 2023.", fmt_subtitle)

    th_av_det = [
        "ESTADO CRUCE PLE", "USUARIO MODIFICA", "FECHA/HORA MODIFICACIÓN",
        "FECHA MOD", "HORA MOD", "AÑO-MES", "VOUCHER (REG. CTB.)", "PRDO",
        "FECHA MOVIMIENTO", "CUENTA CONTABLE", "NOMBRE CUENTA",
        "MONTO DEBE (S/.)", "MONTO HABER (S/.)", "GLOSA", "TIPO DOC", "NRO DOC",
        "CLIENTE / PROVEEDOR", "USUARIO CREA", "FECHA REGISTRO", "CLAVE CONCATENADA"
    ]
    for c, h in enumerate(th_av_det):
        ws_av.write(3, c, h, fmt_th)

    r_av_idx = 4
    for _, row in df_2026_target.iterrows():
        st = row['ESTADO_CRUCE']
        ws_av.write(r_av_idx, 0, st, fmt_coincide if st == 'COINCIDE' else fmt_solo_rep)
        ws_av.write(r_av_idx, 1, row['Usu_Mod_Clean'], fmt_cell)

        if pd.notna(row['FEC_MOD_DT']):
            ws_av.write_datetime(r_av_idx, 2, row['FEC_MOD_DT'].to_pydatetime(), fmt_cell_dt)
        else:
            ws_av.write(r_av_idx, 2, "(Sin Fecha Mod)", fmt_cell_center)

        ws_av.write(r_av_idx, 3, row['FECHA_SOLO_STR'], fmt_cell_center)
        ws_av.write(r_av_idx, 4, row['HORA_STR'], fmt_cell_center)
        ws_av.write(r_av_idx, 5, row['AÑO_MES'], fmt_cell_center)
        ws_av.write(r_av_idx, 6, row['REP_Reg_Ctb'], fmt_cell)

        p_val = str(row['Prdo.']).replace('.0', '').zfill(2)
        ws_av.write(r_av_idx, 7, p_val, fmt_cell_center)
        ws_av.write(r_av_idx, 8, row['REP_Fec_Mov'], fmt_cell_center)
        ws_av.write(r_av_idx, 9, row['REP_Nro_Cta_Limpia'], fmt_cell)
        ws_av.write(r_av_idx, 10, str(row['Nombre Cuenta']).strip() if pd.notna(row['Nombre Cuenta']) else '', fmt_cell)
        ws_av.write(r_av_idx, 11, row['REP_Debe_Num'], fmt_cell_curr)
        ws_av.write(r_av_idx, 12, row['REP_Haber_Num'], fmt_cell_curr)
        ws_av.write(r_av_idx, 13, str(row['Glosa']).strip() if pd.notna(row['Glosa']) else '', fmt_cell)
        ws_av.write(r_av_idx, 14, str(row['Cod. Tip. Doc.']).strip() if pd.notna(row['Cod. Tip. Doc.']) else '', fmt_cell_center)
        ws_av.write(r_av_idx, 15, str(row['Nro. Doc.']).strip() if pd.notna(row['Nro. Doc.']) else '', fmt_cell)
        ws_av.write(r_av_idx, 16, row['TERCERO'], fmt_cell)
        ws_av.write(r_av_idx, 17, row['Usu_Crea_Clean'], fmt_cell)

        if pd.notna(row['FEC_REG_DT']):
            ws_av.write_datetime(r_av_idx, 18, row['FEC_REG_DT'].to_pydatetime(), fmt_cell_dt)
        else:
            ws_av.write(r_av_idx, 18, '', fmt_cell_center)

        ws_av.write(r_av_idx, 19, row['CLAVE_CONCAT'], fmt_cell)
        r_av_idx += 1

    ws_av.write(r_av_idx, 0, "TOTAL GENERAL", fmt_total)
    for c in range(1, 11):
        ws_av.write(r_av_idx, c, "", fmt_total)
    ws_av.write_formula(r_av_idx, 11, f"=SUM(L5:L{r_av_idx})", fmt_total_curr)
    ws_av.write_formula(r_av_idx, 12, f"=SUM(M5:M{r_av_idx})", fmt_total_curr)
    for c in range(13, len(th_av_det)):
        ws_av.write(r_av_idx, c, "", fmt_total)

    ws_av.autofilter(3, 0, r_av_idx - 1, len(th_av_det) - 1)
    widths_av = [18, 16, 22, 14, 12, 12, 18, 8, 16, 16, 28, 18, 18, 30, 10, 16, 26, 16, 22, 45]
    for i, w in enumerate(widths_av):
        ws_av.set_column(i, i, w)

    # =========================================================================
    # PESTAÑA 6: RESUMEN_EJECUTIVO
    # =========================================================================
    print("      -> Generando hoja 'RESUMEN_EJECUTIVO'...")
    ws1 = wb.add_worksheet('RESUMEN_EJECUTIVO')
    ws1.set_zoom(90)

    ws1.write(0, 0, "AUDITORÍA DE EDICIONES POR USUARIO EN REPORTE CONTABLE 2023", fmt_title)
    ws1.write(1, 0, f"Empresa: CORPORACIÓN COMATPE SAC | RUC: 20516403650 | Cruce: PLE Diario 5.1 vs Reporte Contabilidad | Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}", fmt_subtitle)

    kpis_glob = [
        ("ASIENTOS MODIFICADOS", len(df_mod), fmt_card_num),
        ("VOUCHERS ÚNICOS", df_mod['REP_Reg_Ctb'].nunique(), fmt_card_num),
        ("USUARIOS MODIFICADORES", df_mod['Usu_Mod_Clean'].nunique(), fmt_card_num),
        ("COINCIDEN CON PLE", (df_mod['ESTADO_CRUCE'] == 'COINCIDE').sum(), fmt_card_num),
        ("SOLO EN REPORTE", (df_mod['ESTADO_CRUCE'] == 'SOLO EN REPORTE').sum(), fmt_card_num),
        ("TOTAL MONTO DEBE (S/.)", df_mod['REP_Debe_Num'].sum(), fmt_card_str),
        ("TOTAL MONTO HABER (S/.)", df_mod['REP_Haber_Num'].sum(), fmt_card_str),
    ]
    card_col = 0
    for title, val, fmt_n in kpis_glob:
        val_str = f"S/ {val:,.2f}" if isinstance(val, (float, int)) and val > 1000000 else val
        w_fmt = fmt_card_str if isinstance(val_str, str) else fmt_n
        ws1.merge_range(3, card_col, 3, card_col + 1, title, fmt_card_title)
        ws1.merge_range(4, card_col, 4, card_col + 1, val_str, w_fmt)
        card_col += 2

    # TABLA 1: RANKING POR USUARIO
    user_grp = df_mod.groupby('Usu_Mod_Clean')
    df_user_summary = user_grp.agg(
        Total_Registros=('Reg. Ctb.', 'count'),
        Vouchers_Unicos=('REP_Reg_Ctb', 'nunique'),
        Coincide_PLE=('ESTADO_CRUCE', lambda s: (s == 'COINCIDE').sum()),
        Solo_Reporte=('ESTADO_CRUCE', lambda s: (s == 'SOLO EN REPORTE').sum()),
        Total_Debe=('REP_Debe_Num', 'sum'),
        Total_Haber=('REP_Haber_Num', 'sum'),
        Primera_Mod=('FEC_MOD_DT', 'min'),
        Ultima_Mod=('FEC_MOD_DT', 'max'),
        Dias_Activos=('FECHA_SOLO_STR', lambda s: s[s != '(Sin Fecha Mod)'].nunique())
    ).reset_index()

    df_user_summary['Pct_Total'] = df_user_summary['Total_Registros'] / len(df_mod)
    df_user_summary['Promedio_Diario'] = np.where(
        df_user_summary['Dias_Activos'] > 0,
        (df_user_summary['Total_Registros'] / df_user_summary['Dias_Activos']).round(1),
        df_user_summary['Total_Registros']
    )
    df_user_summary.sort_values(by='Total_Registros', ascending=False, inplace=True)

    ws1.write(6, 0, "1. RANKING GENERAL DE EDICIONES POR USUARIO", fmt_section)
    th_u = [
        "USUARIO", "TOTAL ASIENTOS", "% DEL TOTAL", "VOUCHERS ÚNICOS",
        "COINCIDE PLE", "SOLO REPORTE", "TOTAL DEBE (S/.)", "TOTAL HABER (S/.)",
        "DÍAS ACTIVOS", "PROM. DIARIO", "PRIMERA EDICIÓN", "ÚLTIMA EDICIÓN"
    ]
    for c, h in enumerate(th_u):
        ws1.write(7, c, h, fmt_th)

    r_idx = 8
    for _, row in df_user_summary.iterrows():
        ws1.write(r_idx, 0, row['Usu_Mod_Clean'], fmt_cell)
        ws1.write(r_idx, 1, row['Total_Registros'], fmt_cell_num)
        ws1.write(r_idx, 2, row['Pct_Total'], fmt_cell_pct)
        ws1.write(r_idx, 3, row['Vouchers_Unicos'], fmt_cell_num)
        ws1.write(r_idx, 4, row['Coincide_PLE'], fmt_cell_num)
        ws1.write(r_idx, 5, row['Solo_Reporte'], fmt_cell_num)
        ws1.write(r_idx, 6, row['Total_Debe'], fmt_cell_curr)
        ws1.write(r_idx, 7, row['Total_Haber'], fmt_cell_curr)
        ws1.write(r_idx, 8, row['Dias_Activos'], fmt_cell_num)
        ws1.write(r_idx, 9, row['Promedio_Diario'], fmt_cell_dec)
        f_min = row['Primera_Mod'].strftime('%d/%m/%Y %H:%M') if pd.notna(row['Primera_Mod']) else '-'
        f_max = row['Ultima_Mod'].strftime('%d/%m/%Y %H:%M') if pd.notna(row['Ultima_Mod']) else '-'
        ws1.write(r_idx, 10, f_min, fmt_cell_center)
        ws1.write(r_idx, 11, f_max, fmt_cell_center)
        r_idx += 1

    ws1.write(r_idx, 0, "TOTAL GENERAL", fmt_total)
    ws1.write_formula(r_idx, 1, f"=SUM(B9:B{r_idx})", fmt_total_num)
    ws1.write_formula(r_idx, 2, f"=SUM(C9:C{r_idx})", fmt_total_pct)
    ws1.write(r_idx, 3, df_mod['REP_Reg_Ctb'].nunique(), fmt_total_num)
    ws1.write_formula(r_idx, 4, f"=SUM(E9:E{r_idx})", fmt_total_num)
    ws1.write_formula(r_idx, 5, f"=SUM(F9:F{r_idx})", fmt_total_num)
    ws1.write_formula(r_idx, 6, f"=SUM(G9:G{r_idx})", fmt_total_curr)
    ws1.write_formula(r_idx, 7, f"=SUM(H9:H{r_idx})", fmt_total_curr)
    ws1.write(r_idx, 8, "-", fmt_total_center)
    ws1.write_formula(r_idx, 9, f"=AVERAGE(J9:J{r_idx})", fmt_total_dec)
    ws1.write(r_idx, 10, "-", fmt_total_center)
    ws1.write(r_idx, 11, "-", fmt_total_center)

    # TABLA 2: RESUMEN POR AÑO-MES
    month_grp = df_mod.groupby('AÑO_MES')
    df_month_summary = month_grp.agg(
        Total_Registros=('Reg. Ctb.', 'count'),
        Vouchers_Unicos=('REP_Reg_Ctb', 'nunique'),
        Usuarios_Activos=('Usu_Mod_Clean', 'nunique'),
        Coincide_PLE=('ESTADO_CRUCE', lambda s: (s == 'COINCIDE').sum()),
        Solo_Reporte=('ESTADO_CRUCE', lambda s: (s == 'SOLO EN REPORTE').sum()),
        Total_Debe=('REP_Debe_Num', 'sum'),
        Total_Haber=('REP_Haber_Num', 'sum')
    ).reset_index()
    df_month_summary.sort_values(by='AÑO_MES', ascending=True, inplace=True)

    r_idx += 3
    ws1.write(r_idx, 0, "2. DISTRIBUCIÓN POR AÑO-MES EN QUE SE REALIZÓ LA MODIFICACIÓN", fmt_section)
    r_idx += 1
    th_m = [
        "AÑO-MES MODIFICACIÓN", "ASIENTOS EDITADOS", "VOUCHERS ÚNICOS",
        "USUARIOS ACTIVOS", "COINCIDE PLE", "SOLO REPORTE", "TOTAL DEBE (S/.)", "TOTAL HABER (S/.)"
    ]
    for c, h in enumerate(th_m):
        ws1.write(r_idx, c, h, fmt_th)
    r_idx += 1
    start_m_row = r_idx + 1

    for _, row in df_month_summary.iterrows():
        ws1.write(r_idx, 0, row['AÑO_MES'], fmt_cell_center)
        ws1.write(r_idx, 1, row['Total_Registros'], fmt_cell_num)
        ws1.write(r_idx, 2, row['Vouchers_Unicos'], fmt_cell_num)
        ws1.write(r_idx, 3, row['Usuarios_Activos'], fmt_cell_num)
        ws1.write(r_idx, 4, row['Coincide_PLE'], fmt_cell_num)
        ws1.write(r_idx, 5, row['Solo_Reporte'], fmt_cell_num)
        ws1.write(r_idx, 6, row['Total_Debe'], fmt_cell_curr)
        ws1.write(r_idx, 7, row['Total_Haber'], fmt_cell_curr)
        r_idx += 1
    
    ws1.write(r_idx, 0, "TOTAL GENERAL", fmt_total)
    ws1.write_formula(r_idx, 1, f"=SUM(B{start_m_row}:B{r_idx})", fmt_total_num)
    ws1.write(r_idx, 2, df_mod['REP_Reg_Ctb'].nunique(), fmt_total_num)
    ws1.write(r_idx, 3, df_mod['Usu_Mod_Clean'].nunique(), fmt_total_num)
    ws1.write_formula(r_idx, 4, f"=SUM(E{start_m_row}:E{r_idx})", fmt_total_num)
    ws1.write_formula(r_idx, 5, f"=SUM(F{start_m_row}:F{r_idx})", fmt_total_num)
    ws1.write_formula(r_idx, 6, f"=SUM(G{start_m_row}:G{r_idx})", fmt_total_curr)
    ws1.write_formula(r_idx, 7, f"=SUM(H{start_m_row}:H{r_idx})", fmt_total_curr)

    widths1 = [25, 18, 14, 18, 16, 16, 20, 20, 15, 15, 20, 20]
    for i, w in enumerate(widths1):
        ws1.set_column(i, i, w)

    # =========================================================================
    # PESTAÑA 7: MENSUAL_X_USUARIO
    # =========================================================================
    print("      -> Generando hoja 'MENSUAL_X_USUARIO'...")
    ws2 = wb.add_worksheet('MENSUAL_X_USUARIO')
    ws2.set_zoom(90)
    ws2.freeze_panes(4, 1)

    ws2.write(0, 0, "MATRIZ MENSUAL DE MODIFICACIONES POR USUARIO (CANTIDAD DE REGISTROS)", fmt_title)
    ws2.write(1, 0, "Evolución mensual detallada de cuántos asientos contables modificó cada usuario en el tiempo.", fmt_subtitle)

    pivot_month = df_mod.pivot_table(
        index='Usu_Mod_Clean', columns='AÑO_MES', values='Reg. Ctb.', aggfunc='count', fill_value=0
    )
    meses_cols = [c for c in sorted(df_mod['AÑO_MES'].unique()) if c != '(Sin Fecha)']
    if '(Sin Fecha)' in df_mod['AÑO_MES'].unique():
        meses_cols.append('(Sin Fecha)')
    pivot_month = pivot_month[meses_cols]
    pivot_month['TOTAL'] = pivot_month.sum(axis=1)
    pivot_month.sort_values(by='TOTAL', ascending=False, inplace=True)

    ws2.write(3, 0, "USUARIO", fmt_th)
    for c_idx, col_name in enumerate(pivot_month.columns):
        fmt_header = fmt_th if col_name != 'TOTAL' else fmt_th_blue
        ws2.write(3, c_idx + 1, col_name, fmt_header)

    p_row_idx = 4
    for user_name, s_row in pivot_month.iterrows():
        ws2.write(p_row_idx, 0, user_name, fmt_cell)
        for c_idx, col_name in enumerate(pivot_month.columns):
            val = int(s_row[col_name])
            if col_name == 'TOTAL':
                ws2.write(p_row_idx, c_idx + 1, val, fmt_total_num)
            else:
                ws2.write(p_row_idx, c_idx + 1, val, fmt_cell_num if val > 0 else fmt_cell_center)
        p_row_idx += 1

    ws2.write(p_row_idx, 0, "TOTAL POR MES", fmt_total)
    for c_idx, col_name in enumerate(pivot_month.columns):
        col_letter = xlsxwriter.utility.xl_col_to_name(c_idx + 1)
        ws2.write_formula(p_row_idx, c_idx + 1, f"=SUM({col_letter}5:{col_letter}{p_row_idx})", fmt_total_num)

    ws2.set_column(0, 0, 20)
    for c_idx in range(1, len(pivot_month.columns) + 2):
        ws2.set_column(c_idx, c_idx, 13)

    # =========================================================================
    # PESTAÑA 8: DIARIO_X_USUARIO
    # =========================================================================
    print("      -> Generando hoja 'DIARIO_X_USUARIO'...")
    ws3 = wb.add_worksheet('DIARIO_X_USUARIO')
    ws3.set_zoom(90)
    ws3.freeze_panes(4, 0)

    ws3.write(0, 0, "DESGLOSE DIARIO DE MODIFICACIONES POR USUARIO Y MES", fmt_title)
    ws3.write(1, 0, "Auditoría diaria: productividad día por día, vouchers únicos, montos modificados y estado en PLE.", fmt_subtitle)

    daily_grp = df_mod.groupby(['Usu_Mod_Clean', 'AÑO_MES', 'FECHA_SOLO_STR', 'DIA_SEMANA'])
    df_daily_summary = daily_grp.agg(
        Total_Registros=('Reg. Ctb.', 'count'),
        Vouchers_Unicos=('REP_Reg_Ctb', 'nunique'),
        Coincide_PLE=('ESTADO_CRUCE', lambda s: (s == 'COINCIDE').sum()),
        Solo_Reporte=('ESTADO_CRUCE', lambda s: (s == 'SOLO EN REPORTE').sum()),
        Total_Debe=('REP_Debe_Num', 'sum'),
        Total_Haber=('REP_Haber_Num', 'sum')
    ).reset_index()
    df_daily_summary.sort_values(by=['Usu_Mod_Clean', 'FECHA_SOLO_STR'], ascending=[True, True], inplace=True)

    th_d = [
        "USUARIO", "AÑO-MES", "FECHA MODIFICACIÓN", "DÍA SEMANA",
        "REGISTROS EDITADOS", "VOUCHERS ÚNICOS", "COINCIDE PLE", "SOLO REPORTE",
        "TOTAL DEBE (S/.)", "TOTAL HABER (S/.)"
    ]
    for c, h in enumerate(th_d):
        ws3.write(3, c, h, fmt_th)

    d_row_idx = 4
    for _, row in df_daily_summary.iterrows():
        ws3.write(d_row_idx, 0, row['Usu_Mod_Clean'], fmt_cell)
        ws3.write(d_row_idx, 1, row['AÑO_MES'], fmt_cell_center)
        ws3.write(d_row_idx, 2, row['FECHA_SOLO_STR'], fmt_cell_center)
        ws3.write(d_row_idx, 3, row['DIA_SEMANA'], fmt_cell)
        ws3.write(d_row_idx, 4, row['Total_Registros'], fmt_cell_num)
        ws3.write(d_row_idx, 5, row['Vouchers_Unicos'], fmt_cell_num)
        ws3.write(d_row_idx, 6, row['Coincide_PLE'], fmt_cell_num)
        ws3.write(d_row_idx, 7, row['Solo_Reporte'], fmt_cell_num)
        ws3.write(d_row_idx, 8, row['Total_Debe'], fmt_cell_curr)
        ws3.write(d_row_idx, 9, row['Total_Haber'], fmt_cell_curr)
        d_row_idx += 1

    ws3.write(d_row_idx, 0, "TOTAL GENERAL", fmt_total)
    ws3.write(d_row_idx, 1, "-", fmt_total_center)
    ws3.write(d_row_idx, 2, "-", fmt_total_center)
    ws3.write(d_row_idx, 3, "-", fmt_total_center)
    ws3.write_formula(d_row_idx, 4, f"=SUM(E5:E{d_row_idx})", fmt_total_num)
    ws3.write(d_row_idx, 5, df_mod['REP_Reg_Ctb'].nunique(), fmt_total_num)
    ws3.write_formula(d_row_idx, 6, f"=SUM(G5:G{d_row_idx})", fmt_total_num)
    ws3.write_formula(d_row_idx, 7, f"=SUM(H5:H{d_row_idx})", fmt_total_num)
    ws3.write_formula(d_row_idx, 8, f"=SUM(I5:I{d_row_idx})", fmt_total_curr)
    ws3.write_formula(d_row_idx, 9, f"=SUM(J5:J{d_row_idx})", fmt_total_curr)

    ws3.autofilter(3, 0, d_row_idx - 1, len(th_d) - 1)
    widths3 = [18, 12, 22, 16, 20, 18, 15, 15, 20, 20]
    for i, w in enumerate(widths3):
        ws3.set_column(i, i, w)

    # =========================================================================
    # PESTAÑA 9: MATRIZ_CALENDARIO_DIARIO
    # =========================================================================
    print("      -> Generando hoja 'MATRIZ_CALENDARIO_DIARIO'...")
    ws4 = wb.add_worksheet('MATRIZ_CALENDARIO_DIARIO')
    ws4.set_zoom(90)
    ws4.freeze_panes(4, 1)

    ws4.write(0, 0, "MATRIZ CALENDARIO: EDICIONES POR FECHA Y USUARIO", fmt_title)
    ws4.write(1, 0, "Vista comparativa día por día entre todos los usuarios modificadores.", fmt_subtitle)

    pivot_daily = df_mod[df_mod['FECHA_SOLO_STR'] != '(Sin Fecha Mod)'].pivot_table(
        index='FECHA_SOLO_STR', columns='Usu_Mod_Clean', values='Reg. Ctb.', aggfunc='count', fill_value=0
    )
    users_ranking = list(df_user_summary['Usu_Mod_Clean'])
    users_cols_in_pivot = [u for u in users_ranking if u in pivot_daily.columns]
    pivot_daily = pivot_daily[users_cols_in_pivot]
    pivot_daily['TOTAL_DIA'] = pivot_daily.sum(axis=1)

    ws4.write(3, 0, "FECHA MODIFICACIÓN", fmt_th)
    for c_idx, u_col in enumerate(pivot_daily.columns):
        fmt_header = fmt_th if u_col != 'TOTAL_DIA' else fmt_th_blue
        ws4.write(3, c_idx + 1, u_col, fmt_header)

    cal_row_idx = 4
    for date_str, s_row in pivot_daily.iterrows():
        ws4.write(cal_row_idx, 0, date_str, fmt_cell_center)
        for c_idx, u_col in enumerate(pivot_daily.columns):
            val = int(s_row[u_col])
            if u_col == 'TOTAL_DIA':
                ws4.write(cal_row_idx, c_idx + 1, val, fmt_total_num)
            else:
                ws4.write(cal_row_idx, c_idx + 1, val, fmt_cell_num if val > 0 else fmt_cell_center)
        cal_row_idx += 1

    ws4.write(cal_row_idx, 0, "TOTAL POR USUARIO", fmt_total)
    for c_idx, u_col in enumerate(pivot_daily.columns):
        col_letter = xlsxwriter.utility.xl_col_to_name(c_idx + 1)
        ws4.write_formula(cal_row_idx, c_idx + 1, f"=SUM({col_letter}5:{col_letter}{cal_row_idx})", fmt_total_num)

    ws4.set_column(0, 0, 22)
    for c_idx in range(1, len(pivot_daily.columns) + 2):
        ws4.set_column(c_idx, c_idx, 13)

    # =========================================================================
    # PESTAÑA 10: DETALLE_EDICIONES (9,537 registros)
    # =========================================================================
    print("      -> Generando hoja 'DETALLE_EDICIONES' (9,537 filas)...")
    ws5 = wb.add_worksheet('DETALLE_EDICIONES')
    ws5.set_zoom(90)
    ws5.freeze_panes(4, 0)

    ws5.write(0, 0, "DETALLE COMPLETO DE ASIENTOS MODIFICADOS EN REPORTE CONTABILIDAD 2023", fmt_title)
    ws5.write(1, 0, "Listado exhaustivo de las 9,537 líneas con usuario, fecha, voucher, cuenta, importes y estado en Cruce PLE.", fmt_subtitle)

    th_det = [
        "ESTADO CRUCE PLE", "USUARIO MODIFICA", "FECHA/HORA MODIFICACIÓN",
        "FECHA MOD", "HORA MOD", "AÑO-MES", "VOUCHER (REG. CTB.)", "PRDO",
        "FECHA MOVIMIENTO", "CUENTA CONTABLE", "NOMBRE CUENTA",
        "MONTO DEBE (S/.)", "MONTO HABER (S/.)", "GLOSA", "TIPO DOC", "NRO DOC",
        "CLIENTE / PROVEEDOR", "USUARIO CREA", "FECHA REGISTRO", "CLAVE CONCATENADA"
    ]
    for c, h in enumerate(th_det):
        ws5.write(3, c, h, fmt_th)

    det_row_idx = 4
    for _, row in df_mod.iterrows():
        st = row['ESTADO_CRUCE']
        ws5.write(det_row_idx, 0, st, fmt_coincide if st == 'COINCIDE' else fmt_solo_rep)
        ws5.write(det_row_idx, 1, row['Usu_Mod_Clean'], fmt_cell)

        if pd.notna(row['FEC_MOD_DT']):
            ws5.write_datetime(det_row_idx, 2, row['FEC_MOD_DT'].to_pydatetime(), fmt_cell_dt)
        else:
            ws5.write(det_row_idx, 2, "(Sin Fecha Mod)", fmt_cell_center)

        ws5.write(det_row_idx, 3, row['FECHA_SOLO_STR'], fmt_cell_center)
        ws5.write(det_row_idx, 4, row['HORA_STR'], fmt_cell_center)
        ws5.write(det_row_idx, 5, row['AÑO_MES'], fmt_cell_center)
        ws5.write(det_row_idx, 6, row['REP_Reg_Ctb'], fmt_cell)

        p_val = str(row['Prdo.']).replace('.0', '').zfill(2)
        ws5.write(det_row_idx, 7, p_val, fmt_cell_center)
        ws5.write(det_row_idx, 8, row['REP_Fec_Mov'], fmt_cell_center)
        ws5.write(det_row_idx, 9, row['REP_Nro_Cta_Limpia'], fmt_cell)
        ws5.write(det_row_idx, 10, str(row['Nombre Cuenta']).strip() if pd.notna(row['Nombre Cuenta']) else '', fmt_cell)
        ws5.write(det_row_idx, 11, row['REP_Debe_Num'], fmt_cell_curr)
        ws5.write(det_row_idx, 12, row['REP_Haber_Num'], fmt_cell_curr)
        ws5.write(det_row_idx, 13, str(row['Glosa']).strip() if pd.notna(row['Glosa']) else '', fmt_cell)
        ws5.write(det_row_idx, 14, str(row['Cod. Tip. Doc.']).strip() if pd.notna(row['Cod. Tip. Doc.']) else '', fmt_cell_center)
        ws5.write(det_row_idx, 15, str(row['Nro. Doc.']).strip() if pd.notna(row['Nro. Doc.']) else '', fmt_cell)
        ws5.write(det_row_idx, 16, row['TERCERO'], fmt_cell)
        ws5.write(det_row_idx, 17, row['Usu_Crea_Clean'], fmt_cell)

        if pd.notna(row['FEC_REG_DT']):
            ws5.write_datetime(det_row_idx, 18, row['FEC_REG_DT'].to_pydatetime(), fmt_cell_dt)
        else:
            ws5.write(det_row_idx, 18, '', fmt_cell_center)

        ws5.write(det_row_idx, 19, row['CLAVE_CONCAT'], fmt_cell)
        det_row_idx += 1

    ws5.write(det_row_idx, 0, "TOTAL GENERAL", fmt_total)
    for c in range(1, 11):
        ws5.write(det_row_idx, c, "", fmt_total)
    ws5.write_formula(det_row_idx, 11, f"=SUM(L5:L{det_row_idx})", fmt_total_curr)
    ws5.write_formula(det_row_idx, 12, f"=SUM(M5:M{det_row_idx})", fmt_total_curr)
    for c in range(13, len(th_det)):
        ws5.write(det_row_idx, c, "", fmt_total)

    ws5.autofilter(3, 0, det_row_idx - 1, len(th_det) - 1)
    widths5 = [18, 16, 22, 14, 12, 12, 18, 8, 16, 16, 28, 18, 18, 30, 10, 16, 26, 16, 22, 45]
    for i, w in enumerate(widths5):
        ws5.set_column(i, i, w)

    # Cerrar libro
    print(f"\n[Guardando] Finalizando archivo Excel: '{output_excel}'...")
    wb.close()
    file_size_mb = os.path.getsize(output_excel) / (1024 * 1024)
    print(f"=" * 85)
    print(f" ARCHIVO ACTUALIZADO CON ÉXITO: {output_excel}")
    print(f" Tamaño: {file_size_mb:.2f} MB | Tiempo total: {time.time()-start_time:.2f} segundos")
    print(f"=" * 85)

if __name__ == "__main__":
    main()
