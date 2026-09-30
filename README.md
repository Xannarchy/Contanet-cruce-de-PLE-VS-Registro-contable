# Contanet / Starsoft - Cruce de PLE Diario vs Registro Contable

Herramienta de auditoría y conciliación contable automatizada entre el **Reporte de Contabilidad** (exportado de Starsoft / ERP) y el **PLE Libro Diario 5.1** declarado a SUNAT.

## Versión 2.1 (Novedades)
- **Soporte Multianual Completo (2022, 2023, 2024):** Compatible con cualquier ejercicio contable vía `--year` o scripts dedicados.
- **Detección Automática de Archivos y Hojas:** Localiza automáticamente los archivos en la carpeta de descargas y detecta las hojas consolidadas (`'AÑO 2024 '`, `'AÑO 2023'`, `'2022 ANUAL'`, `'Contabilidad'`).
- **Lectura Ultra-Rápida con Rust:** Uso del motor `python-calamine` para cargar hojas de 400,000+ filas en ~15-30 segundos.
- **Normalización Contable Robusta:** Limpieza de cuentas (eliminación de puntos, ej. `10.1.1.01` -> `101101`), normalización de fechas (con protección contra valores nulos o timestamp con horas) y montos formateados a 2 decimales.
- **Cruce 1 a 1 sin Duplicación Cartesiana:** Asignación de índices de ocurrencia enteros garantizados (`#0`, `#1`, `#2`...) para conciliar asientos con líneas idénticas sin explosión de filas.
- **Exportación Streaming con `xlsxwriter`:** Genera archivos Excel de más de 500,000 filas con uso mínimo de memoria RAM y 3 pestañas estructuradas.

---

## Estructura del Archivo Excel Generado
1. **`RESUMEN`**: Resumen ejecutivo con totales de registros, sumatorias de Debe/Haber y tabla comparativa mes a mes con fila de Total General.
2. **`DIFERENCIAS`**: Registros que no cuadraron (`SOLO EN REPORTE` y `SOLO EN PLE`), ordenados y con formato condicional para auditoría inmediata.
3. **`CRUCE_COMPLETO`**: Full Outer Join total con todas las columnas de ambos sistemas y la columna de estado al inicio.

---

## Uso

### Requisitos previos
```bash
pip install pandas python-calamine xlsxwriter
```

### Ejecución para año 2024
```bash
python cruzar_contabilidad_ple.py --year 2024
# O alternativamente:
python cruzar_contabilidad_ple_2024.py
```

### Ejecución para año 2023
```bash
python cruzar_contabilidad_ple.py --year 2023
```

### Ejecución para año 2022
```bash
python cruzar_contabilidad_ple.py --year 2022
```

---

## Regla de Conciliación
$$\text{Clave} = \text{Reg. Ctb.} + \text{"\|"} + \text{Fec. Mov.} + \text{"\|"} + \text{Nro. Cta. (sin puntos)} + \text{"\|"} + \text{Mto. Debe} + \text{"\|"} + \text{Mto. Haber} + \text{"\#"} + \text{Ocurrencia}$$
