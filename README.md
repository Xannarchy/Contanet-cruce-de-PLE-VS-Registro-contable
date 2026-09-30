# Contanet / Starsoft - Cruce de PLE Diario vs Registro Contable

Herramienta de auditoría y conciliación contable automatizada entre el **Reporte de Contabilidad** (exportado de Starsoft / ERP) y el **PLE Libro Diario 5.1** declarado a SUNAT.

## Versión 2.0 (Novedades)
- **Soporte Multianual Dinámico:** Compatible con años 2022, 2023 y futuros vía parámetro `--year`.
- **Detección Automática de Archivos y Hojas:** Localiza automáticamente los archivos en la carpeta de descargas y detecta las hojas consolidadas (`'2022 ANUAL'`, `'AÑO 2023'`, `'Contabilidad'`).
- **Lectura Ultra-Rápida con Rust:** Uso del motor `python-calamine` para cargar hojas de 400,000+ filas en ~10-20 segundos.
- **Normalización Contable Inteligente:** Limpieza de cuentas (eliminación de puntos, ej. `10.1.1.01` -> `101101`) y formato estandarizado de montos y fechas.
- **Cruce 1 a 1 sin Duplicación Cartesiana:** Asignación de índices de ocurrencia para conciliar asientos con líneas idénticas sin explosión de filas.
- **Exportación Streaming con `xlsxwriter`:** Genera archivos Excel de más de 500,000 filas con uso mínimo de memoria RAM y 3 pestañas estructuradas.

---

## Estructura del Archivo Excel Generado
1. **`RESUMEN`**: Resumen ejecutivo con totales de registros, sumatorias de Debe/Haber y tabla comparativa mes a mes.
2. **`DIFERENCIAS`**: Registros que no cuadraron (`SOLO EN REPORTE` y `SOLO EN PLE`), ordenados y con formato condicional para auditoría inmediata.
3. **`CRUCE_COMPLETO`**: Full Outer Join total con todas las columnas de ambos sistemas y la columna de estado al inicio.

---

## Uso

### Requisitos previos
```bash
pip install pandas python-calamine xlsxwriter
```

### Ejecución para año 2022
```bash
python cruzar_contabilidad_ple.py --year 2022
```

### Ejecución para año 2023
```bash
python cruzar_contabilidad_ple.py --year 2023
```

---

## Regla de Conciliación
$$\text{Clave} = \text{Reg. Ctb.} + \text{"\|"} + \text{Fec. Mov.} + \text{"\|"} + \text{Nro. Cta. (sin puntos)} + \text{"\|"} + \text{Mto. Debe} + \text{"\|"} + \text{Mto. Haber} + \text{"\#"} + \text{Ocurrencia}$$
