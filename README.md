# Facturación Emitida y Recibida — JOCS (Carga Aérea)

Tablero Streamlit de facturación, rentabilidad por aduana, modelo de
originación y nómina, para la operación de carga aérea (CARGA AEREA
CONSOLIDADA). Construido a partir de `Dashboard Carga-Aérea - Requerimientos
y Simulación Visual.docx`.

## Contenido

| Página | Qué muestra |
|---|---|
| Panorama | KPIs, emitida/recibida por mes, centros de costo, detalle dinámico de conceptos, Pareto 80/20 de aduanas |
| Gastos por Concepto | Dona de gasto (o ingreso) por concepto, una aduana a la vez, con drilldown de proveedores/clientes |
| Por Aduana (Rentabilidad) | Matriz de 4 cuadrantes, comparativo emitida/recibida y margen por aduana, detalle por centro de costo |
| Modelo y Originación | Corresponsalías vs. asesoría aduanal, honorarios de Josefa, costos de referencia |
| Nómina | Nómina real por empleado (nombre y sueldo), con cascada de costo por trabajador |
| Detalle de Facturas | Buscador línea por línea, emitidas y recibidas |

## Correr en local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Datos

**Desde 2026-10-12, en vivo desde Google Sheets** (antes: Excel local commiteado
en `EXCEL/`). El tablero descarga el libro completo cada 5 minutos
(`_descargar_libro_sheets()`, `st.cache_data(ttl=300)`) desde:

```
https://docs.google.com/spreadsheets/d/1F-b12XUlASAXNnhgYm5G-PRB46s2Ms9RjJ3j6Rj_QVU/export?format=xlsx
```

Esto funciona sin API key ni cuenta de servicio porque la hoja está compartida
como **"Cualquiera con el link" (solo lectura)** — si algún día se quita ese
permiso, el tablero deja de poder descargarla. Se intentó primero una cuenta de
servicio de Google Cloud, pero la organización `jocs.com.mx` bloquea la
creación de llaves de cuentas de servicio por política; de ahí el cambio a
link público de solo lectura.

Solo se usan 6 de las pestañas del libro: `emitidas`, `recibidas`, `sueldos`,
`corresponsalias`, `honorarios josefa`, `costos de referencia`. El resto
(`UTILIDAD OPERATIVA`, `reglas`, `CC`, `Aduanas`, `imile laredo`) no se usa
todavía.

Cosas que hay que saber:

1. **El monto oficial es `MONTO`/`monto 1`, no `Importe`.** Hay líneas con
   importe pero sin monto (sobre todo IMSS/Infonavit/honorarios en STAFF) que
   el usuario deja fuera del cierre a propósito — no se suman a las cifras.
2. **"Ingreso Asesoría" en Modelo y Originación es una suposición**: el
   documento de requerimientos no deja claro qué hoja la alimenta. Se usa
   `costos de referencia`; con `honorarios josefa` el margen salía muy
   negativo e implausible. Ver el aviso en la propia página.
3. **La cascada de nómina incluye "Excedente"**, que no aparece en la maqueta
   del documento pero sí es parte real del total (`TOTAL2`).
4. **"Cuotas Patronales" en Nómina está pendiente de revisión** (aviso visible
   en la propia pestaña): al conectar a Sheets se detectó que esa columna llega
   sistemáticamente más baja que en el Excel original que se uso para armar el
   tablero por primera vez (la proporción varía por empleado, no es un factor
   fijo) — probablemente una fórmula que no se recalculó igual al pasar el
   archivo a Sheets. El resto de los 9 componentes de la cascada sí coincide
   exacto contra ese Excel.
5. `COLS_NOMINA`/`ALIAS_NOMINA` en el código toleran que "Total cuotas patrón
   (N)" y "Comisión (lira)" lleguen sin acento (`patron`, `Comision`) — así
   quedaron tras corregir un acento que había llegado dañado (carácter de
   reemplazo) en esas dos cabeceras de la hoja `sueldos`.

## Confidencial

Contiene nombre y sueldo real de cada empleado (pestaña Nómina), además de
facturación y márgenes de la operación. Publicado como repo público por
decisión explícita del usuario (Streamlit Community Cloud gratis solo permite
1 app privada, ya ocupada por `tablero-ancora`).
