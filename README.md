# Facturación Emitida y Recibida — JOCS (Carga Aérea)

Tablero Streamlit de facturación, rentabilidad por aduana, modelo de
originación y nómina, para la operación de carga aérea (CARGA AEREA
CONSOLIDADA). Construido a partir de `Dashboard Carga-Aérea - Requerimientos
y Simulación Visual.docx`.

## Contenido

| Página | Qué muestra |
|---|---|
| Utilidad Operativa | Cascada Ingresos→Honorarios AA→Sueldos→Gastos→Utilidad Operativa, barras + ficha de detalle por aduana, Pareto de Utilidad Operativa (todo CC), Costos de Originación (Honorarios Josefa + Costos de Referencia) |
| Gastos por Concepto | Dona de gasto (o ingreso) por concepto, una aduana a la vez, con drilldown de proveedores/clientes |
| Nómina | Treemap de 3 niveles con drill-down: aduanas → empleados de la aduana → componentes de su percepción |
| Tarjeta de Crédito | Pendiente de datos — necesita hojas de movimientos AMEX (Alejandro y Josefa) que aún no están en el Sheets |

Reestructurado a estas 4 pestañas el 2026-10-13 siguiendo un PDF de referencia
del usuario (antes había 6: Panorama, Gastos por Concepto, Por Aduana
/Rentabilidad, Modelo y Originación, Nómina, Detalle de Facturas — las 3
últimas se eliminaron por decisión explícita).

## Gráficas: Recharts, no Altair

Las gráficas se renderizan con **Recharts** (React/JS), no con Vega-Altair —
cambio explícito a pedido del usuario (2026-10-13), solo en este tablero (los
otros 3 de Reportes Pao siguen en Altair). No hay paquete de Streamlit ya hecho
para esto: se incrusta a mano con `st.components.v1.html()`, cargando
React/ReactDOM/Recharts desde unpkg y armando el árbol con `React.createElement`
(sin JSX, para no tener que compilar con Babel en el navegador en cada rerun).
Helpers en `app_facturacion.py`: `recharts_component()`, `waterfall_recharts()`,
`bar_h_recharts()`, `bar_h_signed_recharts()`, `donut_recharts()`,
`pareto_recharts()`, `treemap_recharts()` (este último usa el componente nativo
`Treemap` de Recharts). Cada gráfica vive en su propio iframe aislado, sin
estado compartido con Streamlit ni entre gráficas.

## Correr en local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Datos

**En vivo desde Google Sheets** (desde 2026-10-12; antes: Excel local
commiteado en `EXCEL/`). El tablero descarga el libro completo cada 5 minutos
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

Se usan 7 de las pestañas del libro: `emitidas`, `recibidas`, `sueldos`,
`honorarios josefa`, `costos de referencia`, `UTILIDAD OPERATIVA`. El resto
(`corresponsalias` -- ya no se usa tras el rediseño a 4 pestañas --, `reglas`,
`CC`, `Aduanas`, `imile laredo`) no se usa todavía.

Cosas que hay que saber:

1. **El monto oficial es `MONTO`/`monto 1`, no `Importe`.** Hay líneas con
   importe pero sin monto (sobre todo IMSS/Infonavit/honorarios en STAFF) que
   el usuario deja fuera del cierre a propósito — no se suman a las cifras.
2. **Utilidad Operativa = Ingresos + Honorarios AA − Sueldos − Gastos**
   (confirmado con el usuario; Honorarios AA es ingreso adicional, no costo).
   El Pareto de esta página ordena por **magnitud absoluta** (no por el valor
   con signo) ya que la utilidad puede ser negativa por aduana — un % acumulado
   no tiene sentido sobre una suma que cruza por cero.
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
