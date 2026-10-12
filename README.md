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

`EXCEL/base de prueba agosto carga aerea.xlsx`. Solo se usan 6 de sus 10 hojas:
`emitidas`, `recibidas`, `sueldos`, `corresponsalias`, `honorarios josefa`,
`costos de referencia`. El resto (`reglas`, `CC`, `Aduanas`, `imile laredo`)
no se usa todavía.

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
4. Los Excel que lleguen a futuro deben nombrarse con la palabra **"carga
   aerea"** en el nombre de archivo — así los descubre el tablero solo.

## Confidencial

Contiene nombre y sueldo real de cada empleado (pestaña Nómina), además de
facturación y márgenes de la operación. Publicado como repo público por
decisión explícita del usuario (Streamlit Community Cloud gratis solo permite
1 app privada, ya ocupada por `tablero-ancora`).
