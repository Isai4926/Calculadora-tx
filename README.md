# 🧮 Calculadora Tx

Aplicación web (Streamlit) para analizar facturas emitidas, gastos (facturas
recibidas), movimientos bancarios y retenciones de ISR: calcula ingresos
facturados vs. cobrados, gastos deducibles/no deducibles, IVA trasladado /
acreditable / neto, ISR retenido y concilia facturas contra movimientos
bancarios, generando alertas de validación y un reporte Excel descargable.

> ⚠️ **Esta herramienta es un auxiliar de análisis.** Los resultados deben
> ser revisados por un contador público o especialista fiscal antes de
> presentar declaraciones o tomar decisiones fiscales.

---

## 1. Instalación

Requiere Python 3.9 o superior.

```bash
# 1. Clona el repositorio
git clone https://github.com/tu-usuario/calculadora-tx.git
cd calculadora-tx

# 2. (Recomendado) crea un entorno virtual
python -m venv venv
source venv/bin/activate        # En Windows: venv\Scripts\activate

# 3. Instala las dependencias
pip install -r requirements.txt
```

## 2. Ejecutar la aplicación

```bash
streamlit run app.py
```

Se abrirá automáticamente en tu navegador en `http://localhost:8501`.
Desde la barra lateral podrás cargar las 4 bases de datos y elegir el
periodo y filtros a analizar.

## 3. Estructura de las 4 bases de datos

La app acepta archivos **.xlsx** o **.csv**. Los nombres de columna no
distinguen mayúsculas/minúsculas ni espacios extra, pero deben estar
presentes con estos nombres:

### a) Facturas emitidas
| Columna | Descripción |
|---|---|
| `folio` | Identificador único de la factura (p. ej. F-0001) |
| `fecha` | Fecha de emisión |
| `cliente` | Nombre del cliente |
| `concepto` | Concepto facturado |
| `subtotal` | Importe antes de IVA |
| `iva` | IVA trasladado |
| `total` | Subtotal + IVA |

### b) Gastos (facturas recibidas)
| Columna | Descripción |
|---|---|
| `folio_gasto` | Identificador único del gasto (p. ej. G-0001) |
| `fecha` | Fecha del gasto |
| `proveedor` | Nombre del proveedor |
| `concepto` | Categoría del gasto |
| `subtotal` | Importe antes de IVA |
| `iva` | IVA pagado |
| `total` | Subtotal + IVA |
| `comprobante` | `Si` / `No` — si cuenta con comprobante fiscal (CFDI) |

### c) Movimientos bancarios
| Columna | Descripción |
|---|---|
| `fecha` | Fecha del movimiento |
| `monto` | Importe (positivo para ingresos, negativo para egresos) |
| `tipo` | `Ingreso`, `Cobro`, `Deposito`, `Abono` (cobros) o `Egreso` (pagos) |
| `referencia` | Folio de factura/gasto relacionado, o una clave propia del banco |

### d) Retenciones de ISR
| Columna | Descripción |
|---|---|
| `fecha` | Fecha de la retención |
| `referencia` | Clave interna de la retención |
| `concepto` | Concepto de la retención |
| `isr_retenido` | Importe retenido |

Si falta alguna columna obligatoria, la app lo indica claramente (archivo y
columna) y detiene el análisis hasta que se corrija.

## 4. Cómo funcionan los cálculos y la conciliación

### Clasificación de gastos (deducible / no deducible)
1. Conceptos como *Donativo no autorizado, Gasto personal, Recargo* o
   *Multa* se marcan **No deducible** sin importar si hay comprobante.
2. Cualquier gasto sin comprobante fiscal (`comprobante = No`) se marca
   **No deducible**, sin importar el concepto.
3. Un gasto con comprobante y un concepto reconocido como deducible
   (Transporte, Publicidad, Renta de oficina, etc.) se marca **Deducible**.
4. Un concepto que no está catalogado en ninguna de las reglas anteriores
   se marca **Pendiente de clasificación** — la app nunca asume
   automáticamente si algo desconocido es o no deducible.

### IVA
- **IVA trasladado** = suma del IVA de todas las facturas emitidas.
- **IVA acreditable** = suma del IVA únicamente de los gastos clasificados
  como *Deducibles*.
- **IVA neto** = IVA trasladado − IVA acreditable.
- Se marca como *inconsistencia de IVA* cualquier registro donde el IVA no
  corresponda a subtotal × 16% (con tolerancia de $1 peso por redondeos).

### Conciliación de facturas contra movimientos bancarios
Se busca, para cada factura, un movimiento bancario de tipo ingreso cuya
`referencia` sea igual al `folio` de la factura. Tolerancias: **±$10.00**
en importe y **±5 días naturales** en fecha.

- Si no hay ningún movimiento con esa referencia → **No conciliado
  (pendiente de cobro)**.
- Si hay un solo movimiento y coincide exacto → **Conciliado**.
- Si hay un solo movimiento dentro de tolerancia pero no exacto →
  **Conciliado con diferencia dentro de tolerancia**.
- Si hay **más de un** movimiento con la misma referencia (posible pago
  dividido o duplicado), se suman los montos: si la suma cae dentro de
  tolerancia se marca **Conciliado (pago dividido)**; si no, **Posible
  duplicado / revisión manual** (la app nunca decide sola cuál movimiento
  es el correcto).
- Cualquier otro caso fuera de tolerancia → **Requiere revisión manual**.

Los movimientos de ingreso cuya referencia no corresponde a ninguna
factura se listan aparte como **posibles ingresos no facturados**. Los
egresos (pagos) cuya referencia no corresponde a ningún `folio_gasto` se
listan como **egresos sin gasto relacionado**: la app no intenta adivinar
a qué gasto pertenecen ni suma facturas de gasto para "hacerlas calzar".

### Retenciones de ISR
Las retenciones no traen folio de factura, UUID ni RFC, y no existe una
proporción fija y confiable entre el monto retenido y el subtotal de
ninguna factura cercana en fecha. Por eso **la app no vincula
automáticamente cada retención a una factura**: reporta el total del
periodo y el detalle completo por separado, marcado como pendiente de
conciliación manual. Si tu base sí incluye folio de factura, RFC o UUID en
el futuro, ese campo puede usarse para una vinculación exacta.

### Alertas
Cada hallazgo se clasifica en prioridad **Alta**, **Media** o **Baja** y
cubre, entre otros: facturas sin cobro, depósitos sin factura, egresos sin
gasto, diferencias fuera de tolerancia, duplicados, datos faltantes,
importes en cero o negativos, gastos sin clasificar, inconsistencias de
IVA y retenciones de ISR sin vincular.

## 5. Supuestos y limitaciones de esta versión

- Los datos se asumen en **pesos mexicanos (MXN)**; no hay manejo de
  moneda extranjera ni tipo de cambio (las bases de origen no incluyen esa
  columna).
- No se contemplan notas de crédito, devoluciones ni cancelaciones (no
  aparecen en la estructura de datos definida).
- La conciliación usa folio/referencia + tolerancias de importe y fecha,
  no RFC ni UUID (esas columnas no forman parte de las 4 bases definidas).
  Si tus archivos reales sí las incluyen, se puede extender fácilmente la
  función `conciliar_facturas_cobros()` en `app.py` para usarlas como
  criterio adicional o preferente.

## 6. Reporte Excel descargable

Desde la pestaña **Resumen ejecutivo** puedes descargar un Excel con las
hojas: Resumen ejecutivo, Ingresos facturados y cobrados, Gastos
deducibles, Gastos no deducibles, IVA (facturas y gastos), Retenciones
ISR, Conciliación bancaria, Partidas no conciliadas, Depósitos sin
factura, Egresos sin gasto, Alertas y validaciones, y los datos originales
procesados. Las columnas calculadas siempre están separadas de las
columnas originales del archivo fuente, para conservar trazabilidad.

## 7. Publicar en GitHub

```bash
git init
git add app.py requirements.txt README.md
git commit -m "Calculadora Tx: primera versión"
git branch -M main
git remote add origin https://github.com/tu-usuario/calculadora-tx.git
git push -u origin main
```

## 8. Desplegar en Streamlit Community Cloud

1. Sube el repositorio a GitHub (paso anterior).
2. Entra a [share.streamlit.io](https://share.streamlit.io) e inicia sesión
   con tu cuenta de GitHub.
3. Haz clic en **"New app"**, selecciona el repositorio, la rama (`main`)
   y el archivo principal `app.py`.
4. Haz clic en **"Deploy"**. En unos minutos tendrás una URL pública para
   compartir la app.
5. Cada vez que hagas `git push` a la rama configurada, Streamlit Cloud
   actualizará la app automáticamente.

## 9. Estructura del proyecto

```
calculadora-tx/
├── app.py            # Aplicación completa (carga, cálculos, alertas, tablero, reporte)
├── requirements.txt  # Dependencias de Python
└── README.md         # Este archivo
```

El código de `app.py` está dividido en 6 secciones claramente comentadas
(configuración, carga/validación, cálculo, alertas, reporte Excel e
interfaz), pensadas para que puedas ubicar y modificar cada regla de
negocio sin tener que leer todo el archivo.
