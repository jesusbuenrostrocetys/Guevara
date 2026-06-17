# 📘 SPRINT 1 — Catálogo Maestro · Documentación de Entrega

**Empresa:** SYSEA Radiadores · **Propietario:** Mario Guevara
**Entregable:** `ENTREGAS/SYSEA_Radiadores_SPRINT1.xlsx`
**Hoja:** `Catálogo Maestro` · **Estado:** ✅ Funcional y validado

---

## 1. Resumen

Este Sprint entrega el **Catálogo Maestro** con:

- ✅ Las **21 columnas oficiales** (A → U) con los nombres exactos del briefing.
- ✅ **20 productos ficticios** de validación.
- ✅ **6 fórmulas automáticas** (precios, utilidades, rotación, alerta AutoZone).
- ✅ **6 validaciones de datos** (DPI único, precios/stock no negativos, dropdown de Estado).
- ✅ **Formato condicional** en Stock, Rotación y Utilidad Cliente.
- ✅ Verificado con un motor de fórmulas real: **todos los precios terminan en 00/50**,
  utilidad de cliente **≥ $1,200** y de taller **≥ $800** en los 20 productos.

> El archivo se puede **regenerar** en cualquier momento con
> `python scripts/generar_catalogo_sprint1.py` (requiere `pip install openpyxl`).

---

## 2. Dos correcciones importantes respecto al prompt

Para que el archivo funcione de verdad en Excel, ajusté dos cosas **sin tocar la
política de precios ni la estructura de columnas**:

### 2.1 Referencia de "Precio Compra": `C2` → `F2`
En el prompt, las fórmulas usaban `C2` para *Precio Compra*. Pero en la estructura
oficial de 21 columnas, *Precio Compra* es la **columna F** (es la columna #6). La `C`
venía de un borrador anterior. Remapeé `C2 → F2` para que la fórmula apunte a la celda
correcta. *AutoZone* (`G`), *Cliente Final* (`H`) y *Última Venta* (`Q`) ya estaban bien.

### 2.2 Función de redondeo con su 2º argumento
`REDONDEAR.MAS` (ROUNDUP) **exige dos argumentos**: el número y los decimales.
El prompt mostraba `REDONDEAR.MAS((C2+1500)/50)*50` (le faltaba el `;0`/`,0`), lo que
daría error en Excel. Se corrigió a `REDONDEAR.MAS((F2+1500)/50, 0)*50`. El resultado
es el mismo que se buscaba: **redondear hacia arriba al múltiplo de 50**.

### 📝 Nota sobre idioma de las fórmulas (¡no te asustes!)
Todo archivo `.xlsx` guarda las fórmulas **en inglés** internamente (`IF`, `ROUNDUP`,
`TODAY`, `MAX`). **Excel las traduce solas al abrirlas**: en español verás `SI`,
`REDONDEAR.MAS`, `HOY`, `MAX`. Si tu Excel usa punto y coma (`;`) como separador,
también lo ajusta automáticamente. **No tienes que hacer nada.**

---

## 3. Estructura de columnas (A → U)

| Col | Campo | Tipo | Cómo se llena |
|-----|-------|------|---------------|
| A | DPI | Texto | Manual (único) |
| B | Descripción | Texto | Manual |
| C | Marca | Texto | Manual |
| D | Modelo | Texto | Manual |
| E | Año | Número | Manual |
| F | Precio Compra | Moneda MXN | Manual |
| G | Precio AutoZone | Moneda MXN | Manual |
| **H** | **Precio Cliente Final** | Moneda MXN | 🧮 **Fórmula** |
| **I** | **Precio Taller** | Moneda MXN | 🧮 **Fórmula** |
| **J** | **Utilidad Cliente** | Moneda MXN | 🧮 **Fórmula** |
| **K** | **Utilidad Taller** | Moneda MXN | 🧮 **Fórmula** |
| L | Stock Actual | Número | Manual (futuro: Entradas) |
| M | Stock Mínimo | Número | Parámetro (=1) |
| N | Stock Ideal | Número | Parámetro (=3) |
| O | Proveedor Principal | Texto | Manual |
| P | Fecha Última Compra | Fecha | Manual (futuro: Entradas) |
| Q | Fecha Última Venta | Fecha | Manual (futuro: Ventas) |
| R | Estado | Dropdown | Activo / Inactivo |
| **S** | **Rotación** | Texto | 🧮 **Fórmula** |
| **T** | **Alerta AutoZone** | Texto | 🧮 **Fórmula** |
| U | Observaciones | Texto | Manual |

> Las columnas calculadas (H, I, J, K, S, T) tienen el encabezado en **azul claro**.
> En el Sprint 1, *Stock Actual*, *Última Compra* y *Última Venta* se cargan a mano;
> en los Sprints 2 y 3 se vincularán automáticamente a las hojas de Entradas y Ventas.

---

## 4. Explicación detallada de cada fórmula

> Se muestra la versión que **verás en tu Excel en español**. (Internamente el archivo
> las guarda en inglés; Excel hace la traducción al abrir.)

### Fórmula 1 — Precio Cliente Final (columna **H**)

```excel
=SI(G2>0,
    REDONDEAR.MAS(SI(G2-F2-250<1200, F2+1200, G2-250)/50, 0)*50,
    REDONDEAR.MAS((F2+1500)/50, 0)*50)
```

**Lógica paso a paso:**
1. **¿Hay precio de AutoZone (`G2`)?**
   - **Sí:** apunto a quedar **$250 por debajo** de AutoZone (`G2-250`)…
     - …pero si con eso la utilidad quedaría **por debajo de $1,200**
       (`G2-F2-250 < 1200`), entonces subo el precio a `F2+1200` para
       **garantizar la utilidad mínima de $1,200**.
   - **No:** parto de **Precio Compra + $1,500** (utilidad ideal).
2. **Redondeo hacia arriba al múltiplo de 50** con `REDONDEAR.MAS(.../50, 0)*50`.

**Ejemplos reales del archivo:**
- `RVVW029`: AutoZone $3,900 → $3,900−$250 = **$3,650** (utilidad $2,150 ✓).
- `BUICK22`: AutoZone $1,900 → $1,650 daría utilidad $1,050 (< $1,200), así que
  sube a $600+$1,200 = **$1,800** (utilidad exactamente $1,200 ✓).

### Fórmula 2 — Precio Taller (columna **I**)

```excel
=REDONDEAR.MAS(MAX(F2+800, H2*0.82)/50, 0)*50
```

**Lógica:** parte del Precio Cliente Final con un **descuento ≈18%** (`H2*0.82`),
pero `MAX(...)` garantiza que **nunca baje de $800 de utilidad** (`F2+800`).
Luego redondea hacia arriba a múltiplo de 50.

- `RVVW029`: MAX($2,300, $2,993) = $2,993 → **$3,000** (utilidad $1,500 ✓).
- `BUICK22`: MAX($1,400, $1,476) = $1,476 → **$1,500** (utilidad $900 ✓).

### Fórmula 3 — Utilidad Cliente (columna **J**)

```excel
=H2-F2
```
Precio Cliente Final menos Precio Compra.

### Fórmula 4 — Utilidad Taller (columna **K**)

```excel
=I2-F2
```
Precio Taller menos Precio Compra.

### Fórmula 5 — Rotación (columna **S**)

```excel
=SI(Q2="", "Sin venta",
    SI(HOY()-Q2<=90,  "Estrella",
    SI(HOY()-Q2<=180, "Media",
    SI(HOY()-Q2<=365, "Lenta", "Muerta"))))
```

Calcula los **días transcurridos desde la última venta** (`HOY()-Q2`) y clasifica:
**Estrella** ≤90 días · **Media** 91–180 · **Lenta** 181–365 · **Muerta** >365 ·
**Sin venta** si nunca se ha vendido (`Q2` vacío).

### Fórmula 6 — Alerta AutoZone (columna **T**)

```excel
=SI(G2="", "Pendiente", SI(G2>0, "OK", "Revisar"))
```

**Pendiente** si falta el precio de AutoZone · **OK** si hay precio válido ·
**Revisar** si es 0 o inválido.

---

## 5. Validaciones de datos (ya activas)

| # | Columna | Regla | Mensaje de error |
|---|---------|-------|------------------|
| 1 | A — DPI | Sin duplicados (`CONTAR.SI`) | "Ese DPI ya existe…" |
| 2 | F — Precio Compra | Número > 0 | "Debe ser mayor a 0" |
| 3 | L — Stock Actual | Entero ≥ 0 | "No puede ser negativo" |
| 4 | M — Stock Mínimo | Entero ≥ 0 | "No puede ser negativo" |
| 5 | N — Stock Ideal | Entero ≥ 0 | "No puede ser negativo" |
| 6 | R — Estado | Lista: Activo / Inactivo | Dropdown |

> El **redondeo a 00/50** no necesita validación: lo garantizan las fórmulas H e I.
> Las reglas se aplican hasta la fila 1000 para que sigan funcionando al agregar productos.

---

## 6. Formato condicional (colores automáticos)

**Stock Actual (L)** — rojo `=0` · naranja `0<stock<Mín` · amarillo `Mín≤stock<Ideal` · verde `≥Ideal`.
> Con Stock Mínimo = 1, el naranja solo aparecería si subes el mínimo a 2 o más.

**Rotación (S)** — Estrella=verde · Media=amarillo · Lenta=naranja · Muerta=rojo · Sin venta=gris.

**Utilidad Cliente (J)** — rojo `<$1,200` · amarillo `$1,200–$1,499` · verde `≥$1,500`.

---

## 7. Instrucciones de uso

### Para abrir y probar
1. Abre `SYSEA_Radiadores_SPRINT1.xlsx` en **Excel 365**.
2. Si Excel pregunta por actualizar/calcular, acepta (las fórmulas se calculan al abrir).
3. Revisa la hoja `Catálogo Maestro`: las columnas H, I, J, K, S, T ya muestran valores.

### Para agregar un producto nuevo
1. Escribe en la primera fila vacía las columnas **manuales**:
   DPI, Descripción, Marca, Modelo, Año, **Precio Compra (F)**, **Precio AutoZone (G)**,
   Stock, Proveedor, fechas y Estado.
2. **Copia las fórmulas hacia abajo:** selecciona las celdas H:K, S y T de la fila de
   arriba y arrástralas (o copia/pega) a la fila nueva. Se ajustan solas.
3. Los precios, utilidades, rotación y alerta se calculan automáticamente.

### Para editar
- Cambia el **Precio Compra** o el **AutoZone** y los precios sugeridos se recalculan al instante.
- Puedes editar libremente: **las fórmulas no se rompen** (la hoja no está bloqueada).

### Para regenerar el archivo desde cero
```bash
pip install openpyxl
python scripts/generar_catalogo_sprint1.py
```

---

## 8. Pruebas realizadas ✅

Se verificó con un **motor de fórmulas de Excel** (no solo en teoría):

| Prueba | Resultado |
|--------|-----------|
| ¿Precio Cliente Final correcto? | ✅ Sí (20/20) |
| ¿Todos los precios terminan en 00 o 50? | ✅ Sí |
| ¿Utilidad Taller ≥ $800? | ✅ Sí (mín. real: $900) |
| ¿Utilidad Cliente ≥ $1,200? | ✅ Sí (mín. real: $1,200) |
| ¿Rotación se clasifica sola? | ✅ Sí (se incluyen los 5 casos) |
| ¿Formato condicional aplicado? | ✅ Sí (Stock, Rotación, Utilidad) |
| ¿Fórmulas sin errores (#REF!, #DIV/0!)? | ✅ Sí |
| ¿Se puede editar sin romper fórmulas? | ✅ Sí (hoja desbloqueada) |

Tabla de valores calculados (muestra):

| DPI | Compra | AutoZone | Cliente | Taller | Util.Cli | Util.Tall | Rotación |
|-----|-------:|---------:|--------:|-------:|---------:|----------:|----------|
| RVVW029 | $1,500 | $3,900 | $3,650 | $3,000 | $2,150 | $1,500 | Estrella |
| 13015 | $800 | $2,500 | $2,250 | $1,850 | $1,450 | $1,050 | Lenta |
| BUICK22 | $600 | $1,900 | $1,800 | $1,500 | $1,200 | $900 | Muerta |
| MERC77 | $3,500 | $8,900 | $8,650 | $7,100 | $5,150 | $3,600 | Lenta |

---

## 9. ¿Hace falta VBA? (Opcional para Sprint 1)

**No es necesario.** Todo el Sprint 1 funciona con fórmulas nativas + validación +
formato condicional, sin macros. Esto mantiene el archivo como `.xlsx` (sin alertas de
seguridad por macros).

Si más adelante quieres una **alerta emergente** al capturar un DPI duplicado (más
visible que la validación nativa), puedes guardar el archivo como **`.xlsm`** y pegar
esta macro en el módulo de la hoja (`Alt+F11` → doble clic en *Catálogo Maestro*):

```vba
' OPCIONAL — avisa con un mensaje si se captura un DPI repetido
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A2:A1000")) Is Nothing Then
        Dim c As Range
        For Each c In Intersect(Target, Me.Range("A2:A1000"))
            If c.Value <> "" Then
                If Application.WorksheetFunction.CountIf( _
                       Me.Range("A2:A1000"), c.Value) > 1 Then
                    MsgBox "El DPI '" & c.Value & "' ya existe.", _
                           vbExclamation, "DPI duplicado"
                End If
            End If
        Next c
    End If
End Sub
```

Y para **proteger las fórmulas** dejando editables solo los datos (opcional):

```vba
' OPCIONAL — bloquea columnas calculadas (H,I,J,K,S,T) y protege la hoja
Sub ProtegerFormulas()
    With Worksheets("Catálogo Maestro")
        .Cells.Locked = False
        .Range("H:K,S:T").Locked = True
        .Protect Password:="", UserInterfaceOnly:=True
    End With
End Sub
```

---

## 10. Próximo paso

Una vez valides este Sprint 1, seguimos con **SPRINT 2: Módulo de Entradas**, donde
*Stock Actual* y *Fecha Última Compra* se calcularán automáticamente desde las compras
registradas.
