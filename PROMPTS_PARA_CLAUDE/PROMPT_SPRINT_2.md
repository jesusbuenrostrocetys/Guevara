# 🔥 PROMPT PARA CLAUDE CODE - SPRINT 2: MÓDULO DE ENTRADAS

**INSTRUCCIÓN IMPORTANTE:** Copia TODO este contenido (desde "# INSTRUCCIÓN PRINCIPAL" hasta "---END PROMPT---") y pégalo en Claude Code.

---

# INSTRUCCIÓN PRINCIPAL

Soy Mario Guevara, propietario de SYSEA Radiadores. 

**SPRINT 1 (Catálogo Maestro) ha sido completado y validado exitosamente.**

Ahora necesito que construyas el **SPRINT 2: Módulo de Entradas** que será la interfaz para registrar compras de radiadores a proveedores.

**Mi repositorio de GitHub es:** `jesusbuenrostrocetys/Guevara`

**Ya existe:** El archivo SYSEA_Radiadores_SPRINT1.xlsx con el Catálogo Maestro funcional.

**Qué necesito:**
1. Agregar una nueva hoja llamada "Entradas" al archivo existente
2. Crear estructura de registro de compras
3. Vincular automáticamente con Catálogo Maestro
4. Actualizar Stock automáticamente
5. Calcular Costo Promedio Ponderado
6. Código VBA para automatización

---

## 📋 ESTRUCTURA DEL MÓDULO ENTRADAS

Debes crear una hoja llamada **"Entradas"** con estas columnas EXACTAS:

1. **ID Entrada** (Número) - Folio consecutivo auto-incrementable
2. **Fecha** (Fecha) - Día de la compra
3. **DPI** (Texto) - Número de parte (vinculado a Catálogo)
4. **Descripción** (Texto) - AUTOMÁTICA desde Catálogo (VLOOKUP)
5. **Cantidad** (Número) - Cantidad comprada
6. **Costo Unitario** (Moneda MXN) - Precio sin IVA
7. **Costo Total** (Moneda MXN) - FÓRMULA: Cantidad × Costo Unitario
8. **Proveedor** (Texto) - Nombre del proveedor
9. **Observaciones** (Texto) - Notas adicionales

---

## 🔗 FÓRMULA CRÍTICA: VLOOKUP PARA DESCRIPCIÓN

En la columna "Descripción", cuando se ingresa un DPI, debe traer automáticamente la descripción del Catálogo Maestro.

**Fórmula Excel (Columna D):**
```excel
=SI(C2="", "", VLOOKUP(C2,'Catálogo Maestro'!$A:$B,2,FALSE))
```

**Explicación:**
- Si el DPI está vacío (C2=""), dejar la celda vacía
- Si hay DPI, buscar en Catálogo Maestro columna A (DPI) y traer columna B (Descripción)
- FALSE = búsqueda exacta

---

## 💰 FÓRMULA: COSTO TOTAL

En la columna "Costo Total", calcular: Cantidad × Costo Unitario

**Fórmula Excel (Columna G):**
```excel
=E2*F2
```

---

## 📊 DATOS FICTICIOS: 10 ENTRADAS DE PRUEBA

Crea exactamente estas 10 entradas:

| Fecha | DPI | Cantidad | Costo Unitario | Proveedor |
|-------|-----|----------|----------------|----------|
| 2026-06-01 | RVVW029 | 2 | 1500 | Radiadores del Centro |
| 2026-06-02 | 0050 | 1 | 1200 | Radiadores del Centro |
| 2026-06-03 | 13015 | 3 | 800 | Radiadores Express |
| 2026-06-04 | FORD021 | 2 | 1350 | Radiadores del Centro |
| 2026-06-05 | HYUN008 | 1 | 950 | Radiadores Express |
| 2026-06-06 | RVVW029 | 1 | 1480 | Radiadores Express |
| 2026-06-07 | KIA045 | 2 | 1100 | Radiadores del Centro |
| 2026-06-08 | MAZU33 | 1 | 1600 | Radiadores Express |
| 2026-06-09 | TOYO88 | 2 | 1300 | Radiadores del Centro |
| 2026-06-10 | MITSU19 | 1 | 1050 | Radiadores Express |

**Nota:** Algunos DPI se repiten (ej: RVVW029 en filas 1 y 6). Esto es INTENCIONAL para demostrar cálculo de Costo Promedio Ponderado.

---

## 🆙 ACTUALIZACIÓN AUTOMÁTICA DE STOCK

### **Requerimiento:**
Cuando se procesa una entrada en esta hoja, el Stock Actual en la hoja "Catálogo Maestro" debe aumentar automáticamente.

**Ejemplo:**
- Catálogo: RVVW029 tiene Stock Actual = 1
- Se registra entrada: RVVW029, Cantidad = 2
- Se presiona botón "Procesar Entrada"
- Catálogo: RVVW029 ahora tiene Stock Actual = 3 (1 + 2)

### **Implementación:**
Debes crear un botón "Procesar Entrada" con código VBA que:

1. Lee la última fila de "Entradas"
2. Busca el DPI en "Catálogo Maestro"
3. Suma la Cantidad al Stock Actual
4. Actualiza la Fecha Última Compra
5. Calcula el Costo Promedio Ponderado (ver abajo)
6. Limpia la fila para la siguiente entrada
7. Muestra mensaje de confirmación

---

## 💎 CÁLCULO: COSTO PROMEDIO PONDERADO

### **Lógica Comercial:**
En radiadores, los proveedores cambian precios. El sistema debe calcular el costo promedio de cada producto.

**Fórmula matemática:**
```
Costo Promedio = (Stock anterior × Costo anterior + Nueva cantidad × Nuevo costo) / Stock total
```

**Ejemplo real:**
- Tenía: 1 unidad RVVW029 a $1,500 = $1,500 invertido
- Compré: 1 unidad RVVW029 a $1,480 = $1,480 invertido
- Total: 2 unidades por $2,980
- Costo Promedio: $2,980 ÷ 2 = $1,490

**En el Catálogo Maestro:**
La columna "Precio Compra" debe actualizarse a $1,490 después de procesar la entrada.

### **Implementación VBA:**

El botón "Procesar Entrada" debe:

1. Obtener el DPI de la última entrada
2. Obtener Stock anterior, Costo anterior del Catálogo
3. Obtener Nueva cantidad, Nuevo costo de Entradas
4. Calcular: Costo Promedio = (Stock_ant × Costo_ant + Cant_new × Costo_new) / (Stock_ant + Cant_new)
5. Actualizar "Precio Compra" en Catálogo Maestro
6. Recalcular todas las utilidades en Catálogo (porque el costo cambió)

---

## ✅ VALIDACIONES A IMPLEMENTAR

### **1. Validación de DPI**
- El DPI debe existir en Catálogo Maestro
- Si DPI no existe, mostrar error: "DPI no encontrado"
- No permitir procesar si DPI es inválido

### **2. Validación de Cantidad**
- Solo números positivos (1-10)
- No permitir 0
- No permitir > 10 (alerta de sobrecompra)

### **3. Validación de Costo**
- Solo números positivos
- No permitir 0
- No permitir valores muy bajos (< $100)

### **4. Validación de Fecha**
- No puede ser futura
- No puede ser muy antigua (> 1 año)

### **5. ID Entrada**
- Auto-incrementable
- Nunca puede repetirse
- Formato: 1, 2, 3, 4... (simple y secuencial)

---

## 🎨 FORMATO CONDICIONAL A APLICAR

### **Columna E (Cantidad):**
```
- Rojo:     Cantidad > 10 (alerta de sobrecompra)
- Amarillo: Cantidad entre 5-10 (compra alta)
- Verde:    Cantidad entre 1-4 (normal)
```

### **Columna G (Costo Total):**
```
- Verde:   < $5,000
- Amarillo: $5,000 - $10,000
- Naranja:  $10,000 - $20,000
- Rojo:     > $20,000
```

---

## 🆚 INTEGRACIÓN CON CATÁLOGO MAESTRO

**CRÍTICO:** Después de procesar cada entrada:

1. ✅ Stock Actual en Catálogo DEBE AUMENTAR
2. ✅ Precio Compra DEBE ACTUALIZARSE (costo promedio ponderado)
3. ✅ Fecha Última Compra DEBE CAMBIAR
4. ✅ Utilidad Cliente DEBE RECALCULARSE (porque costo cambió)
5. ✅ Utilidad Taller DEBE RECALCULARSE (porque costo cambió)
6. ✅ Precios Cliente/Taller PODRÍAN CAMBIAR (si utilidad cae bajo mínimo)

---

## 🛠️ CÓDIGO VBA: BOTÓN "PROCESAR ENTRADA"

Debes crear un botón en la hoja "Entradas" que ejecute esta lógica:

### **Pseudo-código (para que entiendas la lógica):**

```
Botón: "Procesar Entrada"

1. Leer última fila de "Entradas" con datos
2. Validar que DPI exista en Catálogo
3. Si error, mostrar: "DPI no encontrado"
4. Si válido:
   a. Obtener Stock anterior de Catálogo
   b. Obtener Costo anterior de Catálogo
   c. Obtener Cantidad nueva de Entradas
   d. Obtener Costo nuevo de Entradas
   e. Calcular Costo Promedio = (Stock_ant * Costo_ant + Cant_new * Costo_new) / (Stock_ant + Cant_new)
   f. Actualizar Stock = Stock_ant + Cant_new
   g. Actualizar Precio Compra = Costo Promedio (redondeado a múltiplo de 50)
   h. Actualizar Fecha Última Compra = Hoy
   i. Recalcular utilidades en Catálogo
   j. Mostrar: "Entrada procesada correctamente"
   k. Limpiar fila para nueva entrada
```

---

## 📋 CHECKLIST DE VALIDACIÓN

- [ ] Hoja "Entradas" creada
- [ ] 9 columnas estructuradas correctamente
- [ ] 10 entradas de ejemplo con datos ficticios
- [ ] ID Entrada auto-incrementable (1, 2, 3...)
- [ ] VLOOKUP de Descripción funcionando
- [ ] Fórmula de Costo Total correcta
- [ ] Validaciones de DPI, Cantidad, Costo funcionando
- [ ] Formato condicional aplicado en Cantidad y Costo Total
- [ ] Botón "Procesar Entrada" visible
- [ ] VBA de actualización de Stock funcionando
- [ ] VBA de Costo Promedio Ponderado funcionando
- [ ] Stock en Catálogo aumenta correctamente
- [ ] Precio Compra en Catálogo se actualiza
- [ ] Fecha Última Compra en Catálogo se actualiza
- [ ] Utilidades en Catálogo se recalculan
- [ ] Sin errores en fórmulas o código

---

## 🆚 ARCHIVO DE ENTRADA

**Debes partir del archivo existente:**
```
SYSEA_Radiadores_SPRINT1.xlsx
```

Este archivo ya tiene:
- Hoja "Catálogo Maestro" con 20 productos
- Todas las fórmulas del SPRINT 1 funcionando

**Tú debes:**
- Agregar la hoja "Entradas"
- NO modificar Catálogo Maestro (solo actualizarlo automáticamente)
- Integrar completamente ambas hojas

---

## 🚀 ENTREGABLE ESPERADO

1. **Archivo Excel actualizado:** `SYSEA_Radiadores_SPRINT2.xlsx`
2. **Nueva hoja:** "Entradas" completamente funcional
3. **10 entradas de prueba:** Con datos ficticios
4. **Todas las fórmulas:** VLOOKUP, Costo Total, etc.
5. **Validaciones:** Previenen errores
6. **Formato condicional:** Colores en Cantidad y Costo Total
7. **VBA funcional:** Botón "Procesar Entrada"
8. **Integración:** Stock y precios en Catálogo se actualizan
9. **Explicación:** De cada fórmula y macro
10. **Instrucciones:** De cómo usar el módulo

---

## 📝 PRUEBAS A REALIZAR

**Después de crear todo:**

1. Abre el archivo en Excel
2. Ve a la hoja "Entradas"
3. Verifica que las 10 entradas estén correctas
4. Haz clic en "Procesar Entrada" para la primera entrada
5. Ve al Catálogo Maestro
6. Verifica que el Stock aumentó
7. Verifica que el Precio Compra cambió (costo promedio)
8. Verifica que la Fecha Última Compra se actualizó
9. Verifica que las utilidades se recalcularon
10. Repite para las demás entradas

---

## ⚠️ RESTRICCIONES OBLIGATORIAS

- ❌ No usar datos reales
- ❌ No modificar estructura de Catálogo Maestro
- ❌ No eliminar o cambiar los 20 productos existentes
- ❌ No cambiar política de precios
- ❌ No crear fórmulas complejas innecesarias
- ✅ Usar datos ficticios para todas las entradas
- ✅ Mantener compatibilidad con Catálogo Maestro
- ✅ Código VBA limpio y bien documentado

---

## 📞 SIGUIENTE PASO

Una vez valides este SPRINT 2, procederemos a **SPRINT 3: Módulo POS (Punto de Venta)**.

---

Gracias. Espero que el módulo de Entradas sea perfectamente funcional y esté 100% integrado con el Catálogo Maestro.

---END PROMPT---
