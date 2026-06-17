# SPRINT 2: MÓDULO DE ENTRADAS

## 🎯 OBJETIVO

Construir el módulo de **Entradas (Compras a Proveedores)** que:
- Registre compras de radiadores a proveedores
- Actualice automáticamente el stock del Catálogo Maestro
- Calcule costo promedio ponderado
- Registre fecha de última compra
- Valide datos y prevea errores

---

## 📋 ESTRUCTURA DEL MÓDULO DE ENTRADAS

Debes crear una hoja en Excel llamada **"Entradas"** con estas columnas:

| Columna | Tipo | Descripción | Automático |
|---------|------|-------------|------------|
| ID Entrada | Número | Folio consecutivo único | **Fórmula** |
| Fecha | Fecha | Fecha de la compra | Manual (default: hoy) |
| DPI | Texto | Número de parte (vinculado a Catálogo) | Manual |
| Descripción | Texto | Se obtiene del Catálogo | **Automática** |
| Cantidad | Número | Cantidad comprada | Manual |
| Costo Unitario | Moneda | Precio sin IVA | Manual |
| Costo Total | Moneda | Cantidad × Costo Unitario | **Fórmula** |
| Proveedor | Texto | Nombre del proveedor | Manual |
| Observaciones | Texto | Notas adicionales | Manual |

---

## 🔄 VÍNCULOS CON CATÁLOGO MAESTRO

### **Entrada automática de Descripción**
Cuando se registra un DPI, la descripción debe traerse automáticamente del Catálogo Maestro.

**Fórmula Excel:**
```excel
=VLOOKUP(D2,'Catálogo Maestro'!$A:$B,2,FALSE)
```

---

## ✅ ACTUALIZACIÓN AUTOMÁTICA DEL STOCK

### **Lógica:**
Cada vez que se registra una entrada, el Stock Actual del Catálogo debe aumentar.

**Proceso:**
1. Se registra DPI en Entradas
2. Se ingresa Cantidad
3. Al guardar, el Stock en Catálogo = Stock anterior + Cantidad nueva

**Esto requiere VBA** para que sea automático en tiempo real.

---

## 💰 COSTO PROMEDIO PONDERADO

### **Lógica Comercial:**
Si compras radiadores a precios diferentes, el sistema debe calcular el costo promedio.

**Ejemplo:**
- Tenías: 2 unidades a $1,500 = $3,000 invertido
- Compras: 3 unidades a $1,400 = $4,200 invertido
- Total: 5 unidades por $7,200
- Costo promedio: $7,200 ÷ 5 = $1,440

**En el Catálogo:**
El "Precio Compra" debe actualizarse al costo promedio ponderado después de cada entrada.

---

## 🔗 ACTUALIZACIÓN DE FECHA ÚLTIMA COMPRA

Cuando se registra una entrada para un DPI, la "Fecha Última Compra" en el Catálogo debe actualizarse automáticamente.

**Esto también requiere VBA.**

---

## 📊 VALIDACIONES A IMPLEMENTAR

1. **DPI válido:**
   - Debe existir en Catálogo Maestro
   - No permitir DPI inválidos

2. **Cantidad:**
   - Solo números positivos
   - Máximo 10 (no comprar más de lo necesario)

3. **Costo Unitario:**
   - Solo números positivos
   - No permitir cero

4. **Fecha:**
   - No puede ser futura
   - Debe ser válida

5. **ID Entrada:**
   - Auto-incrementable
   - Nunca repetir

---

## 📝 DATOS FICTICIOS PARA PRUEBA

Crea 10 entradas de ejemplo:

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

---

## 🛠️ CÓDIGO VBA REQUERIDO

### **Macro 1: Actualizar Stock al registrar Entrada**

```vba
Sub ActualizarStock()
' Esta macro se ejecuta cuando se completa una entrada
' Busca el DPI en Catálogo Maestro
' Aumenta el Stock Actual con la cantidad de la entrada
' Actualiza Fecha Última Compra
End Sub
```

### **Macro 2: Calcular Costo Promedio Ponderado**

```vba
Sub CalcularCostoPromedio()
' Calcula: (Stock anterior × Costo anterior + Nueva cantidad × Nuevo costo) / Stock total
' Actualiza el "Precio Compra" en Catálogo Maestro
End Sub
```

### **Botón de acción:**
Un botón **"Procesar Entrada"** que ejecute ambas macros.

---

## 🎨 FORMATO CONDICIONAL

### **En columna Cantidad:**
```
- Rojo: Cantidad > 10 (alerta de sobrecompra)
- Amarillo: Cantidad entre 5-10 (normal alto)
- Verde: Cantidad entre 1-4 (normal)
```

### **En columna Costo Total:**
```
- Verde: < $5,000
- Amarillo: $5,000 - $10,000
- Naranja: $10,000 - $20,000
- Rojo: > $20,000
```

---

## 📋 CHECKLIST DE VALIDACIÓN

- [ ] Hoja "Entradas" creada
- [ ] Columnas estructuradas correctamente
- [ ] 10 entradas de ejemplo con datos realistas
- [ ] Fórmulas de VLOOKUP funcionando (Descripción automática)
- [ ] Fórmula de Costo Total correcta
- [ ] Validaciones de datos implementadas
- [ ] Formato condicional aplicado
- [ ] VBA para actualizar Stock funcionando
- [ ] VBA para Costo Promedio funcionando
- [ ] Botón "Procesar Entrada" visible y funcional
- [ ] Catálogo Maestro actualiza correctamente
- [ ] Sin errores en fórmulas

---

## 🆚 INTEGRACIÓN CON SPRINT 1

**Después de procesar entradas:**
1. El Stock Actual en Catálogo Maestro debe aumentar
2. El Precio Compra debe reflejar costo promedio ponderado
3. La Fecha Última Compra debe actualizarse
4. Las utilidades (Cliente y Taller) deben recalcularse automáticamente

---

## 🚀 ENTREGABLE

- Archivo Excel actualizado: `SYSEA_Radiadores_SPRINT2.xlsx`
- Hoja: "Entradas" completamente funcional
- Con 10 entradas de prueba
- Todas las fórmulas y VBA implementados
- Integración con Catálogo Maestro verificada
- Listo para procesar compras reales

---

## ✅ VALIDACIÓN REQUERIDA

Antes de avanzar a SPRINT 3:

1. ¿Las descripciones se traen automáticamente del Catálogo?
2. ¿El Costo Total se calcula correctamente?
3. ¿Al procesar una entrada, el Stock aumenta en Catálogo?
4. ¿La Fecha Última Compra se actualiza?
5. ¿El Costo Promedio Ponderado funciona?
6. ¿Las utilidades en Catálogo se recalculan?
7. ¿Hay validaciones que previenen errores?
8. ¿El formato condicional funciona?

**Aprobación requerida para continuar** ✓
