# SPRINT 1: CATÁLOGO MAESTRO CON DATOS FICTICIOS

## 🎯 OBJETIVO

Construir la base de datos maestro de productos con estructura completa, fórmulas automáticas y 20 productos de ejemplo para validar la lógica.

---

## 📋 COLUMNAS DEL CATÁLOGO

| Columna | Tipo | Descripción | Automático |
|---------|------|-------------|------------|
| DPI | Texto | Número de parte único | No |
| Descripción | Texto | Vehículo/marca/modelo | No |
| Marca | Texto | Marca del vehículo | No |
| Modelo | Texto | Modelo del vehículo | No |
| Año | Número | Año de fabricación | No |
| Precio Compra | Moneda | Costo unitario sin IVA | No |
| Precio AutoZone | Moneda | Referencia de competencia | Manual |
| Precio Cliente Final | Moneda | Sugerido para cliente | **Fórmula** |
| Precio Taller | Moneda | Sugerido para taller | **Fórmula** |
| Utilidad Cliente | Moneda | Ganancia por venta a cliente | **Fórmula** |
| Utilidad Taller | Moneda | Ganancia por venta a taller | **Fórmula** |
| Stock Actual | Número | Cantidad disponible | **Automático** |
| Stock Mínimo | Número | Límite mínimo (alerta) | Parámetro |
| Stock Ideal | Número | Cantidad objetivo | Parámetro |
| Proveedor Principal | Texto | Proveedor habitual | No |
| Fecha Última Compra | Fecha | Última entrada registrada | **Automático** |
| Fecha Última Venta | Fecha | Última salida registrada | **Automático** |
| Estado | Dropdown | Activo/Inactivo | Manual |
| Rotación | Texto | Estrella/Media/Lenta/Muerta | **Fórmula** |
| Alerta AutoZone | Texto | Sí/No/Pendiente | **Fórmula** |
| Observaciones | Texto | Notas adicionales | No |

---

## 🧮 FÓRMULAS CRÍTICAS

### 1. PRECIO CLIENTE FINAL

**Lógica:**
1. Si AutoZone existe:
   - Descuento deseado = $250 (promedio)
   - Precio inicial = AutoZone - $250
   - Si Utilidad < $1,200: Aumentar hasta garantizar $1,200
   - Si Utilidad >= $1,500: Mantener
2. Si AutoZone NO existe:
   - Base = Precio Compra + $1,500 (utilidad mínima ideal)
3. Redondear a múltiplo de 50 hacia arriba

**Fórmula Excel (versión simplificada para comenzar):**
```excel
=SI(D2>0, 
  REDONDEAR.MAS((D2-250)/50)*50,
  REDONDEAR.MAS((C2+1500)/50)*50
)
```

### 2. PRECIO TALLER

**Lógica:**
1. Base = Precio Cliente Final
2. Descuento atractivo (~15-20%)
3. Garantizar utilidad mínima $800
4. Redondear a múltiplo de 50 hacia arriba

**Fórmula Excel:**
```excel
=REDONDEAR.MAS(MAX(C2+800, H2*0.80)/50)*50
```

### 3. UTILIDAD CLIENTE

**Fórmula:**
```excel
=H2-C2
```

### 4. UTILIDAD TALLER

**Fórmula:**
```excel
=I2-C2
```

### 5. ROTACIÓN (basada en fecha última venta)

**Lógica:**
- Estrella: ≤ 90 días
- Media: 91-180 días
- Lenta: 181-365 días
- Muerta: > 365 días

**Fórmula Excel:**
```excel
=SI(Q2="", "Sin venta",
  SI(HOY()-Q2<=90, "Estrella",
    SI(HOY()-Q2<=180, "Media",
      SI(HOY()-Q2<=365, "Lenta", "Muerta")
    )
  )
)
```

### 6. ALERTA AUTOZONE

**Fórmula:**
```excel
=SI(D2="", "Pendiente", SI(D2>0, "OK", "Revisar"))
```

---

## 📊 DATOS FICTICIOS DE EJEMPLO

*(Tabla con 20 productos para validación)*

Ejemplo de 3 productos:

| DPI | Descripción | Marca | Modelo | Año | P.Compra | P.AutoZone |
|-----|-------------|-------|--------|-----|----------|------------|
| RVVW029 | VW Jetta Radiador | VW | Jetta | 2010 | 1500 | 3900 |
| 0050 | Nissan Sentra Motor | Nissan | Sentra | 2015 | 1200 | 3200 |
| 13015 | Chevy Aveo Plastico | Chevrolet | Aveo | 2012 | 800 | 2500 |

---

## ✅ VALIDACIONES A IMPLEMENTAR

1. **DPI único** - No permitir duplicados
2. **Precios positivos** - Mayor a cero
3. **Stock no negativo** - No permitir valores < 0
4. **Utilidad mínima** - Alertar si está por debajo de límite
5. **Redondeo correcto** - Solo 00 y 50

---

## 🎨 FORMATO CONDICIONAL

1. **Stock Actual:**
   - Verde: Stock ≥ Stock Ideal
   - Amarillo: Stock Mínimo ≤ Stock < Stock Ideal
   - Rojo: Stock = 0
   - Naranja: Stock < Stock Mínimo

2. **Rotación:**
   - Verde: Estrella
   - Amarillo: Media
   - Naranja: Lenta
   - Rojo: Muerta

3. **Utilidad Cliente:**
   - Verde: ≥ $1,500
   - Amarillo: $1,200-$1,499
   - Rojo: < $1,200

---

## 📋 CHECKLIST DE VALIDACIÓN

- [ ] Tabla estructurada con todas las columnas
- [ ] 20 productos de ejemplo con datos realistas
- [ ] Fórmulas de precios funcionando correctamente
- [ ] Redondeo a múltiplos de 50 verificado
- [ ] Cálculo de utilidades correcto
- [ ] Clasificación de rotación automática
- [ ] Formato condicional aplicado
- [ ] Validaciones de datos activadas
- [ ] Sin errores en fórmulas (#DIV/0!, #N/A, etc.)
- [ ] Hoja protegida (opcional en esta etapa)

---

## 🚀 ENTREGABLE

- Archivo Excel: `SYSEA_Radiadores_SPRINT1.xlsx`
- Hoja: "Catálogo Maestro"
- Con 20 productos de prueba completos
- Todas las fórmulas funcionando
- Listo para importar datos reales en etapa posterior

---

## ✋ VALIDACIÓN REQUERIDA

Antes de avanzar a SPRINT 2:

1. ¿Precios se calculan correctamente?
2. ¿Redondeos son siempre hacia arriba a múltiplos de 50?
3. ¿Utilidades están dentro de parámetros?
4. ¿Rotación se clasifica automáticamente?
5. ¿Formatos condicionales funcionan?

**Aprobación requerida para continuar** ✓
