# 🎯 PROMPT PARA CLAUDE CODE - SPRINT 1: CATÁLOGO MAESTRO

**INSTRUCCIÓN IMPORTANTE:** Copia TODO este contenido (desde "# INSTRUCCIÓN PRINCIPAL" hasta "---END PROMPT---") y pégalo en Claude Code.

---

# INSTRUCCIÓN PRINCIPAL

Soy Mario Guevara, propietario de SYSEA Radiadores. Estoy construyendo un sistema integral en Excel 365 para administrar mi negocio de venta de radiadores automotrices.

Este es el **SPRINT 1**, enfocado en crear el Catálogo Maestro con estructura completa, fórmulas automáticas de precios, y 20 productos de ejemplo para validación.

**Mi repositorio de GitHub es:** `jesusbuenrostrocetys/Guevara`

Por favor:
1. Entiende que es Sprint 1 de un proyecto mayor (9 Sprints totales)
2. Usa SOLO datos ficticios (sin datos reales aún)
3. Crea un Catálogo Maestro perfectamente funcional
4. Explica cada fórmula en detalle
5. Proporciona instrucciones de instalación
6. Incluye código VBA si es necesario

---

## 🏢 CONTEXTO DEL NEGOCIO

**Empresa:** SYSEA Radiadores  
**Giro:** Venta de radiadores automotrices  
**Tipo de cliente:** Clientes finales y talleres  
**Proveedores:** Reposición 2-4 días  
**Inventario típico:** 2-3 piezas por DPI, máximo 4  
**Objetivo:** Maximizar utilidad sin perder competitividad

---

## 📊 ESTRUCTURA DEL CATÁLOGO MAESTRO

Debes crear una hoja en Excel llamada **"Catálogo Maestro"** con estas columnas EXACTAS:

### COLUMNAS OBLIGATORIAS

1. **DPI** (Texto) - Número de parte único. Ej: RVVW029, 0050, 13015
2. **Descripción** (Texto) - Ej: "VW Jetta 2010-2018 Radiador"
3. **Marca** (Texto) - Marca del vehículo. Ej: "Volkswagen"
4. **Modelo** (Texto) - Modelo del vehículo. Ej: "Jetta"
5. **Año** (Número) - Año de fabricación. Ej: 2010
6. **Precio Compra** (Moneda MXN) - Costo sin IVA. Ej: $1,500
7. **Precio AutoZone** (Moneda MXN) - Referencia de competencia. Ej: $3,900
8. **Precio Cliente Final** (Moneda MXN) - **FÓRMULA AUTOMÁTICA**
9. **Precio Taller** (Moneda MXN) - **FÓRMULA AUTOMÁTICA**
10. **Utilidad Cliente** (Moneda MXN) - **FÓRMULA AUTOMÁTICA**
11. **Utilidad Taller** (Moneda MXN) - **FÓRMULA AUTOMÁTICA**
12. **Stock Actual** (Número) - Cantidad disponible. Vinculado a "Entradas"
13. **Stock Mínimo** (Número) - Límite mínimo (parámetro). Típicamente: 1-2
14. **Stock Ideal** (Número) - Cantidad objetivo. Típicamente: 3
15. **Proveedor Principal** (Texto) - Ej: "Radiadores del Centro"
16. **Fecha Última Compra** (Fecha) - **AUTOMÁTICA** desde "Entradas"
17. **Fecha Última Venta** (Fecha) - **AUTOMÁTICA** desde "Ventas"
18. **Estado** (Dropdown: Activo/Inactivo) - Estado del producto
19. **Rotación** (Texto) - **FÓRMULA AUTOMÁTICA** (Estrella/Media/Lenta/Muerta)
20. **Alerta AutoZone** (Texto) - **FÓRMULA AUTOMÁTICA** (OK/Pendiente/Revisar)
21. **Observaciones** (Texto) - Notas adicionales

---

## 🧮 FÓRMULAS CRÍTICAS A IMPLEMENTAR

### FÓRMULA 1: PRECIO CLIENTE FINAL (Columna H)

**Lógica Comercial:**
1. Si existe precio AutoZone:
   - Buscar quedarse $250 MXN por debajo
   - Garantizar utilidad mínima de $1,200 MXN
   - Si margen es muy amplio, puedo bajar hasta $500 MXN
2. Si NO existe precio AutoZone:
   - Partir de: Precio Compra + $1,500 MXN (utilidad ideal)
3. Redondear SIEMPRE a múltiplo de 50 hacia ARRIBA
   - Ej: $2,320 → $2,350
   - Ej: $2,380 → $2,400

**Fórmula Excel (Versión Simplificada - es decir, versión funcional para el Sprint 1):**
```excel
=SI(G2>"",
  REDONDEAR.MAS((SI(G2-C2-250<1200, C2+1200, G2-250))/50)*50,
  REDONDEAR.MAS((C2+1500)/50)*50
)
```

**Explicación:**
- Si G2 (AutoZone) existe: Restar $250, pero garantizar utilidad mínima de $1,200
- Si no existe: Sumar $1,500 a precio compra
- Redondear siempre hacia arriba a múltiplo de 50

### FÓRMULA 2: PRECIO TALLER (Columna I)

**Lógica Comercial:**
1. Base: Precio Cliente Final (Columna H)
2. Objetivo: Otorgar descuento atractivo (~15-20%)
3. Restricción: NUNCA menos de $800 MXN de utilidad
4. Redondear a múltiplo de 50 hacia arriba

**Fórmula Excel:**
```excel
=REDONDEAR.MAS(MAX(C2+800, H2*0.82)/50)*50
```

**Explicación:**
- MAX asegura que haya mínimo $800 de utilidad
- H2*0.82 da aproximadamente 18% de descuento sobre Cliente Final
- Redondea hacia arriba a múltiplo de 50

### FÓRMULA 3: UTILIDAD CLIENTE (Columna J)

**Fórmula Excel:**
```excel
=H2-C2
```

**Explicación:** Precio Cliente Final menos Precio Compra

### FÓRMULA 4: UTILIDAD TALLER (Columna K)

**Fórmula Excel:**
```excel
=I2-C2
```

**Explicación:** Precio Taller menos Precio Compra

### FÓRMULA 5: ROTACIÓN (Columna S)

**Lógica de Clasificación:**
- **Estrella:** Última venta ≤ 90 días
- **Media:** Última venta 91-180 días
- **Lenta:** Última venta 181-365 días
- **Muerta:** Última venta > 365 días
- **Sin venta:** Nunca ha sido vendido

**Fórmula Excel:**
```excel
=SI(Q2="",
  "Sin venta",
  SI(HOY()-Q2<=90,
    "Estrella",
    SI(HOY()-Q2<=180,
      "Media",
      SI(HOY()-Q2<=365,
        "Lenta",
        "Muerta"
      )
    )
  )
)
```

**Explicación:**
Q2 es "Fecha Última Venta". La fórmula calcula días desde la última venta y clasifica automáticamente.

### FÓRMULA 6: ALERTA AUTOZONE (Columna T)

**Lógica:**
- Si G2 (Precio AutoZone) está vacío: "Pendiente"
- Si G2 tiene valor > 0: "OK"
- Si G2 = 0 o error: "Revisar"

**Fórmula Excel:**
```excel
=SI(G2="",
  "Pendiente",
  SI(G2>0,
    "OK",
    "Revisar"
  )
)
```

---

## 📋 DATOS FICTICIOS: 20 PRODUCTOS PARA VALIDACIÓN

Por favor, crea exactamente estos 20 productos con datos realistas pero ficticios:

| DPI | Descripción | Marca | Modelo | Año | Precio Compra | Precio AutoZone |
|-----|-------------|-------|--------|-----|---------------|----------------|
| RVVW029 | VW Jetta Radiador | Volkswagen | Jetta | 2010 | 1500 | 3900 |
| 0050 | Nissan Sentra Radiador | Nissan | Sentra | 2015 | 1200 | 3200 |
| 13015 | Chevy Aveo Radiador | Chevrolet | Aveo | 2012 | 800 | 2500 |
| FORD021 | Ford Focus Radiador | Ford | Focus | 2013 | 1350 | 3600 |
| HYUN008 | Hyundai Elantra Radiador | Hyundai | Elantra | 2011 | 950 | 2800 |
| KIA045 | Kia Cerato Radiador | Kia | Cerato | 2014 | 1100 | 3100 |
| MAZU33 | Mazda 3 Radiador | Mazda | Mazda 3 | 2016 | 1600 | 4200 |
| MERC77 | Mercedes C200 Radiador | Mercedes-Benz | C200 | 2012 | 3500 | 8900 |
| SUBR12 | Subaru Impreza Radiador | Subaru | Impreza | 2010 | 1450 | 3700 |
| TOYO88 | Toyota Corolla Radiador | Toyota | Corolla | 2015 | 1300 | 3400 |
| MITSU19 | Mitsubishi Lancer Radiador | Mitsubishi | Lancer | 2011 | 1050 | 2900 |
| VOLKS30 | VW Passat Radiador | Volkswagen | Passat | 2009 | 2000 | 5200 |
| CHRY55 | Chrysler Neon Radiador | Chrysler | Neon | 2008 | 700 | 2200 |
| HONDA64 | Honda Civic Radiador | Honda | Civic | 2012 | 1250 | 3300 |
| SUZUK16 | Suzuki Swift Radiador | Suzuki | Swift | 2014 | 900 | 2600 |
| BUICK22 | Buick Century Radiador | Buick | Century | 2005 | 600 | 1900 |
| JEEP99 | Jeep Wrangler Radiador | Jeep | Wrangler | 2013 | 1800 | 4800 |
| LEXUS50 | Lexus RX350 Radiador | Lexus | RX350 | 2010 | 4000 | 9500 |
| PEUGEOT11 | Peugeot 206 Radiador | Peugeot | 206 | 2006 | 750 | 2400 |
| RENAULT88 | Renault Duster Radiador | Renault | Duster | 2015 | 1200 | 3300 |

**Para cada producto:**
- Stock Actual: Valor aleatorio entre 0 y 4
- Stock Mínimo: 1
- Stock Ideal: 3
- Proveedor Principal: Radiadores del Centro
- Fecha Última Compra: Entre 1-30 días atrás
- Fecha Última Venta: Valores variados (algunos sin venta)
- Estado: Activo
- Observaciones: Dejar en blanco o notas relevantes

---

## ✅ VALIDACIONES A IMPLEMENTAR

1. **Columna DPI:** No permitir duplicados
2. **Columna Precio Compra:** Solo números positivos
3. **Columna Stock Actual:** No permitir negativos
4. **Columna Stock Mínimo:** No permitir negativos
5. **Columna Stock Ideal:** No permitir negativos
6. **Redondeo de Precios:** Verificar que TODOS terminen en 00 o 50

---

## 🎨 FORMATO CONDICIONAL A APLICAR

### En Columna L (Stock Actual):
```
- Rojo:        Stock = 0
- Naranja:     0 < Stock < Stock Mínimo
- Amarillo:    Stock Mínimo ≤ Stock < Stock Ideal
- Verde:       Stock ≥ Stock Ideal
```

### En Columna S (Rotación):
```
- Verde:    "Estrella"
- Amarillo: "Media"
- Naranja:  "Lenta"
- Rojo:     "Muerta"
- Gris:     "Sin venta"
```

### En Columna J (Utilidad Cliente):
```
- Rojo:     < $1,200
- Amarillo: $1,200 - $1,499
- Verde:    ≥ $1,500
```

---

## 📝 INSTRUCCIONES DE INSTALACIÓN

Debes incluir en tu respuesta:

1. **Paso 1:** Crear la hoja "Catálogo Maestro"
2. **Paso 2:** Crear encabezados exactos (nombres de columnas)
3. **Paso 3:** Cargar los 20 productos ficticios
4. **Paso 4:** Insertar todas las fórmulas
5. **Paso 5:** Aplicar validaciones
6. **Paso 6:** Aplicar formatos condicionales
7. **Paso 7:** Proteger la hoja (opcional para esta etapa)
8. **Paso 8:** Crear archivo Excel descargable

---

## 🧪 PRUEBAS A REALIZAR

Una vez creado, debes verificar:

1. ¿El precio Cliente Final se calcula correctamente?
2. ¿Todos los precios terminan en 00 o 50?
3. ¿Las utilidades son ≥ $800 para talleres?
4. ¿Las utilidades son ≥ $1,200 para clientes?
5. ¿La rotación se clasifica automáticamente?
6. ¿El formato condicional se aplica correctamente?
7. ¿Las fórmulas no generan errores (#DIV/0!, #REF!, etc.)?
8. ¿Puedo editar datos sin romper fórmulas?

---

## 📊 ENTREGABLES ESPERADOS

1. **Archivo Excel:** SYSEA_Radiadores_SPRINT1.xlsx
2. **Hoja:** "Catálogo Maestro" completamente funcional
3. **20 productos:** Con datos realistas y ficticios
4. **Todas las fórmulas:** Correctamente implementadas
5. **Validaciones:** Activas y funcionando
6. **Formato condicional:** Aplicado en columnas clave
7. **Explicación:** De cada fórmula en detalle
8. **Instrucciones:** De cómo usar y mantener
9. **Código VBA:** Si es necesario (mínimo para Sprint 1)

---

## 🚫 RESTRICCIONES OBLIGATORIAS

- ❌ No incluir datos reales
- ❌ No asumir reglas no definidas
- ❌ No modificar la política de precios
- ❌ No cambiar estructura de columnas
- ❌ No incluir módulos de Entradas o Ventas (solo Catálogo)
- ❌ No usar datos de proveedores reales

---

## ✋ VALIDACIÓN FINAL

Antes de entregar, confirma:

- [ ] Archivo Excel cargable sin errores
- [ ] 20 productos con datos ficticios realistas
- [ ] Todas las fórmulas funcionan correctamente
- [ ] Precios terminan en 00 o 50
- [ ] Utilidades respetan márgenes mínimos
- [ ] Formatos condicionales visibles
- [ ] Explicación clara de cada fórmula
- [ ] Instrucciones de instalación paso a paso
- [ ] Listo para descargar y probar en Excel

---

## 📞 PRÓXIMO PASO

Una vez valides este Sprint 1, procederemos a **SPRINT 2: Módulo de Entradas**.

---

Gracias por tu ayuda. Espero que el resultado sea un Catálogo Maestro perfectamente funcional y listo para integrar con el resto de módulos.

---END PROMPT---
