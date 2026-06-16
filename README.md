# 📊 Guevara - Sistema de Ventas

Un sistema completo de gestión de ventas y control de inventario diseñado para facilitar la administración de transacciones comerciales.

## 🎯 Características

- ✅ Gestión de productos y catálogo
- ✅ Registro de ventas con detalles de transacciones
- ✅ Control de inventario en tiempo real
- ✅ Generación de reportes de ventas
- ✅ Cálculo automático de totales, impuestos y descuentos
- ✅ Historial de transacciones
- ✅ Gestión de clientes
- ✅ Exportación a Excel

## 🚀 Inicio Rápido

### Requisitos Previos

- Python 3.8+ (si es backend)
- Node.js 14+ (si es frontend)
- MySQL/PostgreSQL (para base de datos, opcional)
- Excel (para importar/exportar datos)

### Instalación

```bash
# Clona el repositorio
git clone https://github.com/jesusbuenrostrocetys/Guevara.git
cd Guevara

# Instala dependencias
pip install -r requirements.txt
# o
npm install
```

### Configuración

1. Configura tu base de datos en el archivo `.env`
2. Ejecuta las migraciones necesarias
3. Inicia la aplicación

```bash
python app.py
# o
npm start
```

## 📋 Estructura del Proyecto

```
Guevara/
├── src/
│   ├── models/          # Modelos de datos
│   ├── controllers/     # Lógica de negocio
│   └── utils/          # Funciones auxiliares
├── tests/              # Pruebas unitarias
├── data/               # Archivos Excel y datos
└── requirements.txt    # Dependencias
```

## 💼 Funcionalidades Principales

### 1. Gestión de Ventas
- Crear nueva venta
- Ver historial de ventas
- Editar/eliminar ventas
- Generar recibos

### 2. Control de Inventario
- Actualizar stock de productos
- Alertas de inventario bajo
- Reportes de movimiento de inventario

### 3. Reportes
- Reporte de ventas por período
- Análisis por producto
- Rendimiento de ventas

### 4. Integración Excel
- Importar catálogo de productos desde Excel
- Exportar ventas a Excel
- Plantillas descargables

## 🛠 Tecnologías Utilizadas

- **Backend**: Python/Flask o Node.js/Express
- **Frontend**: HTML/CSS/JavaScript o React/Vue
- **Base de Datos**: MySQL/PostgreSQL/SQLite
- **Excel**: Openpyxl o Pandas (Python)
- **IA**: Claude Code (para desarrollo asistido)

## 📝 Uso Básico

```
1. Registra tus productos en el sistema
2. Crea nuevas ventas
3. El sistema actualizará automáticamente el inventario
4. Genera reportes cuando lo necesites
5. Exporta datos a Excel si es necesario
```

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Fork el repositorio
2. Crea una rama para tu feature (`git checkout -b feature/NuevaFuncionalidad`)
3. Commit tus cambios (`git commit -m 'Add: Nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/NuevaFuncionalidad`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está bajo la licencia MIT. Ver el archivo [LICENSE](LICENSE) para más detalles.

## 👤 Autor

**jesusbuenrostrocetys**

## 📞 Contacto & Soporte

Si tienes preguntas o sugerencias, abre un [Issue](https://github.com/jesusbuenrostrocetys/Guevara/issues) en el repositorio.

---

**Nota**: Este sistema está en desarrollo. Con la ayuda de Claude Code, se irán agregando más funcionalidades y mejoras de forma continua.

*Desarrollado con ❤️ y Claude Code*
