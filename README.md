# Excel SQL Analyzer

Una aplicación de escritorio para analizar archivos Excel mediante consultas SQL, con interfaz gráfica.

## Características principales

- **Carga de archivos Excel**: Soporte para múltiples hojas de cálculo
- **Editor SQL con resaltado de sintaxis**: Reconocimiento de palabras clave SQL
- **Ejecución de consultas**: Visualización de resultados en formato tabla
- **Gestión de historial**:
  - Guardado automático de consultas
  - Posibilidad de renombrar y eliminar consultas
  - Recuperación rápida de consultas anteriores
- **Temas visuales**: Alternancia entre modo claro y oscuro
- **Base de datos en memoria**: Uso de SQLite para procesamiento rápido

## Requisitos

- Python 3.7 o superior
- Dependencias:
  ```
  pandas
  PyQt5
  openpyxl (para soporte de Excel)
  ```

## Uso

1. Haz clic en "Cargar Excel" para seleccionar un archivo
2. Selecciona la hoja de cálculo que deseas analizar
3. Escribe tu consulta SQL en el editor
4. Ejecuta la consulta con el botón "Ejecutar SQL"
5. Los resultados se mostrarán en la tabla inferior
6. Puedes guardar consultas frecuentes con un nombre descriptivo

## Historial de consultas

Las consultas se guardan automáticamente en un archivo JSON (`query_history.json`) en el mismo directorio que la aplicación. Puedes:

- Hacer doble clic en una consulta del historial para cargarla
- Renombrar consultas con clic derecho
- Eliminar consultas no deseadas

**🔧 Próximas Funcionalidades**

📌 Interfaz de Usuario Mejorada

Pestañas múltiples: Trabaja con varias consultas al mismo tiempo.

Autocompletado SQL: Sugerencias de tablas, columnas y palabras clave.

Formateador SQL: Indentación automática para consultas legibles.

Gráficos integrados: Visualización de datos con Matplotlib/Plotly.

Perfiles de usuario: Guardar configuraciones y consultas favoritas.


📊 Soporte para Más Formatos

Importar/Exportar a CSV, JSON y otros formatos.

Conexión a bases de datos externas (MySQL, PostgreSQL, SQL Server).

⚡ Análisis de Datos Avanzado

Estadísticas rápidas (media, moda, percentiles).

Limpieza de datos (eliminar nulos, normalizar texto).

Transformaciones (pivotar, agrupar, filtrar con interfaz gráfica).

📂 Gestión de Consultas y Datos

carpeta de quarys de distintos trabajos

Exportar resultados a Excel, CSV o copiar al portapapeles.

Generador visual de consultas (drag-and-drop para JOINs y WHERE).

Variables en consultas (ej: SELECT * FROM ventas WHERE fecha = '{{fecha}}').

🔐 Seguridad y Rendimiento

Advertencia antes de ejecutar consultas peligrosas (DROP, DELETE sin WHERE).

Caché de consultas frecuentes para mayor velocidad.

🚀 Ideas a Largo Plazo

🤖 Automatización


## **ROADMAP**

🚀 Phase 1: Core UX & Productivity (Weeks 1-4)
High Priority - Immediate Value

📌 Interfaz de Usuario Mejorada

Pestañas múltiples (essential for workflow efficiency)

✅ Autocompletado SQL (reduces errors, speeds up queries)

✅ Formateador SQL (improves readability instantly)

📂 Gestión de Consultas Básica

✅ Exportar resultados a CSV/Excel (critical for sharing data)

✅ Carpeta de queries por trabajo (basic organization)

⚡ Phase 2: Data Handling & Security (Weeks 5-8)
Medium Priority - Foundational Features
3. 📊 Soporte para Formatos

Importar/Exportar CSV/JSON (interoperability)

Conexión a DBs externas (MySQL, PostgreSQL) (expands use cases)

🔐 Seguridad Básica

Advertencia para consultas peligrosas (prevents accidents)

Variables en consultas (e.g., {{fecha}}) (adds flexibility)

📈 Phase 3: Advanced Features (Weeks 9-12)
High Value - Deeper Functionality
5. ⚡ Análisis de Datos Avanzado

Estadísticas rápidas (media, percentiles) (quick insights)

Limpieza de datos (nulos, normalización) (saves manual work)

📊 Visualización

Gráficos integrados (Matplotlib/Plotly) (helps non-SQL users)

🛠️ Generador Visual de Consultas

Drag-and-drop para JOINs/WHERE (beginners/low-code)

🌟 Phase 4: Long-Term (Months 3-6+)
Strategic & Scalability
8. 🚀 Automatización

Programar consultas recurrentes (e.g., daily reports)

📱 Multiplataforma

Versión web (Flask/Django + SQL.js) (broader access)

🔐 Rendimiento

Caché de consultas (optimization for heavy users)

👤 Perfiles de usuario

Guardar configuraciones/favoritos (personalization)
Programar consultas recurrentes (ejecutar cada día a una hora específica).

📱 Multiplataforma
Versión web (usando Flask/Django + SQL.js).


# **more fix**

better ui
auto to exel on export to exel 
auto to csv same as exel
fast fill fix like not geting stuck when moving and better reading on table
