
---

### 📘 `README.md` – Excel SQL Analyzer Pro

````markdown
# 📊 Excel SQL Analyzer Pro

Aplicación de escritorio en PyQt5 para analizar archivos Excel (`.xlsx`) usando consultas SQL. Permite combinar múltiples hojas, visualizar resultados, guardar consultas personalizadas y crear dashboards simples con estadísticas y gráficos.

---

## 🚀 Características

- Carga múltiple de archivos y hojas Excel
- Ejecución de consultas SQL sobre los datos
- Resaltado de sintaxis y autocompletado inteligente
- Filtro rápido en tabla
- Visualización estadística (`describe`) y gráficos (`matplotlib`)
- Exportación de resultados a CSV / Excel
- Historial de consultas organizadas por carpetas
- Modo oscuro activado por defecto 🌙

---

## 💻 Requisitos

- Python 3.10 o superior
- Sistema operativo: probado en Arch Linux y otras distros basadas en Linux
- Entorno virtual (`venv`) recomendado

---

## 🧰 Instalación

### 1. Clonar repositorio

```bash
git clone https://github.com/usuario/excel-sql-analyzer.git
cd excel-sql-analyzer
````

### 2. Crear entorno virtual

```bash
python -m venv venv
source venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

Si no tienes el archivo `requirements.txt`, instala manualmente:

```bash
pip install pyqt5 pandas openpyxl matplotlib
```

---

## ▶️ Ejecutar la aplicación

Con el entorno virtual activo:

```bash
python exelsql.py
```

---

## 📝 Uso general

1. Haz clic en **Cargar Excel(s)** y selecciona uno o más archivos `.xlsx`.
2. Selecciona una hoja desde el combo desplegable.
3. Visualiza los datos automáticamente.
4. Escribe consultas SQL en el editor (ej. `SELECT * FROM archivo_hoja`).
5. Ejecuta con **Ctrl+R** o con el botón `Ejecutar SQL`.
6. Usa `Visualizar Gráfico` o `Resumen Estadístico` para análisis rápido.
7. Exporta los resultados a `.csv` o `.xlsx`.

---

## 🧠 Consejos

* Las hojas se renombran automáticamente como `nombrearchivo_nombredelaHoja` para evitar colisiones.
* Usa el menú `☰ Plantillas SQL` para insertar rápidamente comandos comunes.
* Guarda tus consultas y organízalas en carpetas.
* Puedes mover o eliminar consultas desde el historial.

---

## 🐞 Problemas comunes

### `ModuleNotFoundError: No module named 'matplotlib'`

Solución:

```bash
pip install matplotlib
```

---

## 📂 Estructura recomendada del proyecto

```text
excel-sql-analyzer/
├── exelsql.py               # Código principal
├── query_history.json       # (autogenerado) historial de consultas
├── README.md                # este archivo
└── venv/                    # entorno virtual
```

---
