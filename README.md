# 📄 Generador de Documentos Word y Resumen de Datos

## 📝 Descripción

Este script de Python 🐍 se utiliza para generar documentos de Word 📝 a partir de un archivo CSV 📊 que contiene cierta información. Además, actualiza un archivo CSV de resumen con las rutas de los archivos de Word generados. Finalmente, ordena el DataFrame por una columna específica antes de guardarlo.

## 🛠️ Dependencias

- Python 3.x 🐍
- pandas 🐼
- python-docx 📝
- os 💻

## 🔄 Funcionamiento

1. **📂 Directorio**: Se establece el directorio donde se encuentra el archivo CSV y donde se guardarán los documentos de Word generados.
2. **🔍 Buscar archivo CSV**: Busca un archivo `.csv` en el directorio especificado que contenga el texto 'resumen.csv' en su nombre.
3. **📖 Leer archivo CSV**: Lee el archivo `.csv` y selecciona columnas específicas en un DataFrame.
4. **📝 Generar documentos de Word**: Itera sobre las filas del DataFrame y genera un documento de Word para cada una, basándose en la información de ciertas columnas.
5. **🔄 Actualizar DataFrame**: Agrega la ruta de cada documento de Word generado a una nueva columna en el DataFrame.
6. **🔢 Ordenar DataFrame**: Ordena el DataFrame por la columna 'km_a_reclamar' en orden descendente.
7. **💾 Guardar DataFrame actualizado**: Guarda el DataFrame actualizado en un nuevo archivo `.csv`.

## 🚀 Cómo usar

1. Asegúrese de tener todas las dependencias instaladas 🛠️.
2. Cambie la variable `dir_path` al directorio en el que desea trabajar 📂.
3. Ejecute el script 🚀.

