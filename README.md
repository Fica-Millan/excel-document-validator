<p align="center">
  <img src="icono_chequeo.ico" width="20%" />
</p>

<div align="center">

![Python](https://img.shields.io/badge/Python-3.13-blue?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458?logo=pandas&logoColor=white)
![OpenPyXL](https://img.shields.io/badge/Excel-OpenPyXL-217346?logo=microsoftexcel&logoColor=white)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-orange)
![Platform](https://img.shields.io/badge/Platform-Windows-blue?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

![Status](https://img.shields.io/badge/Status-Terminado-green)
![Author](https://img.shields.io/badge/Autor-Yesica%20Fica%20Mill%C3%A1n-purple)

</div>

# Control de Documentación

Aplicación de escritorio desarrollada en Python para verificar la correspondencia entre los números internos de una planilla Excel y los archivos almacenados en una carpeta.

Está pensada para facilitar controles de documentación, por ejemplo, comprobar que las facturas en PDF registradas en un Excel estén disponibles e identificar archivos faltantes, mal nombrados o que no correspondan a ningún registro.

## 🔹​Funcionalidades

- **Control de documentación:** verifica si cada número interno del Excel tiene al menos un archivo asociado.
- **Identificación de faltantes:** muestra los números internos sin archivos y los datos complementarios de sus filas, como Fecha, COD y Comprobante.
- **Detección de archivos adicionales:** identifica archivos que no coinciden con ningún número interno del Excel.
- **Selección flexible de datos:** permite elegir la hoja, la columna del número interno y las columnas adicionales que se incluirán en el detalle.
- **Configuración de encabezados:** permite indicar la fila donde se encuentran los títulos de las columnas.
- **Búsqueda en subcarpetas:** permite incluir archivos ubicados dentro de subcarpetas.
- **Prefijo configurable:** permite agregar una letra al número interno para buscar archivos con nombres como `I14505 - Factura.pdf`.
- **Exportación a Excel:** genera un reporte con los resultados del control.

## 🔹​¿Cómo funciona?

La aplicación compara los números internos de una planilla Excel con los nombres de los archivos de una carpeta.

Por ejemplo, si el Excel contiene el número `14505` y el prefijo configurado es `I`, la aplicación buscará archivos cuyos nombres comiencen con `I14505`.

Si encuentra uno o más archivos que cumplen esa condición, el número se considera presente. Si no encuentra ninguno, lo informa como faltante.

También puede identificar archivos que no se asociaron con ningún número interno del Excel.

## 🔹​Requisitos de los archivos

Para realizar el control, se necesitan tres elementos:

1. **Un archivo Excel** con una columna que contenga los números internos. Por ejemplo, `14505`, sin la letra del prefijo si este se configura en la aplicación.
2. **Una carpeta con los documentos** cuyos nombres comiencen con el prefijo y el número interno correspondiente. Por ejemplo, `I14505 - Factura.pdf`.
3. **La fila de encabezados correctamente identificada**, ya que los títulos pueden estar en la primera fila o en otra ubicación de la planilla.

## 🔹​Instructivo de uso

### 1. Seleccionar el Excel

En la sección **Archivo Excel**, hacé clic en `Buscar...` y elegí la planilla que querés controlar.

En el campo **Fila de títulos**, indicá el número de fila donde están los nombres de las columnas. Por ejemplo, si los encabezados están en la fila 4, ingresá `4`.

### 2. Elegir la hoja y las columnas

En la sección **Hoja y columna con el número interno**:

- **Hoja:** seleccioná la hoja que contiene los registros.
- **Columna:** elegí la columna que contiene el número interno, sin el prefijo si este se agregará mediante la configuración.
- **Columnas a incluir en el detalle de faltantes:** marcá los datos adicionales que necesitás consultar cuando falte un archivo.

Por defecto, la aplicación selecciona las columnas correspondientes a Fecha, COD y Comprobante, siempre que existan en el Excel.

Para seleccionar o desmarcar varias columnas, mantené presionada la tecla `Ctrl` mientras hacés clic.

### 3. Seleccionar la carpeta

En la sección **Carpeta con los archivos**, hacé clic en `Buscar...` y elegí la carpeta donde se encuentra la documentación.

### 4. Configurar las opciones

- **Incluir subcarpetas:** activa la búsqueda de archivos dentro de las subcarpetas.
- **Mostrar también archivos que no corresponden a ningún número del Excel:** permite visualizar los archivos adicionales o sin correspondencia.
- **Prefijo a agregar al número del Excel antes de comparar:** establece la letra que se antepone al número interno. El valor predeterminado es `I`.

Si el Excel ya contiene el prefijo, dejá el campo vacío para evitar duplicarlo.

### 5. Ejecutar el control

Hacé clic en **Chequear**.

La aplicación mostrará el total de números internos encontrados en el Excel, la cantidad de archivos de la carpeta y los resultados de la comparación.

El resultado incluye dos listados:

- **FALTANTES:** números internos para los que no se encontró ningún archivo.
- **ARCHIVOS SIN NÚMERO CORRESPONDIENTE:** archivos que no coinciden con ningún número interno del Excel.

Las fechas del detalle de faltantes se muestran en formato `dd/mm/aa`.

### 6. Exportar el reporte

La exportación es opcional. Hacé clic en **Exportar reporte a Excel**, elegí la ubicación y el nombre del archivo.

El reporte generado contiene dos hojas:

| Hoja | Contenido |
|---|---|
| `Faltantes` | Números internos sin archivos asociados y los datos complementarios seleccionados. |
| `Archivos_extra` | Archivos que no se asociaron con ningún número interno. |

## 🔹​Problemas frecuentes

| Problema | Qué revisar |
|---|---|
| Todos los números figuran como faltantes. | Verificá el prefijo, la columna seleccionada y la carpeta elegida. |
| Todos los archivos figuran como adicionales. | Revisá la columna del número interno, la fila de títulos y el prefijo. |
| Las columnas aparecen como `Unnamed...`. | Corregí el número indicado en Fila de títulos. |
| Aparece el mensaje «Sin datos». | Comprobá que la columna seleccionada contenga números internos válidos. |
| No se detecta una factura que existe. | Verificá que el nombre del archivo comience con el prefijo y el número, sin espacios al principio. |
| No se detectan archivos de subcarpetas. | Activá la opción Incluir subcarpetas. |

**Consejo:** cada vez que cambies de Excel o de hoja, verificá nuevamente la fila de títulos y la columna del número interno antes de ejecutar el control.

## 🔹​Tecnologías utilizadas

- **Python:** lógica de la aplicación.
- **Tkinter y ttk:** interfaz gráfica de escritorio.
- **pandas:** lectura y procesamiento de los datos del Excel.
- **openpyxl:** soporte para la lectura y generación de archivos Excel.

## 🔹​Instalación y ejecución desde el código fuente

### Requisitos

- Python instalado en Windows.
- Las dependencias `pandas` y `openpyxl`.
- Tkinter, que suele estar incluido en las instalaciones habituales de Python para Windows.

### Instalar dependencias

Desde una terminal, ejecutá:

```bash
pip install pandas openpyxl
```

### Ejecutar la aplicación

Desde la carpeta del proyecto, ejecutá el archivo Python principal:

```bash
python chequeo_archivos_vs_excel.py
```

## 🔹​Ejecutar la versión para Windows (.exe)

El repositorio incluye una versión ejecutable de la aplicación para Windows.

1. Descargá el archivo `.exe` desde el repositorio de GitHub.
2. Guardalo en una carpeta de tu equipo.
3. Hacé doble clic en el ejecutable para abrir la aplicación.

La versión ejecutable permite utilizar la aplicación sin ejecutar manualmente el código Python ni instalar Python por separado, siempre que el paquete incluya las dependencias necesarias.

Si modificás el código fuente y querés distribuir los cambios, tenés que generar una nueva versión del ejecutable y reemplazar la anterior. El icono `icono_chequeo.ico` debe incluirse en el empaquetado para que aparezca en la aplicación compilada.

## 🔹​Alcance y limitaciones

- La comparación se basa en los nombres de los archivos, no en su contenido.
- Un número interno se considera presente cuando se encuentra al menos un archivo cuyo nombre comienza con el valor buscado.
- La aplicación no verifica si una factura es válida ni si el documento corresponde efectivamente al registro contable.
- Los archivos adicionales identificados deben revisarse para determinar si están mal nombrados o si no corresponden a la planilla.
- Se recomienda revisar los resultados antes de tomar decisiones sobre documentación pendiente.

## 🔹​Licencia

Este proyecto se distribuye bajo la licencia MIT, que permite utilizar, copiar, modificar y distribuir el software, incluso con fines comerciales, siempre que se conserve el aviso de copyright y el texto de la licencia.

Consultá el archivo `LICENSE` incluido en este repositorio para conocer los términos completos.
