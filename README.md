# RPA-Python

# Generador de Facturas en PDF

Este proyecto automatiza la generación de facturas en formato PDF a partir de un archivo de datos en Excel. El script toma los datos de las empresas, productos, cantidades, precios unitarios y otras informaciones, y genera un PDF con el formato estándar de una factura, incluyendo IVA y total.

## Requisitos

Asegúrate de tener instaladas las siguientes librerías:

- `pandas`: Para leer y manejar los datos del archivo Excel.
- `reportlab`: Para generar los archivos PDF.

Puedes instalarlas usando `pip`

# Estructura del Proyecto

- `facturas.xlsx`: Archivo Excel que contiene los datos para generar las facturas. Este archivo debe tener las siguientes columnas:
  - **Empresa**: Nombre de la empresa.
  - **Dirección**: Dirección de la empresa.
  - **Producto**: Nombre del producto.
  - **Cantidad**: Cantidad de productos vendidos.
  - **Precio Unitario (€)**: Precio unitario de los productos.
  - **IVA (%)**: Porcentaje de IVA aplicado a la factura.
- `generar_factura.py`: Script principal que lee el archivo Excel, procesa los datos y genera las facturas en formato PDF.

## Uso

1. Coloca el archivo `facturas.xlsx` en la ruta especificada en el script (puedes modificar la ruta del archivo en el código según sea necesario).
2. Ejecuta el script `generar_factura.py` para generar las facturas en formato PDF.
3. Las facturas se guardarán en una carpeta llamada `facturas_generadas` en el mismo directorio donde se ejecuta el script.

## Descripción del Código

### 1. Cargar y procesar datos

El script comienza cargando el archivo Excel con `pandas`. Se verifica si el archivo existe y si es posible leer los datos. Si el archivo no existe, se muestra un mensaje de error y el script termina.

### 2. Generación de facturas

La función `generar_factura_pdf` se encarga de crear un archivo PDF para cada empresa listada en el archivo Excel. Se incluye el logo de la empresa, un encabezado con los detalles de la factura (número de factura, nombre de la empresa y dirección), una tabla de productos con sus precios, cantidades y subtotales, y el cálculo del IVA y total de la factura.

Cada factura se guarda en el directorio `facturas_generadas` con un nombre basado en la empresa y el número de factura.

### 3. Carpeta de salida

El script crea una carpeta llamada `facturas_generadas` en el directorio actual del script para almacenar los archivos PDF generados.

## Personalización

- **Logo**: Asegúrate de actualizar la ruta del logo en el script para que apunte al archivo correcto.
- **Ruta de entrada/salida**: Puedes modificar las rutas de entrada y salida en el script para adaptarlas a tu sistema.
- **Formato de factura**: Puedes personalizar el formato de la factura modificando el código dentro de la función `generar_factura_pdf`.

## Ejemplo de salida

Una vez ejecutado el script, se generarán archivos PDF con el siguiente formato de nombre:

Factura_<Nombre_Empresa>_<Numero_Factura>.pdf

Cada archivo PDF contendrá los detalles de la factura correspondiente.

## Contribuciones

Si deseas contribuir a este proyecto, por favor, crea un fork del repositorio y envía un pull request con tus cambios.
