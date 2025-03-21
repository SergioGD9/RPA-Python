import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os

# Ruta absoluta del archivo Excel
excel_path = r"C:\Users\sergi\Desktop\proyecto rpa\facturas.xlsx"
print(f"📄 Ruta del Excel: {excel_path}")

# Verificar si el archivo Excel existe antes de cargarlo
if not os.path.exists(excel_path):
    print(f"❌ Error: No se encontró el archivo {excel_path}")
    exit()

# Cargar datos en un DataFrame
df = pd.read_excel(excel_path)
print("✅ Archivo Excel cargado correctamente.")

# Obtener la ruta del directorio actual
script_dir = os.getcwd()
print(f"📂 Directorio actual: {script_dir}")

# Crear la carpeta de salida dentro del mismo directorio del script
output_folder = os.path.join(script_dir, "facturas_generadas")

try:
    os.makedirs(output_folder, exist_ok=True)
    print(f"📂 Carpeta creada: {output_folder}")
except Exception as e:
    print(f"❌ Error al crear la carpeta: {e}")

def generar_factura_pdf(empresa, direccion, productos, cantidades, precios, iva, factura_num):
    pdf_path = os.path.join(output_folder, f"Factura_{empresa.replace(' ', '_')}_{factura_num}.pdf")
    
    try:
        c = canvas.Canvas(pdf_path, pagesize=A4)
        c.setTitle(f"Factura {factura_num}")

        # Ruta del logo (ajústala según la ubicación real del archivo)
        logo_path = r"C:\Users\sergi\Desktop\proyecto rpa\489.jpg"

        # **Dibujar el logo en la parte superior izquierda**
        if os.path.exists(logo_path):
            c.drawImage(logo_path, 50, 760, width=100, height=50)  # Ajustar posición y tamaño del logo
        else:
            print(f"⚠️ Advertencia: No se encontró el logo en {logo_path}")

        # Márgenes y coordenadas base
        margin_x = 50  # Margen izquierdo
        margin_y = 700  # Ajuste de altura del encabezado

        # **Encabezado con borde (FACTURA, EMPRESA, DIRECCIÓN)**
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(300, margin_y + 25, f"Factura #{factura_num}")  # Título centrado
        c.setFont("Helvetica", 12)
        c.drawCentredString(300, margin_y, f"Empresa: {empresa}")
        c.drawCentredString(300, margin_y - 20, f"Dirección: {direccion}")

        # Ajuste del recuadro del encabezado (Empresa y Dirección)
        c.rect(margin_x, margin_y - 30, 500, 70)  # Cambia el segundo valor para bajar, el cuarto para aumentar altura

        # **Tabla de productos**
        y_position = margin_y - 90  # Ajuste de altura para la tabla de productos
        col_widths = [200, 80, 100, 120]  # Anchos de las columnas

        # **Encabezado de la tabla**
        c.setFont("Helvetica-Bold", 12)
        col_titles = ["Producto", "Cantidad", "Precio Unitario (€)", "Subtotal (€)"]

        # Ajustar posiciones de cada columna
        x_positions = [
            margin_x,  # Producto
            margin_x + col_widths[0] + 50,  # Cantidad
            margin_x + col_widths[0] + col_widths[1] + 50 ,  # Precio Unitario (€)
            margin_x + col_widths[0] + col_widths[1] + col_widths[2] + 80  # Subtotal (€) - Modificar este valor para moverlo
        ]

        # Dibujar encabezados de la tabla
        for i, title in enumerate(col_titles):
            c.drawString(x_positions[i] + 5, y_position, title)

        y_position -= 20  # Espacio antes de los productos

        # **Dibujar productos**
        total = 0
        c.setFont("Helvetica", 12)
        for producto, cantidad, precio in zip(productos, cantidades, precios):
            subtotal = cantidad * precio
            total += subtotal

            # Dibujar cada elemento alineado en su columna
            c.drawString(x_positions[0] + 5, y_position, producto)  # Producto
            c.drawRightString(x_positions[1] + 30, y_position, str(cantidad))  # Cantidad
            c.drawRightString(x_positions[2] + 60, y_position, f"{precio:.2f}")  # Precio Unitario
            c.drawRightString(x_positions[3] + 60, y_position, f"{subtotal:.2f}")  # Subtotal

            y_position -= 20  # Bajar una línea por cada producto agregado

        # **IVA y Total en un mismo recuadro**
        iva_total = total * (iva / 100)
        total_final = total + iva_total

        y_position -= 40  # Espacio antes del cuadro de IVA/Total
        cuadro_x = margin_x + 300  # Mueve el recuadro más a la derecha
        cuadro_ancho = 200  # Ancho del recuadro
        cuadro_alto = 50  # Altura del recuadro

        # Dibujar el recuadro del total e IVA
        c.rect(cuadro_x, y_position, cuadro_ancho, cuadro_alto)

        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(cuadro_x + cuadro_ancho / 2, y_position + 30, f"IVA ({iva}%) (€): {iva_total:.2f}")
        c.drawCentredString(cuadro_x + cuadro_ancho / 2, y_position + 10, f"Total (€): {total_final:.2f}")

        # **Mensaje final**
        y_position -= 40
        c.setFont("Helvetica", 10)
        c.drawCentredString(300, y_position, "Gracias por su compra!")

        c.save()
        print(f"✅ Factura generada: {pdf_path}")
    
    except Exception as e:
        print(f"❌ Error al generar el PDF: {e}")

# Procesar los datos del Excel y generar facturas
factura_num = 1
for empresa in df["Empresa"].unique():
    datos_empresa = df[df["Empresa"] == empresa]
    productos = datos_empresa["Producto"].tolist()
    cantidades = datos_empresa["Cantidad"].tolist()
    precios = datos_empresa["Precio Unitario (€)"].tolist()
    direccion = datos_empresa["Dirección"].iloc[0]
    iva = datos_empresa["IVA (%)"].iloc[0]
    
    generar_factura_pdf(empresa, direccion, productos, cantidades, precios, iva, factura_num)
    factura_num += 1

print("🎉 Todas las facturas han sido generadas correctamente.")
