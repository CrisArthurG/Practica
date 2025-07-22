from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, Image, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image as PILImage
from reportlab.pdfgen import canvas

# Configurar Tkinter (ocultar ventana principal)
root = tk.Tk()
root.withdraw()

def seleccionar_archivo(titulo="Seleccionar archivo", tipos_archivo=[("Imágenes", "*.jpg *.jpeg *.png *.bmp")]):
    return filedialog.askopenfilename(
        title=titulo,
        filetypes=tipos_archivo
    )

def solicitar_imagen(prompt, obligatorio=True):
    while True:
        print(prompt)
        ruta = seleccionar_archivo(titulo=prompt)
        
        if ruta:
            try:
                # Verificar si es imagen válida
                with PILImage.open(ruta) as img:
                    img.verify()
                return ruta
            except Exception as e:
                messagebox.showerror("Error", f"Archivo no válido: {str(e)}")
        elif not obligatorio:
            return None
        
        if not ruta and obligatorio:
            messagebox.showwarning("Archivo obligatorio", "Debe seleccionar un archivo")

def solicitar_imagen_opcional(prompt):
    respuesta = messagebox.askyesno("Logo adicional", prompt)
    if respuesta:
        return solicitar_imagen("Seleccione el logo adicional", obligatorio=False)
    return None

def solicitar_imagenes_asistencia():
    imagenes = []
    messagebox.showinfo("Imágenes del trabajo", "Adjunte imágenes del trabajo (mínimo 1, máximo 4)")
    
    while len(imagenes) < 4:
        img = solicitar_imagen(f"Seleccione la imagen {len(imagenes)+1} (mínimo 1)", obligatorio=(len(imagenes) == 0))
        if img:
            imagenes.append(img)
        
        if len(imagenes) >= 4:
            break
            
        if not messagebox.askyesno("Imagen adicional", "¿Desea agregar otra imagen?"):
            if len(imagenes) >= 1:  # Aseguramos mínimo 1 imagen
                break
    
    return imagenes

def crear_encabezado_pagina(canvas, doc, datos):
    """Función para dibujar el encabezado en cada página"""
    ancho, alto = A4
    
    # Dibujar logos
    x_logo = 1 * cm
    y_logo = alto - 2.5 * cm
    
    # Logo Equifar
    if datos['logo_equifar'] and os.path.exists(datos['logo_equifar']):
        try:
            canvas.drawImage(datos['logo_equifar'], x_logo, y_logo, width=2.5*cm, height=2.5*cm, preserveAspectRatio=True)
        except:
            pass
    
    # Logo adicional (si existe)
    if datos['logo_marca'] and os.path.exists(datos['logo_marca']):
        try:
            canvas.drawImage(datos['logo_marca'], x_logo + 2.5*cm, y_logo, width=2.5*cm, height=2.5*cm, preserveAspectRatio=True)
        except:
            pass
    
    # Número de asistencia (centrado)
    texto_numero = f"N° Asistencia: {datos['numero_asistencia']}"
    canvas.setFont("Helvetica-Bold", 12)
    canvas.drawCentredString(ancho/2, alto - 1.7*cm, texto_numero)
    
    # Información de contacto (derecha)
    contacto = [
        "Equifar S.A",
        "Fono: (56) 23 274 2230",
        "Email: info@equifar.cl",
        "www.equifar.cl"
    ]
    
    x_contacto = ancho - 1*cm
    y_contacto = alto - 1.5*cm
    canvas.setFont("Helvetica", 9)
    
    for i, linea in enumerate(contacto):
        canvas.drawRightString(x_contacto, y_contacto - i*0.5*cm, linea)

def generar_pdf(datos, imagenes, firma, nombre_archivo="asistencia.pdf"):
    # Crear documento con margen superior para el encabezado
    doc = SimpleDocTemplate(nombre_archivo, pagesize=A4, topMargin=3*cm)
    estilos = getSampleStyleSheet()
    
    # Crear estilos personalizados
    estilo_derecha = ParagraphStyle(name='Derecha', alignment=2, parent=estilos['Normal'])
    estilo_centrado = ParagraphStyle(name='Centrado', alignment=1, parent=estilos['Normal'])
    
    # Crear el contenido principal
    story = []
    
    # ========== CUERPO DEL DOCUMENTO ==========
    # Información general
    story.append(Paragraph(f"<b>Fecha:</b> {datos['fecha']}", estilos['Normal']))
    
    # Horas de trabajo
    story.append(Paragraph(f"<b>Horas de trabajo:</b> {datos['hora_inicio']} | {datos['hora_termino']}", estilos['Normal']))
    
    # Horas de traslado
    story.append(Paragraph(f"<b>Horas de traslado:</b> {datos['horas_traslado']}", estilos['Normal']))
    
    story.append(Paragraph(f"<b>Técnico Responsable:</b> {datos['tecnico']}", estilos['Normal']))
    story.append(Paragraph(f"<b>Técnicos Adicionales:</b> {', '.join(datos['tecnicos_adicionales'])}", estilos['Normal']))

    story.append(Spacer(1, 0.3*cm))

    # Información del cliente
    story.append(Paragraph(f"<b>Cliente:</b> {datos['cliente_nombre']}", estilos['Normal']))
    story.append(Paragraph(f"<b>Email:</b> {datos['cliente_email']}", estilos['Normal']))
    story.append(Paragraph(f"<b>Teléfono:</b> {datos['cliente_telefono']}", estilos['Normal']))
    story.append(Paragraph(f"<b>Dirección:</b> {datos['cliente_direccion']}", estilos['Normal']))
    story.append(Paragraph(f"<b>Observaciones Cliente:</b> {datos['observaciones_cliente']}", estilos['Normal']))

    story.append(Spacer(1, 0.3*cm))

    # Información de la máquina/equipo
    story.append(Paragraph(f"<b>Máquina/Equipo:</b> {datos['maquina_equipo']}", estilos['Normal']))
    story.append(Paragraph(f"<b>Modelo:</b> {datos['modelo_equipo']}", estilos['Normal']))
    story.append(Paragraph(f"<b>Número de serie:</b> {datos['numero_serie']}", estilos['Normal']))

    story.append(Spacer(1, 0.3*cm))

    # Servicio
    story.append(Paragraph(f"<b>Motivo de asistencia:</b> {datos['motivo']}", estilos['Normal']))
    story.append(Paragraph(f"<b>Descripción del trabajo:</b> {datos['descripcion_trabajo']}", estilos['Normal']))

    story.append(Spacer(1, 0.3*cm))

    # Repuestos
    story.append(Paragraph("<b>Repuestos:</b>", estilos['Normal']))
    if datos["repuestos"]:
        repuestos_tabla = [["Repuesto", "Cantidad"]] + [[r["nombre"], str(r["cantidad"])] for r in datos["repuestos"]]
        tabla = Table(repuestos_tabla, hAlign='LEFT')
        story.append(tabla)
    else:
        story.append(Paragraph("No se utilizaron repuestos.", estilos['Normal']))

    story.append(Spacer(1, 0.5*cm))

    # Imágenes del trabajo
    if imagenes:
        story.append(Paragraph("<b>Imágenes del trabajo:</b>", estilos['Normal']))
        
        # Preparar datos para la tabla de imágenes
        imagenes_data = []
        fila_actual = []
        
        for i, img_path in enumerate(imagenes):
            if os.path.exists(img_path):
                # Todas las imágenes con el mismo tamaño
                img_element = Image(img_path, width=8*cm, height=6*cm)
                fila_actual.append(img_element)
                
                # Cada 2 imágenes, creamos una nueva fila
                if len(fila_actual) == 2 or i == len(imagenes) - 1:
                    imagenes_data.append(fila_actual)
                    fila_actual = []
        
        # Crear tabla con las imágenes
        tabla_imagenes = Table(imagenes_data, colWidths=[8*cm, 8*cm])
        tabla_imagenes.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        
        story.append(tabla_imagenes)
        story.append(Spacer(1, 0.5*cm))

    # Firma del cliente
    story.append(Paragraph("<b>Firma del Cliente:</b>", estilos['Normal']))
    if firma and os.path.exists(firma):
        story.append(Image(firma, width=6*cm, height=3*cm))
    else:
        story.append(Paragraph("Firma no disponible.", estilos['Normal']))

    # Crear función de encabezado con los datos
    def encabezado(canvas, doc):
        crear_encabezado_pagina(canvas, doc, datos)
    
    # Construir el documento con el encabezado en cada página
    doc.build(story, onFirstPage=encabezado, onLaterPages=encabezado)

def generar_pdf_interactivo():
    print("📄 Generación de PDF de asistencia técnica\n")

    # Logos
    logo_equifar = solicitar_imagen("🔷 Seleccione el logotipo de EQUIFAR (obligatorio)")
    logo_marca = solicitar_imagen_opcional("¿Desea agregar un logotipo adicional de la marca?")

    # Datos simulados con los nuevos campos
    datos = {
        "numero_asistencia": "ATT-005",
        "fecha": "02/07/2025",
        "hora_inicio": "09:30",
        "hora_termino": "11:15",
        "horas_traslado": "1.5 horas",
        "tecnico": "Cristopher Arthur",
        "tecnicos_adicionales": ["Juan Pérez", "María López"],
        "cliente_nombre": "Farmacia Vida",
        "cliente_email": "contacto@farmaciavida.cl",
        "cliente_telefono": "+56 9 1234 5678",
        "cliente_direccion": "Av. Providencia 1001, Santiago",
        "observaciones_cliente": "Solicita revisión en segundo módulo.",
        "motivo": "Mantenimiento Preventivo",
        "descripcion_trabajo": "Se realizó mantenimiento del dispensador automático. Se cambiaron sensores y se ajustó calibración.",
        "maquina_equipo": "Dispensador automático",
        "modelo_equipo": "RX-2000",
        "numero_serie": "SN-123456789",
        "repuestos": [{"nombre": "Sensor óptico", "cantidad": 2}],
        "logo_equifar": logo_equifar,
        "logo_marca": logo_marca
    }

    # Imágenes del trabajo
    imagenes = solicitar_imagenes_asistencia()

    # Firma
    firma = solicitar_imagen("✍️ Seleccione la imagen de la firma del cliente: ")

    # Crear PDF
    nombre_archivo = "asistencia_generada.pdf"
    generar_pdf(datos, imagenes, firma, nombre_archivo)
    messagebox.showinfo("Éxito", f"PDF generado correctamente:\n{nombre_archivo}")

# Ejecutar el generador
if __name__ == "__main__":
    generar_pdf_interactivo()