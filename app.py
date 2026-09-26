import io
import re
from datetime import date

import streamlit as st
from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Registro y Facturación - Toyota Corregidora", layout="centered")

# ----------------------------------------------------------------------------
# Coordenadas del formato ANEXO 1 (Carta de Validación de Datos CFDI V 4.0)
# Este PDF NO tiene campos de formulario interactivos reales: es un PDF plano.
# Las coordenadas fueron extraídas directamente del PDF que subiste (página
# carta, 612x792 pts). Si Toyota cambia el formato, hay que re-calibrar esto.
# ----------------------------------------------------------------------------
PDF_TEXT_FIELDS = {
    # clave interna: (x, y, ancho_max, tamaño_fuente)
    "nombre": (40, 585, 490, 10),
    "regimen": (110, 565, 415, 10),
    "rfc": (60, 544, 140, 10),
    "cp": (246, 544, 115, 10),
    "movil": (68, 524, 140, 10),
    "correo": (308, 524, 225, 9),
}
CHECKBOX_S01 = (223, 445)          # centro de la casilla "S01 Sin efectos Fiscales"
SIGNATURE_NAME = (305.9, 98, 470)  # centro de la línea "Nombre / firma / fecha" (x, y, ancho_max)
# Rectángulo que cubre la casilla "WhatsApp": el PDF plantilla que compartiste ya
# trae una "x" fija ahí (no la pone este código); se blanquea esa zona para que
# el resultado final salga sin marcar. (x0, y0, ancho, alto, en coords. reportlab)
WHATSAPP_BOX_COVER = (226, 227, 24, 19)


def draw_fitted_text(c, x, y, text, max_width, base_size=10, min_size=6, font="Helvetica"):
    """Dibuja texto reduciendo el tamaño de fuente si no cabe en max_width."""
    if not text:
        return
    text = str(text)
    size = base_size
    while size > min_size and c.stringWidth(text, font, size) > max_width:
        size -= 0.5
    c.setFont(font, size)
    c.drawString(x, y, text)


def fill_pdf(template_bytes: bytes, data: dict) -> bytes:
    reader = PdfReader(io.BytesIO(template_bytes))
    page = reader.pages[0]
    width = float(page.mediabox.width)
    height = float(page.mediabox.height)

    overlay_buffer = io.BytesIO()
    c = canvas.Canvas(overlay_buffer, pagesize=(width, height))

    # Blanquea la casilla "WhatsApp" para tapar la "x" que trae la plantilla original,
    # y vuelve a dibujar el borde de la casilla (vacía) para no perder el recuadro.
    c.setFillColorRGB(1, 1, 1)
    c.rect(*WHATSAPP_BOX_COVER, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    c.setLineWidth(0.75)
    c.rect(*WHATSAPP_BOX_COVER, fill=0, stroke=1)

    for key, (x, y, max_w, size) in PDF_TEXT_FIELDS.items():
        draw_fitted_text(c, x, y, data.get(key, ""), max_w, base_size=size)

    # Marca fija requerida: S01 (Sin efectos Fiscales)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(*CHECKBOX_S01, "X")

    # Nombre del cliente en la línea de "Nombre / firma / fecha"
    x_sig, y_sig, max_w_sig = SIGNATURE_NAME
    nombre_txt = data.get("nombre", "")
    size = 10
    while size > 6 and c.stringWidth(nombre_txt, "Helvetica", size) > max_w_sig:
        size -= 0.5
    c.setFont("Helvetica", size)
    c.drawCentredString(x_sig, y_sig, nombre_txt)

    c.save()
    overlay_buffer.seek(0)

    overlay_reader = PdfReader(overlay_buffer)
    page.merge_page(overlay_reader.pages[0])

    writer = PdfWriter()
    writer.add_page(page)
    # Si el PDF original tuviera más páginas, se agregan tal cual
    for p in reader.pages[1:]:
        writer.add_page(p)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def adjust_formula_row(formula: str, old_row: int, new_row: int) -> str:
    """Ajusta referencias relativas de fila (ej. B2 -> B15) dentro de una fórmula,
    sin tocar referencias absolutas como $C$2:$C$201 o Config!$B$1."""
    pattern = re.compile(r"(?<![A-Za-z0-9$])([A-Z]{1,3})" + str(old_row) + r"(?![0-9])")
    return pattern.sub(lambda m: f"{m.group(1)}{new_row}", formula)


def next_empty_row(ws, key_col: int, template_row: int, max_row: int) -> int:
    row = template_row
    while ws.cell(row=row, column=key_col).value not in (None, ""):
        row += 1
        if row > max_row:
            raise ValueError(
                f"La hoja '{ws.title}' ya no tiene filas libres dentro de la tabla "
                f"(límite fila {max_row}). Agrega más filas a la tabla en Excel."
            )
    return row


def copy_row_formulas(ws, template_row: int, new_row: int, max_col: int):
    """Copia SOLO las celdas con fórmula de la fila plantilla a la fila nueva,
    ajustando las referencias relativas de fila. Las celdas de datos (valores)
    se dejan intactas para llenarlas explícitamente después."""
    for col in range(1, max_col + 1):
        src = ws.cell(row=template_row, column=col)
        if isinstance(src.value, str) and src.value.startswith("="):
            new_formula = adjust_formula_row(src.value, template_row, new_row)
            dst = ws.cell(row=new_row, column=col)
            dst.value = new_formula
            dst.number_format = src.number_format


def update_excel(excel_bytes: bytes, datos: dict) -> bytes:
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=False)

    # ---- Hoja "Clientes" ----
    if "Clientes" in wb.sheetnames:
        ws_cli = wb["Clientes"]
        row = next_empty_row(ws_cli, key_col=2, template_row=2, max_row=201)  # col B = Nombre
        copy_row_formulas(ws_cli, template_row=2, new_row=row, max_col=ws_cli.max_column)
        ws_cli.cell(row=row, column=2, value=datos["nombre"])          # Nombre
        ws_cli.cell(row=row, column=3, value=datos["telefono_wa"])     # Teléfono (para wa.me, solo dígitos)
        ws_cli.cell(row=row, column=4, value=datos["unidad"])          # Vehículo
        ws_cli.cell(row=row, column=14, value=date.today())            # Fecha Registro

    # ---- Hoja "Facturación" ----
    if "Facturación" in wb.sheetnames:
        ws_fac = wb["Facturación"]
        row = next_empty_row(ws_fac, key_col=1, template_row=2, max_row=101)  # col A = Cliente
        copy_row_formulas(ws_fac, template_row=2, new_row=row, max_col=ws_fac.max_column)
        valores = [
            datos["nombre"], datos["rfc"], datos["domicilio"], datos["uso_cfdi"],
            datos["regimen"], datos["telefono"], datos["correo"], datos["contacto_emergencia"],
            datos["conductor_nombre"], datos["conductor_tel"], datos["conductor_correo"],
            datos["resp_nombre"], datos["resp_tel"], datos["resp_correo"],
            datos["unidad"], datos["valor_factura"],
        ]
        for i, val in enumerate(valores, start=1):
            ws_fac.cell(row=row, column=i, value=val)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def build_whatsapp_message(d: dict) -> str:
    return f"""La factura se realizará a nombre de: {d['nombre']}
RFC: {d['rfc']}
Domicilio fiscal conforme a la constancia: {d['domicilio']} CP. {d['cp']}
Uso de CFDI: {d['uso_cfdi']}
Régimen: {d['regimen']}
Tel: {d['telefono']}
Correo: {d['correo']}
Contacto de emergencias. {d['contacto_emergencia']}

Conductor: {d['conductor_nombre']}
Tel: {d['conductor_tel']}
Correo: {d['conductor_correo']}
Responsable Mto: {d['resp_nombre']}
Tel: {d['resp_tel']}
Correo: {d['resp_correo']}

Se factura una unidad: {d['unidad']}
Valor factura: {d['valor_factura']}

Correcto?"""


# ==============================================================================
# INTERFAZ
# ==============================================================================
st.title("Registro y Facturación · Toyota Corregidora")
st.caption("Sube tus archivos base, llena los datos del cliente y genera todo en un paso.")

col_up1, col_up2 = st.columns(2)
with col_up1:
    excel_file = st.file_uploader("Excel de seguimiento (.xlsx)", type=["xlsx"])
with col_up2:
    pdf_file = st.file_uploader("Formato ANEXO 1 - CFDI (.pdf)", type=["pdf"])

st.divider()

with st.form("form_cliente", clear_on_submit=False):
    st.subheader("Datos de Facturación")
    c1, c2 = st.columns(2)
    with c1:
        nombre = st.text_input("Nombre Cliente / Razón Social")
        rfc = st.text_input("RFC")
        domicilio = st.text_input("Domicilio Fiscal Completo")
        cp = st.text_input("C.P.")
    with c2:
        uso_cfdi = st.text_input("Uso de CFDI", value="S01 Sin efectos Fiscales")
        regimen = st.text_input("Régimen Fiscal")
        telefono = st.text_input("Teléfono")
        correo = st.text_input("Correo Electrónico")
    contacto_emergencia = st.text_input("Contacto de Emergencia")

    st.subheader("Datos de Personal")
    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**Conductor**")
        conductor_nombre = st.text_input("Nombre (conductor)")
        conductor_tel = st.text_input("Tel (conductor)")
        conductor_correo = st.text_input("Correo (conductor)")
    with c4:
        st.markdown("**Responsable de Mantenimiento**")
        resp_nombre = st.text_input("Nombre (responsable)")
        resp_tel = st.text_input("Tel (responsable)")
        resp_correo = st.text_input("Correo (responsable)")

    st.subheader("Datos de la Unidad")
    unidad = st.text_input("Descripción de la unidad", placeholder='Ej. "Hilux SR, 2027, color gris, manual"')
    valor_factura = st.text_input("Valor factura")

    submitted = st.form_submit_button("Procesar y Generar Archivos", use_container_width=True)

if submitted:
    if not excel_file or not pdf_file:
        st.error("Sube primero el Excel de seguimiento y el PDF del ANEXO 1.")
    elif not nombre or not rfc:
        st.error("Nombre Cliente y RFC son obligatorios.")
    else:
        telefono_wa = re.sub(r"\D", "", telefono)  # solo dígitos, para wa.me
        datos = dict(
            nombre=nombre, rfc=rfc, domicilio=domicilio, cp=cp, uso_cfdi=uso_cfdi,
            regimen=regimen, telefono=telefono, telefono_wa=telefono_wa, correo=correo,
            contacto_emergencia=contacto_emergencia, conductor_nombre=conductor_nombre,
            conductor_tel=conductor_tel, conductor_correo=conductor_correo,
            resp_nombre=resp_nombre, resp_tel=resp_tel, resp_correo=resp_correo,
            unidad=unidad, valor_factura=valor_factura, movil=telefono,
        )
        try:
            excel_out = update_excel(excel_file.getvalue(), datos)
            pdf_out = fill_pdf(pdf_file.getvalue(), datos)
            mensaje = build_whatsapp_message(datos)

            st.success(f"Listo. Se agregó el registro de **{nombre}**.")

            st.subheader("Descargas")
            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    "⬇️ Excel actualizado",
                    data=excel_out,
                    file_name=excel_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with d2:
                st.download_button(
                    "⬇️ PDF ANEXO 1 rellenado",
                    data=pdf_out,
                    file_name=f"Anexo1_{nombre.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

            st.subheader("Mensaje para WhatsApp")
            st.text_area("Copia este texto", value=mensaje, height=380)

        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Ocurrió un error procesando los archivos: {e}")
