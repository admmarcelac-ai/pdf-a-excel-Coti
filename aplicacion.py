import streamlit as st
from PyPDF2 import PdfReader
import pandas as pd
from io import BytesIO
import re

# =========================
# ✅ INTERFAZ
# =========================

st.title("PDF a Excel")

st.markdown("Subí archivos PDF de facturas")

archivos = st.file_uploader(
    "Seleccionar PDFs",
    type="pdf",
    accept_multiple_files=True
)

procesar = st.button("Procesar facturas")

if not archivos:
    st.info("Esperando que cargues archivos...")

# =========================
# ✅ FUNCIÓN
# =========================

def procesar_pdf(archivo):

    filas = []
    reader = PdfReader(archivo)
    texto = reader.pages[0].extract_text()

    if not texto:
        return filas

    fecha = re.search(r"\d{2}/\d{2}/\d{4}", texto)
    fecha = fecha.group(0) if fecha else ""

    tipo = re.search(r"FACTURA\s+([ABC])", texto)
    tipo = tipo.group(1) if tipo else ""

    cuits = re.findall(r"\d{11}", texto)
    cuit_emisor = cuits[0] if len(cuits) > 0 else ""
    cuit_receptor = cuits[1] if len(cuits) > 1 else ""

    razon_emisor = ""
    for linea in texto.split("\n"):
        if ("SRL" in linea or "S.A" in linea or "SA" in linea):
            razon_emisor = linea.strip()
            break

    razon_receptor = ""
    for linea in texto.split("\n"):
        if cuit_receptor in linea:
            razon_receptor = linea.replace(cuit_receptor, "").strip()
            break

    m = re.search(
        r"Punto de Venta:\s*Comp\.?\s*Nro:\s*(\d+)\s*(\d+)", texto
    )

    punto_venta = m.group(1) if m else ""
    numero = m.group(2) if m else ""

    if "Código Producto" in texto:
        texto = texto.split("Código Producto", 1)[1]

    lineas = [l.strip() for l in texto.split("\n") if l.strip()]
    buffer_desc = []

    for linea in lineas:

        if "unidades" in linea:

            cantidad = 0

            match = re.search(r"(.+?)unidades", linea)

            if match:
                bloque = match.group(1)
                numeros = re.findall(r"(\d+),\d+", bloque)

                if numeros:
                    cantidad_str = numeros[-1]

                    if len(cantidad_str) <= 3:
                        cantidad = int(cantidad_str)
                    else:
                        cantidad = int(cantidad_str[-3:])

            numeros_linea = re.findall(r"\d+,\d+", linea)

            if buffer_desc:
                producto = " ".join(buffer_desc).strip()
            else:
                producto = re.split(r"\d+,\d+", linea)[0].strip()

            if len(numeros_linea) >= 4:
                precio = float(numeros_linea[1].replace(",", "."))
                subtotal = float(numeros_linea[3].replace(",", "."))
                total = float(numeros_linea[-1].replace(",", "."))
            elif len(numeros_linea) >= 3:
                precio = float(numeros_linea[1].replace(",", "."))
                subtotal = float(numeros_linea[2].replace(",", "."))
                total = float(numeros_linea[-1].replace(",", "."))
            else:
                precio = subtotal = total = 0

            # ✅ corrección contable
            if precio > 0:
                cantidad_calc = round(subtotal / precio)
                if cantidad == 0 or abs(cantidad - cantidad_calc) > 1:
                    cantidad = cantidad_calc

            # ✅ validación
            if precio > 0:
                diferencia = round((cantidad * precio) - subtotal, 2)
                validacion = "OK" if abs(diferencia) <= 1 else "ERROR"
            else:
                validacion = "SIN PRECIO"

            producto = re.sub(r"\s+", " ", producto).strip()
            if len(producto) < 3:
                producto = "SIN DESCRIPCIÓN"

            filas.append({
                "Fecha": fecha,
                "CUIT Emisor": cuit_emisor,
                "Razón Emisor": razon_emisor,
                "CUIT Receptor": cuit_receptor,
                "Razón Receptor": razon_receptor,
                "Tipo": tipo,
                "Punto de Venta": punto_venta,
                "Número": numero,
                "Producto": producto,
                "Cantidad": cantidad,
                "Precio Unitario": precio,
                "Subtotal": subtotal,
                "Total c/ IVA": total,
                "Validación": validacion
            })

            buffer_desc = []

        else:
            if not any(x in linea for x in ["Código", "Subtotal", "IVA", "CAE", "%"]):
                buffer_desc.append(linea)

    return filas

# =========================
# ✅ PROCESO
# =========================

if archivos and procesar:

    todas = []

    for pdf in archivos:
        todas.extend(procesar_pdf(pdf))

    if todas:

        df = pd.DataFrame(todas)

        columnas = [
            "Fecha",
            "CUIT Emisor",
            "Razón Emisor",
            "CUIT Receptor",
            "Razón Receptor",
            "Tipo",
            "Punto de Venta",
            "Número",
            "Producto",
            "Cantidad",
            "Precio Unitario",
            "Subtotal",
            "Total c/ IVA",
            "Validación"
        ]

        df = df[columnas]

        st.dataframe(df)

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")

        st.download_button(
            "Descargar Excel",
            buffer.getvalue(),
            "facturas.xlsx"
        )

    else:
        st.warning("No se detectaron datos.")
