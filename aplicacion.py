import streamlit as st
from PyPDF2 import PdfReader
import pandas as pd
from io import BytesIO
import re

# =========================
# ✅ INTERFAZ
# =========================

st.title("PDF a Excel")

st.markdown("### Subí archivos PDF de facturas")
st.markdown("Podés cargar una o varias facturas en formato PDF")

archivos = st.file_uploader(
    "Seleccionar PDFs",
    type="pdf",
    accept_multiple_files=True
)

procesar = st.button("Procesar facturas")

if not archivos:
    st.info("Esperando que cargues archivos...")

# =========================
# ✅ FUNCIÓN PRINCIPAL
# =========================

def procesar_pdf(archivo):

    filas = []
    reader = PdfReader(archivo)
    texto = reader.pages[0].extract_text()

    if not texto:
        return filas

    # =========================
    # ✅ DATOS GENERALES
    # =========================

    fecha = re.search(r"\d{2}/\d{2}/\d{4}", texto)
    fecha = fecha.group(0) if fecha else ""

    tipo = re.search(r"FACTURA\s+([ABC])", texto)
    tipo = tipo.group(1) if tipo else ""

    cuits = re.findall(r"\d{11}", texto)
    cuit_emisor = cuits[0] if len(cuits) > 0 else ""
    cuit_receptor = cuits[1] if len(cuits) > 1 else ""

    # =========================
    # ✅ RAZÓN EMISOR
    # =========================

    razon_emisor = ""

    match_rs = re.search(
        r"Razón Social:\s*([A-Z0-9 .]+?)\s*(Condición|Domicilio|CUIT|FACTURA)",
        texto
    )

    if match_rs:
        razon_emisor = match_rs.group(1).strip()

    if not razon_emisor:
        for linea in texto.split("\n"):
            if ("SRL" in linea or "S.A" in linea or "SA" in linea):
                razon_emisor = linea.strip()
                break

    # =========================
    # ✅ RECEPTOR
    # =========================

    razon_receptor = ""

    for linea in texto.split("\n"):
        if cuit_receptor in linea:
            razon_receptor = linea.replace(cuit_receptor, "").strip()
            break

    # =========================
    # ✅ PV + Nº
    # =========================

    m = re.search(
        r"Punto de Venta:\s*Comp\.?\s*Nro:\s*(\d+)\s*(\d+)", texto
    )

    punto_venta = m.group(1) if m else ""
    numero = m.group(2) if m else ""

    # =========================
    # ✅ DETALLE
    # =========================

    if "Código Producto" in texto:
        texto = texto.split("Código Producto", 1)[1]

    lineas = [l.strip() for l in texto.split("\n") if l.strip()]
    buffer_desc = []

    for linea in lineas:

        if "unidades" in linea:

            # =========================
            # ✅ CANTIDAD DESDE TEXTO
            # =========================

            cantidad = 0

            match = re.search(r"(.+?)unidades", linea)

            if match:
                bloque = match.group(1)
                numeros = re.findall(r"(\d+),\d+", bloque)

                if numeros:
                    cantidad_str = numeros[-1]

                    if len(cantidad_str) <= 3:
                        cantidad = int(cantidad_str)
                    elif len(cantidad_str) == 4:
                        cantidad = int(cantidad_str[-3:])
                    else:
                        cantidad = int(cantidad_str[-3:])

            # =========================
            # ✅ PRODUCTO
            # =========================

            numeros_linea = re.findall(r"\d+,\d+", linea)

            if buffer_desc:
                producto = " ".join(buffer_desc).strip()
            else:
                producto = re.split(r"\d+,\d+", linea)[0].strip()

            # =========================
            # ✅ IMPORTES
            # =========================

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

            # =========================
            # ✅ CORRECCIÓN CONTABLE
            # =========================

            if precio > 0:
                cantidad_calc = round(subtotal / precio)

                if cantidad == 0 or abs(cantidad - cantidad_calc) > 1:
                    cantidad = cantidad_calc

            # =========================
            # ✅ VALIDACIÓN
            # =========================

            if precio > 0:
                diferencia = round((cantidad * precio) - subtotal, 2)

                if abs(diferencia) <= 1:
                    validacion = "OK"
                else:
                    validacion = "ERROR"
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
# ✅ PROCESO GLOBAL
# =========================

if archivos and procesar:

    todas = []

    for pdf in archivos:
        todas.extend(procesar_pdf(pdf))

    if todas:

        df = pd.DataFrame(todas)

        # =========================
        # ✅ TOTAL POR COMPROBANTE
        # =========================

        df["Total Factura"] = df.groupby(
            ["Fecha", "CUIT Emisor", "Tipo", "Punto de Venta", "Número"]
        )["Total c/ IVA"].transform("first")

        # =========================
        # ✅ RESUMEN
        # =========================

        resumen = df.groupby(
            ["Fecha", "CUIT Emisor", "Razón Emisor",
             "CUIT Receptor", "Razón Receptor",
             "Tipo", "Punto de Venta", "Número"],
            as_index=False
        ).agg({
            "Subtotal": "sum",
            "Total c/ IVA": "max"
        })

        resumen.rename(columns={
            "Subtotal": "Subtotal Factura",
            "Total c/ IVA": "Total Factura"
        }, inplace=True)

        # =========================
        # ✅ ORDEN COLUMNAS DETALLE
        # =========================

        columnas_ordenadas = [
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
            "Total Factura",
            "Validación"
        ]

        df = df[columnas_ordenadas]

        # =========================
        # ✅ MOSTRAR
        # =========================

        st.subheader("Detalle de productos")
        st.dataframe(df)

        st.subheader("Resumen por comprobante")
        st.dataframe(resumen)

        # =========================
        # ✅ EXPORTAR
        # =========================

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")

        st.download_button(
            "Descargar Excel",
            buffer.getvalue(),
            "facturas_combinadas.xlsx"
        )

    else:
        st.warning("No se detectaron datos.")
