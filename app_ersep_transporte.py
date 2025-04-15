
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import shutil
from io import BytesIO

# Configuración de la app
st.set_page_config(page_title="ERSEP Transporte", layout="centered")
st.title("🚌 Calculadora Tarifaria - ERSEP Transporte")
st.markdown("Modificá los valores que necesites y luego presioná **Enter** o el botón para obtener el cálculo tarifario completo.")

# Cargar datos de referencia desde la hoja "Hoja Llave"
@st.cache_data
def cargar_referencia():
    hoja = pd.read_excel("Incremento TBK_ Mesa 13 Octubre 2025.xlsx", sheet_name="Hoja Llave", header=None)
    nombres = hoja.iloc[:23, 0].tolist()
    valores_ref = hoja.iloc[:23, 14].tolist()
    nombres = [str(n) if pd.notna(n) else f"Dato {i+1}" for i, n in enumerate(nombres)]
    return nombres, valores_ref

# Aplicar los nuevos valores al archivo Excel y devolverlo
def actualizar_excel_con_datos(entradas_usuario):
    # Crear una copia del archivo original
    ruta_original = "Incremento TBK_ Mesa 13 Octubre 2025.xlsx"
    ruta_temp = "ersep_calculo_actualizado.xlsx"
    shutil.copy(ruta_original, ruta_temp)

    # Cargar la copia
    wb = load_workbook(ruta_temp, data_only=False)
    hoja = wb["Hoja Llave"]

    # Verificar si celda está combinada
    def es_combinada(hoja, fila, columna):
        for rango in hoja.merged_cells.ranges:
            if (fila, columna) in rango.cells:
                return True
        return False

    # Cargar datos en columna B, filas 3 a 25
    for i in range(23):
        fila = i + 3
        columna = 2
        if not es_combinada(hoja, fila, columna):
            hoja.cell(row=fila, column=columna).value = entradas_usuario[i]

    # Guardar y devolver objeto BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, wb

# Extraer resumen de resultados desde la hoja "Resumen de Calculo"
def obtener_resumen(wb):
    hoja_resumen = wb["Resumen de Calculo"]
    resumen = {}
    # Puedes ajustar este rango según tu estructura específica
    for fila in hoja_resumen.iter_rows(min_row=5, max_row=40, min_col=1, max_col=3):
        if fila[0].value and fila[2].value:
            categoria = str(fila[0].value).strip()
            valor = fila[2].value
            resumen[categoria] = valor
    return resumen

# Cargar referencias
nombres, valores_ref = cargar_referencia()
entradas_usuario = []

# Formulario para ingreso de datos
with st.form("formulario_datos"):
    for i in range(23):
        entrada = st.number_input(f"{nombres[i]}", value=float(valores_ref[i]) if isinstance(valores_ref[i], (int, float)) else 0.0, step=1.0, key=f"dato_{i}")
        entradas_usuario.append(entrada)
    calcular = st.form_submit_button("Calcular Tarifa")

# Cuando se aprieta Enter o el botón
if calcular:
    archivo_actualizado, wb_actualizado = actualizar_excel_con_datos(entradas_usuario)
    resumen = obtener_resumen(wb_actualizado)

    st.success("✅ Cálculo completado. Abajo se muestra el resumen tarifario.")
    st.write("### 📊 Resumen por Categorías")
    for categoria, valor in resumen.items():
        st.write(f"**{categoria}**: {valor}")

    st.download_button("📥 Descargar Excel completo con fórmulas",
                       data=archivo_actualizado,
                       file_name="Resumen_ERSEP_Tarifa.xlsx")
