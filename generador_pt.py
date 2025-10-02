# empieza codigo
import streamlit as st
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime
import numpy as np
import pandas as pd
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import textwrap
import math

def mostrar_app():
    st.set_page_config(page_title="Generador CTG - Transformador de Tensión", layout="wide")

    st.title("📄 Generador de Ficha CTG")
    st.subheader("Transformador de Tensión")

    st.markdown("### ⚙️ Parámetros del transformador")

    # 1 al 5
    fabricante = st.text_input("1. Fabricante", key="param_01_fabricante")
    pais = st.text_input("2. País", key="param_02_pais")
    referencia = st.text_input("3. Referencia", key="param_03_referencia")
    norma_fabricacion = st.text_input("4. Norma de fabricación", value="IEC 61869-5", key="param_04_norma_fabricacion")
    norma_calidad = st.text_input("5. Norma de calidad", value="ISO 9001", key="param_05_norma_calidad")

    # 6 al 9
    tipo_ejecucion = st.selectbox("6. Tipo de ejecución", ["Interior", "Exterior"], key="param_06_tipo_ejecucion")
    altura_instalacion = st.number_input("7. Altura de instalación (msnm)", min_value=0, step=100, key="param_07_altura")
    material_aislador = st.selectbox("8. Material del aislador", ["Compuesto siliconado", "Porcelana"], key="param_08_material_aislador")
    tipo_transformador = st.selectbox("8a. Tipo", ["Capacitivo", "Inductivo"], key="param_08a_tipo_transformador")
    tension_um = st.selectbox("9. Tensión más elevada para el material (Um)", ["145 kV", "245 kV", "550 kV"], key="param_09_um")

    # 10. Ud
    st.markdown("### 🔌 10. Tensión asignada soportada a la frecuencia industrial (Ud)")
    ud_interno = {"145 kV": "360 kV", "245 kV": "460 kV", "550 kV": "700 kV"}[tension_um]
    st.text(f"Aislamiento Interno: {ud_interno}")
    st.text(f"Aislamiento Externo (*): {ud_interno} a {int(altura_instalacion)} msnm")

    # 11. Up
    st.markdown("### ⚡ 11. Tensión asignada soportada al impulso tipo rayo (Up)")
    up_interno = {"145 kV": "750 kV", "245 kV": "1050 kV", "550 kV": "1550 kV"}[tension_um]
    st.text(f"Aislamiento Interno: {up_interno}")
    st.text(f"Aislamiento Externo (*): {up_interno} a {int(altura_instalacion)} msnm")

    # 12. Us
    st.markdown("### ⚡ 12. Tensión asignada soportada al impulso tipo maniobra (Us)")
    us_interno = st.text_input("Aislamiento Interno (dejar vacío por ahora)", key="param_12_us_interno")
    us_externo = st.text_input("Aislamiento Externo (*) (dejar vacío por ahora)", key="param_12_us_externo")

    # 13. Frecuencia
    st.markdown("### 📶 13. Frecuencia asignada (fr)")
    st.text("60 Hz")

    # 14. Factor de tensión
    st.markdown("### ⚙️ 14. Factor de tensión asignado")
    st.text("a) Permanente: 1,2")
    st.text("b) Durante 30 s: 1,5")

    # 15. Capacidad total
    st.markdown("### ⚡ 15. Capacidad total")
    capacidad_total = st.number_input("Capacidad total (≥ 4000 VA)", min_value=4000, key="param_15_capacidad_total")

    # 16 al 18
    st.markdown("### 🔧 16-18. Condensadores y tensión intermedia")
    c1 = st.text_input("16. Condensador de alta tensión (C1)", key="param_16_c1")
    c2 = st.text_input("17. Condensador de tensión intermedia (C2)", key="param_17_c2")
    tension_intermedia = st.text_input("18. Tensión intermedia asignada en circuito abierto", key="param_18_tension_intermedia")

    # 19. Número de devanados secundarios
    st.markdown("### 🔁 19. Número de devanados secundarios")
    num_devanados = st.selectbox("Selecciona el número de devanados secundarios", [1, 2, 3], key="param_19_num_devanados")

    # 20. Clase de precisión
    st.markdown("### 🎯 20. Clase de precisión")
    st.markdown("**Entre el 25% y el 100% de la carga de precisión con factor de potencia 0,8 en atraso**")
    clase_precision_a = st.selectbox("a) Entre el 5% y el 80% de la tensión asignada", ["1P", "2P", "3P", "4P", "5P"], key="param_20a_clase_precision")
    clase_precision_b = st.selectbox("b) Entre el 80% y el 120% de la tensión asignada", ["0.1", "0.2", "0.3"], key="param_20b_clase_precision")
    clase_precision_c = st.selectbox("c) Entre el 120% y el 150% de la tensión asignada", ["1P", "2P", "3P", "4P", "5P"], key="param_20c_clase_precision")

    # 21. Carga de precisión
    st.markdown("### ⚙️ 21. Carga de precisión")
    rango_burden = st.selectbox("Rango de burden acorde con IEC 61869-1/3/5", ["I", "II", "III", "IV"], key="param_21_rango_burden")
    st.text("a) Devanado 1: 15 VA")
    st.text("b) Devanado 2: 15 VA")
    st.text("c) Devanado 3: 15 VA")
    st.text("d) Simultánea: 45 VA")
    potencia_termica_limite = st.text_input("e) Potencia térmica límite (dejar vacío por ahora)", key="param_21_potencia_termica_limite")


    # 22. Tensión asignada
    st.markdown("### ⚡ 22. Tensión asignada")
    upn_opciones = [110, 230, 500]
    upn_seleccionada = st.selectbox("a) Tensión primaria (Upn)", upn_opciones)
    upn_calculada = round(upn_seleccionada / math.sqrt(3), 2)
    st.text(f"{upn_seleccionada} V dividido entre √3 ≈ {upn_calculada} V")

    usn_opciones = {
        "115 / √3": round(115 / math.sqrt(3), 2),
        "110 / √3": round(110 / math.sqrt(3), 2)
    }
    usn_seleccionada = st.selectbox("b) Tensión secundaria (Usn)", list(usn_opciones.keys()))
    st.text(f"{usn_seleccionada} ≈ {usn_opciones[usn_seleccionada]} V")


    # BOTÓN PARA GENERAR FICHA
   # 📋 Diccionario con los datos del transformador
    ficha_ctg = {
        "Fabricante": fabricante,
        "País": pais,
        "Referencia": referencia,
        "Norma de fabricación": norma_fabricacion,
        "Norma de calidad": norma_calidad,
        "Tipo de ejecución": tipo_ejecucion,
        "Altura de instalación (msnm)": altura_instalacion,
        "Material del aislador": material_aislador,
        "Tipo de transformador": tipo_transformador,
        "Tensión más elevada para el material (Um)": tension_um,
        "Tensión Ud - Aislamiento Interno": ud_interno,
        "Tensión Ud - Aislamiento Externo": f"{ud_interno} a {int(altura_instalacion)} msnm",
        "Tensión Up - Aislamiento Interno": up_interno,
        "Tensión Up - Aislamiento Externo": f"{up_interno} a {int(altura_instalacion)} msnm",
        "Tensión Us - Aislamiento Interno": us_interno,
        "Tensión Us - Aislamiento Externo": us_externo,
        "Frecuencia asignada (fr)": "60 Hz",
        "Factor de tensión permanente": "1,2",
        "Factor de tensión durante 30 s": "1,5",
        "Capacidad total (VA)": capacidad_total,
        "Condensador de alta tensión (C1)": c1,
        "Condensador de tensión intermedia (C2)": c2,
        "Tensión intermedia en circuito abierto": tension_intermedia,
        "Número de devanados secundarios": num_devanados,
        "Clase de precisión (5%-80%)": clase_precision_a,
        "Clase de precisión (80%-120%)": clase_precision_b,
        "Clase de precisión (120%-150%)": clase_precision_c,
        "Rango de burden (IEC 61869)": rango_burden,
        "Carga Devanado 1 (VA)": "15",
        "Carga Devanado 2 (VA)": "15",
        "Carga Devanado 3 (VA)": "15",
        "Carga Simultánea (VA)": "45",
        "Potencia térmica límite": potencia_termica_limite,
        "Tensión primaria (Upn)": f"{upn_seleccionada} V / √3 ≈ {upn_calculada} V",
        "Tensión secundaria (Usn)": f"{usn_seleccionada} ≈ {usn_opciones[usn_seleccionada]} V"
    }
    
    def exportar_excel(datos, fuente="Calibri", tamaño=9):
        unidades = {
            "Altura de instalación (msnm)": "msnm",
            "Capacidad total (VA)": "VA",
            "Tensión más elevada para el material (Um)": "kV",
            "Tensión Ud - Aislamiento Interno": "kV",
            "Tensión Ud - Aislamiento Externo": "kV",
            "Tensión Up - Aislamiento Interno": "kV",
            "Tensión Up - Aislamiento Externo": "kV",
            "Tensión Us - Aislamiento Interno": "kV",
            "Tensión Us - Aislamiento Externo": "kV",
            "Frecuencia asignada (fr)": "Hz",
            "Factor de tensión permanente": "",
            "Factor de tensión durante 30 s": "",
            "Carga Devanado 1 (VA)": "VA",
            "Carga Devanado 2 (VA)": "VA",
            "Carga Devanado 3 (VA)": "VA",
            "Carga Simultánea (VA)": "VA",
            "Tensión primaria (Upn)": "V",
            "Tensión secundaria (Usn)": "V"
            # Puedes añadir más unidades si lo deseas
        }
    
        df = pd.DataFrame([
            {
                "ÍTEM": i + 1,
                "DESCRIPCIÓN": campo,
                "UNIDAD": unidades.get(campo, ""),
                "REQUERIDO": valor,
                "OFRECIDO": ""
            }
            for i, (campo, valor) in enumerate(datos.items())
        ])
    
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="CTG", startrow=6)
            wb = writer.book
            ws = writer.sheets["CTG"]
    
            # 🖼️ Logo (opcional)
            logo_path = "siemens_logo.png"
            try:
                img = Image(logo_path)
                img.width = 300
                img.height = 100
                ws.add_image(img, "C1")
            except FileNotFoundError:
                st.warning("⚠️ No se encontró el logo 'siemens_logo.png'. Asegúrate de subirlo al repositorio.")
    
            # 🟪 Título
            ws.merge_cells("A2:E4")
            cell = ws.cell(row=2, column=1)
            cell.value = "FICHA TÉCNICA TRANSFORMADOR DE TENSIÓN"
            cell.font = Font(name=fuente, bold=True, size=14, color="000000")
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
            # 🏷️ Subtítulo
            ws.merge_cells("A5:D5")
            ws["A5"] = "CARACTERÍSTICAS GARANTIZADAS"
            ws["A5"].font = Font(name=fuente, bold=True, size=12)
            ws["A5"].alignment = Alignment(horizontal="center")
    
            # 🎨 Encabezados
            header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
            header_font = Font(name=fuente, size=tamaño, color="FFFFFF", bold=True)
            thin_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
    
            for col_num in range(1, 6):
                cell = ws.cell(row=6, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border
            # 📐 Ajuste de columnas
            ws.column_dimensions["A"].width = 4
            ws.column_dimensions["B"].width = 50
            ws.column_dimensions["C"].width = 10
            ws.column_dimensions["D"].width = 12
            ws.column_dimensions["E"].width = 12
    
            # 📋 Formato de filas
            for row in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=5):
                max_lines = 1
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.font = Font(name=fuente, size=tamaño)
    
                    if cell.value and isinstance(cell.value, str):
                        if cell.column_letter == "B":
                            wrapped = textwrap.wrap(cell.value, width=55)
                            max_lines = max(max_lines, len(wrapped))
    
                ws.row_dimensions[row[0].row].height = max_lines * 15
                row[0].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                row[2].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                row[3].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                row[4].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
        output.seek(0)
        return output

    
    ficha_ctg = mostrar_app()
    
    if st.button("📊 Generar archivo CTG"):
        archivo_excel = exportar_excel(ficha_ctg)
        st.download_button(
            label="📥 Descargar archivo CTG en Excel",
            data=archivo_excel,
            file_name="CTG_Transformador_Tension.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )









