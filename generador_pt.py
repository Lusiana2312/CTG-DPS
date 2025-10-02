# empieza codigo
import streamlit as st
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime
import numpy as np

import streamlit as st
import math

def mostrar_app():
    st.set_page_config(page_title="Generador CTG - Transformador de Tensión", layout="wide")

    st.title("📄 Generador de Ficha CTG")
    st.subheader("Transformador de Tensión")

    st.markdown("### ⚙️ Parámetros del transformador")

    # 1 al 5
    fabricante = st.text_input("1. Fabricante")
    pais = st.text_input("2. País")
    referencia = st.text_input("3. Referencia")
    norma_fabricacion = st.text_input("4. Norma de fabricación", value="IEC 61869-5")
    norma_calidad = st.text_input("5. Norma de calidad", value="ISO 9001")

    # 6 al 9
    tipo_ejecucion = st.selectbox("6. Tipo de ejecución", ["Interior", "Exterior"])
    altura_instalacion = st.number_input("7. Altura de instalación (msnm)", min_value=0, step=100)
    material_aislador = st.selectbox("8. Material del aislador", ["Compuesto siliconado", "Porcelana"])
    tipo_transformador = st.selectbox("8a. Tipo", ["Capacitivo", "Inductivo"])
    tension_um = st.selectbox("9. Tensión más elevada para el material (Um)", ["145 kV", "245 kV", "550 kV"])

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
    us_interno = st.text_input("Aislamiento Interno (dejar vacío por ahora)")
    us_externo = st.text_input("Aislamiento Externo (*) (dejar vacío por ahora)")

    # 13. Frecuencia
    st.markdown("### 📶 13. Frecuencia asignada (fr)")
    st.text("60 Hz")

    # 14. Factor de tensión
    st.markdown("### ⚙️ 14. Factor de tensión asignado")
    st.text("a) Permanente: 1,2")
    st.text("b) Durante 30 s: 1,5")

    # 15. Capacidad total
    st.markdown("### ⚡ 15. Capacidad total")
    capacidad_total = st.number_input("Capacidad total (≥ 4000 VA)", min_value=4000)

    # 16 al 18
    st.markdown("### 🔧 16-18. Condensadores y tensión intermedia")
    c1 = st.text_input("16. Condensador de alta tensión (C1)")
    c2 = st.text_input("17. Condensador de tensión intermedia (C2)")
    tension_intermedia = st.text_input("18. Tensión intermedia asignada en circuito abierto")

    # 19. Número de devanados secundarios
    st.markdown("### 🔁 19. Número de devanados secundarios")
    num_devanados = st.selectbox("Selecciona el número de devanados secundarios", [1, 2, 3])

    # 20. Clase de precisión
    st.markdown("### 🎯 20. Clase de precisión")
    st.markdown("**Entre el 25% y el 100% de la carga de precisión con factor de potencia 0,8 en atraso**")
    clase_precision_a = st.selectbox("a) Entre el 5% y el 80% de la tensión asignada", ["1P", "2P", "3P", "4P", "5P"])
    clase_precision_b = st.selectbox("b) Entre el 80% y el 120% de la tensión asignada", ["0.1", "0.2", "0.3"])
    clase_precision_c = st.selectbox("c) Entre el 120% y el 150% de la tensión asignada", ["1P", "2P", "3P", "4P", "5P"])

    # 21. Carga de precisión
    st.markdown("### ⚙️ 21. Carga de precisión")
    rango_burden = st.selectbox("Rango de burden acorde con IEC 61869-1/3/5", ["I", "II", "III", "IV"])
    st.text("a) Devanado 1: 15 VA")
    st.text("b) Devanado 2: 15 VA")
    st.text("c) Devanado 3: 15 VA")
    st.text("d) Simultánea: 45 VA")
    potencia_termica_limite = st.text_input("e) Potencia térmica límite (dejar vacío por ahora)")

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
    if st.button("Generar ficha CTG"):
        ficha_ctg = {
            # Datos manuales
            "Responsable": responsable,
            "Fecha de elaboración": fecha_elaboracion.strftime("%Y-%m-%d"),
            "Área técnica": area_tecnica,
            "Proyecto": proyecto,

            # Datos fijos
            "Tipo de equipo": "Transformador de Tensión",
            "Frecuencia asignada (fr)": frecuencia,
            "Estado": "Operativo",
            "Fecha de registro": datetime.now().strftime("%Y-%m-%d"),

            # Parámetros eléctricos
            "Tensión primaria (kV)": tension_primaria,
            "Tensión secundaria (V)": tension_secundaria,
            "Tensión de aislamiento (kV)": tension_aislamiento,
            "Tensión de impulso (kV)": tension_impulso,
            "Factor de tensión asignado - Permanente": factor_permanente,
            "Factor de tensión asignado - Durante 30s": factor_30s,
            "Tipo de conexión": tipo_conexion,
            "Tipo de aislamiento": tipo_aislamiento,

            # Parámetros adicionales
            "Capacidad total (pF)": capacidad_total,
            "Condensador de alta tensión (C1) (pF)": c1,
            "Condensador de tensión intermedia (C2) (pF)": c2,
            "Tensión intermedia asignada en circuito abierto (kV)": tension_intermedia,
            "Número de devanados secundarios": num_devanados
        }

        # Crear Excel en memoria
        wb = Workbook()
        ws = wb.active
        ws.title = "Ficha CTG"
        ws.append(["Parámetro", "Valor"])
        for parametro, valor in ficha_ctg.items():
            ws.append([parametro, valor])

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Ficha CTG generada correctamente.")
        st.download_button(
            label="📥 Descargar Excel",
            data=output,
            file_name="CTG_TransformadorTension.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


