# empieza codigo
import streamlit as st
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime


def mostrar_app():
    st.set_page_config(page_title="Generador CTG - Interruptor de Potencia", layout="wide")

    st.title("📄 Generador de Ficha CTG")
    st.subheader("Interruptor de Potencia")

    # 1. DATOS GENERALES
    st.markdown("### 🖊️ Datos generales")
    fabricante = st.text_input("Fabricante")
    pais = st.text_input("País")
    referencia = st.text_input("Referencia")
    norma_fabricacion = st.text_input("Norma de fabricación")
    norma_calidad = st.text_input("Norma de calidad")

    # 2. CARACTERÍSTICAS TÉCNICAS
    st.markdown("### ⚙️ Características técnicas")
    medio_extincion = st.selectbox("Medio de extinción", ["Vacío", "SF6", "Aceite", "Aire comprimido"])
    num_polos = st.selectbox("Número de polos", [1, 2, 3])
    camaras_por_polo = st.selectbox("Número de cámaras por polo", [1, 2])
    tipo_ejecucion = st.selectbox("Tipo de ejecución", ["Exterior", "Interior"])
    altura_instalacion = st.number_input("Altura de instalación (m.s.n.m)", min_value=0, value=1000)

    # 3. TEMPERATURA DE OPERACIÓN
    st.markdown("### 🌡️ Temperatura de operación")
    temp_min = st.number_input("a) Temperatura mínima anual (°C)", value=-5)
    temp_max = st.number_input("b) Temperatura máxima anual (°C)", value=40)
    temp_media = st.number_input("c) Temperatura media (24 h) (°C)", value=25)

    # BOTÓN PARA GENERAR FICHA
    if st.button("Generar ficha CTG"):
        ficha_cb = {
            "Fabricante": fabricante,
            "País": pais,
            "Referencia": referencia,
            "Norma de fabricación": norma_fabricacion,
            "Norma de calidad": norma_calidad,
            "Medio de extinción": medio_extincion,
            "Número de polos": num_polos,
            "Número de cámaras por polo": camaras_por_polo,
            "Tipo de ejecución": tipo_ejecucion,
            "Altura de instalación (m.s.n.m)": altura_instalacion,
            "Temperatura mínima anual (°C)": temp_min,
            "Temperatura máxima anual (°C)": temp_max,
            "Temperatura media (24 h) (°C)": temp_media,
            "Fecha de registro": datetime.now().strftime("%Y-%m-%d")
        }

        # Crear Excel en memoria
        wb = Workbook()
        ws = wb.active
        ws.title = "Ficha CTG"
        ws.append(["Parámetro", "Valor"])
        for parametro, valor in ficha_cb.items():
            ws.append([parametro, valor])

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Ficha CTG generada correctamente.")
        st.download_button(
            label="📥 Descargar Excel",
            data=output,
            file_name="CTG_InterruptorPotencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
