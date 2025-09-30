# empieza codigo
import streamlit as st
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime
import numpy as np

def mostrar_app():
    st.set_page_config(page_title="Generador CTG - Transformador de Corriente", layout="wide")

    st.title("📄 Generador de Ficha CTG")
    st.subheader("Transformador de Corriente")

    # 1. DATOS PARA ANOTACIÓN MANUAL
    st.markdown("### 🖊️ Datos para anotación manual")
    responsable = st.text_input("Nombre del responsable")
    fecha_elaboracion = st.date_input("Fecha de elaboración")
    area_tecnica = st.text_input("Área técnica")
    proyecto = st.text_input("Proyecto o ubicación general")

    # 2. PARÁMETROS DE TENSIÓN
    st.markdown("### ⚡ Parámetros de tensión")
    tension_material = st.selectbox(
        "Tensión más elevada para el material (Um)",
        options=["115 kV", "245 kV", "550 kV"]
    )

    # Asignación automática de tensiones e Ith según Um
    if tension_material == "115 kV":
        tension_nominal = "110 kV"
        tension_ensayo = "195 kV"
        tension_impulso = "250 kV"
        ith_sugerido = "35"
    elif tension_material == "245 kV":
        tension_nominal = "230 kV"
        tension_ensayo = "395 kV"
        tension_impulso = "460 kV"
        ith_sugerido = "40"
    elif tension_material == "550 kV":
        tension_nominal = "525 kV"
        tension_ensayo = "610 kV"
        tension_impulso = "1000 kV"
        ith_sugerido = "63"
    else:
        tension_nominal = ""
        tension_ensayo = ""
        tension_impulso = ""
        ith_sugerido = ""

    st.write("**Tensión nominal asignada:**", tension_nominal)
    st.write("**Tensión de ensayo asignada:**", tension_ensayo)
    st.write("**Tensión de impulso asignada:**", tension_impulso)

    # 3. PARÁMETROS ELÉCTRICOS
    st.markdown("### 🔌 Parámetros eléctricos")
    st.markdown("**Frecuencia asignada (fr):** 60 Hz  \n*Valor fijo para sistemas eléctricos en Colombia*")
    frecuencia = "60 Hz"
    st.markdown(f"**Corriente de cortocircuito térmica asignada (Ith):** {ith_sugerido} kA  \n*Asignada automáticamente según Um*")
    ith = f"{ith_sugerido} kA"

    valores_entre_1_y_2 = [f"{v:.1f}" for v in np.arange(1.0, 2.1, 0.1)]
    ipn = st.text_input("Corriente primaria asignada (Ipn) [A]")
    factor_ipn = st.selectbox("Factor de corriente primaria continua asignada", options=valores_entre_1_y_2)
    isn = st.selectbox("Corriente secundaria asignada (Isn) [A]", options=valores_entre_1_y_2)
    idyn = st.text_input("Corriente dinámica asignada (Idyn) [kA]")

    # 4. PARÁMETROS DE NÚCLEOS
    st.markdown("### 🧲 Parámetros de núcleos")
    num_nucleos = st.selectbox("Número de núcleos", options=[1, 2, 3, 4, 5, 6])
    relaciones_fijas = ["625/1", "800/1", "1000/1", "1200/1", "1400/1", "1600/1"]
    cargas_por_nucleo = {}

    for i in range(1, num_nucleos + 1):
        st.markdown(f"#### Núcleo {i}")
        cargas_por_nucleo[f"Núcleo {i}"] = {}
        for relacion in relaciones_fijas:
            carga_va = st.text_input(
                f"Carga (VA) para relación {relacion} en núcleo {i}",
                key=f"carga_{i}_{relacion}"
            )
            cargas_por_nucleo[f"Núcleo {i}"][relacion] = carga_va

    # BOTÓN PARA GENERAR FICHA
    if st.button("Generar ficha CTG"):
        datos_manual = {
            "Responsable": responsable,
            "Fecha de elaboración": fecha_elaboracion.strftime("%Y-%m-%d"),
            "Área técnica": area_tecnica,
            "Proyecto": proyecto
        }

        datos_fijos = {
            "Tipo de equipo": "Transformador de Corriente",
            "Frecuencia asignada (fr)": frecuencia,
            "Frecuencia nominal": "60 Hz",
            "Clase de precisión": "0.5",
            "Estado": "Operativo",
            "Fecha de registro": datetime.now().strftime("%Y-%m-%d")
        }

        datos_tension = {
            "Tensión más elevada para el material (Um)": tension_material,
            "Tensión nominal asignada": tension_nominal,
            "Tensión de ensayo asignada": tension_ensayo,
            "Tensión de impulso asignada": tension_impulso
        }

        datos_electricos = {
            "Corriente primaria asignada (Ipn)": ipn,
            "Factor de corriente primaria continua asignada": factor_ipn,
            "Corriente secundaria asignada (Isn)": isn,
            "Corriente de cortocircuito térmica asignada (Ith)": ith,
            "Corriente dinámica asignada (Idyn)": idyn,
            "Número de núcleos": num_nucleos
        }

        ficha_ctg = {**datos_manual, **datos_fijos, **datos_tension, **datos_electricos}

        for nucleo, relaciones in cargas_por_nucleo.items():
            for relacion, carga in relaciones.items():
                ficha_ctg[f"{nucleo} - Relación {relacion} - Carga (VA)"] = carga

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
            file_name="CTG_TransformadorCorriente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

