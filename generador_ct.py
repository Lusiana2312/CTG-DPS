# empieza codigo
import streamlit as st
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime
import numpy as np

import streamlit as st

def mostrar_app():
    st.set_page_config(page_title="Generador CTG - Transformador de Corriente", layout="wide")

    st.title("📄 Generador de Ficha CTG")
    st.subheader("Transformador de Corriente")

    # 1. DATOS GENERALES
    st.markdown("### 🖊️ Datos generales")
    fabricante = st.text_input("Fabricante")
    pais = st.text_input("País")
    referencia = st.text_input("Referencia")
    norma_fabricacion = "IEC 61869-1 y IEC 61869-2"
    norma_calidad = "ISO 9001"
    st.markdown(f"**Norma de fabricación:** {norma_fabricacion}")
    st.markdown(f"**Norma de calidad:** {norma_calidad}")

    tipo_ejecucion = st.selectbox("Tipo de ejecución", options=["Exterior", "Interior"])
    altura_instalacion = st.number_input("Altura de instalación [msnm]", min_value=0)
    material_aislador = st.selectbox("Material del aislador", options=["Compuesto Siliconado", "Material"])

    # 2. PARÁMETROS DE TENSIÓN
    st.markdown("### ⚡ Parámetros de tensión")
    tension_material = st.selectbox("Tensión más elevada para el material (Um)", options=["145 kV", "245 kV", "550 kV"])

    # Asignación automática de tensiones
    if tension_material == "145 kV":
        ud_interno = "360 kV"
        up_interno = "750 kV"
        ipn = "1000 A"
    elif tension_material == "245 kV":
        ud_interno = "460 kV"
        up_interno = "1050 kV"
        ipn = "2500 A"
    elif tension_material == "550 kV":
        ud_interno = "700 kV"
        up_interno = "1550 kV"
        ipn = "3000 A"
    else:
        ud_interno = ""
        up_interno = ""
        ipn = ""

    ud_externo = f"{ud_interno} a {altura_instalacion} msnm" if ud_interno else ""
    up_externo = f"{up_interno} a {altura_instalacion} msnm" if up_interno else ""

    st.write("**Tensión asignada soportada a la frecuencia industrial (Ud) - Aislamiento Interno:**", ud_interno)
    st.write("**Tensión asignada soportada a la frecuencia industrial (Ud) - Aislamiento Externo:**", ud_externo)
    st.write("**Tensión asignada soportada al impulso tipo rayo (Up) - Aislamiento Interno:**", up_interno)
    st.write("**Tensión asignada soportada al impulso tipo rayo (Up) - Aislamiento Externo:**", up_externo)

    us_interno = st.text_input("Tensión asignada soportada al impulso tipo maniobra (Us) - Aislamiento Interno")
    us_externo = st.text_input("Tensión asignada soportada al impulso tipo maniobra (Us) - Aislamiento Externo")

    # 3. PARÁMETROS ELÉCTRICOS
    st.markdown("### 🔌 Parámetros eléctricos")
    frecuencia = "60 Hz"
    st.markdown(f"**Frecuencia asignada (fr):** {frecuencia}")
    st.markdown(f"**Corriente primaria asignada (Ipn):** {ipn}")
    factor_ipn = "1"
    st.markdown(f"**Factor de la corriente primaria continua asignada:** {factor_ipn}")
    isn = "1"
    st.markdown(f"**Corriente secundaria asignada (Isn):** {isn}")
    ith = "40 kA"
    st.markdown(f"**Corriente de cortocircuito térmica asignada (Ith) en 1 segundo:** {ith}")
    idyn = "2.6 × Ith"
    st.markdown(f"**Corriente dinámica asignada (Idyn):** {idyn}")

    # 4. CANTIDAD Y CLASE DE NÚCLEOS
    st.markdown("### 🧲 Cantidad y clase de núcleos")
    cantidad_nucleos = st.number_input("Cantidad total de núcleos", min_value=1, max_value=6, value=6)

    tipos_nucleo = {}
    for i in range(1, cantidad_nucleos + 1):
        tipo = st.selectbox(
            f"Tipo de núcleo {i}",
            options=["Medida", "Protección convencional"],
            key=f"tipo_nucleo_{i}"
        )
        tipos_nucleo[f"Núcleo {i}"] = tipo

    st.markdown("#### 🧾 Resumen de núcleos")
    for nucleo, tipo in tipos_nucleo.items():
        st.write(f"{nucleo}: {tipo}")

    # 5. DESCRIPCIÓN DE NÚCLEOS
        st.markdown("### 📘 Descripción de núcleos")
    
        opciones_relacion_transformacion = [
            "2500-1250-625/1",
            "2500-1250-600-2"
        ]
    
        clases_exactitud = ["1P", "2P", "3P", "4P", "5P"]
        factores_precision = ["50", "60", "70"]
    
        for i in range(1, cantidad_nucleos + 1):
            tipo = tipos_nucleo[f"Núcleo {i}"]
            st.markdown(f"#### 🔹 Núcleo {i} ({tipo})")
    
            if tipo == "Medida":
                relacion_asignada = st.selectbox(
                    f"a) Relación de transformación asignada - Núcleo {i}",
                    options=opciones_relacion_transformacion,
                    key=f"relacion_asignada_{i}"
                )
    
                relacion_exactitud = st.selectbox(
                    f"b) Relación para la que debe cumplir la exactitud - Núcleo {i}",
                    options=opciones_relacion_transformacion,
                    key=f"relacion_exactitud_{i}"
                )
    
                clase_exactitud = st.selectbox(
                    f"c) Clase de exactitud - Núcleo {i}",
                    options=clases_exactitud,
                    key=f"clase_exactitud_{i}"
                )
    
                if i == 6:
                    factor_precision = st.selectbox(
                        f"d) Factor límite de precisión - Núcleo {i}",
                        options=factores_precision,
                        key=f"factor_precision_{i}"
                    )
    
                    st.markdown("e) Carga de exactitud - Núcleo 6:")
                    st.write("• 625/1 (1S3-1S4): N.A")
                    st.write("• 1250/1 (1S2-1S4): N.A")
                    st.write("• 2500/1 (1S1-1S4): N.A")
                    st.write("• 400/1 (1S3-1S4): N.A")
                    st.write("• 800/1 (1S2-1S4): N.A")
                    st.write("• 1600/1 (1S1-1S4): N.A")
                else:
                    carga_exactitud = st.text_input(
                        f"e) Carga de exactitud (VA) - Núcleo {i}",
                        key=f"carga_exactitud_{i}"
                    )
    
            elif tipo == "Protección convencional":
                st.markdown("*Este núcleo está clasificado como protección convencional. Puedes definir sus parámetros más adelante.*")
    




    

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

