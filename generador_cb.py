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
    norma_fabricacion = "IEC 62271-100 / IEC 62271-110"
    st.markdown(f"**Norma de fabricación:** {norma_fabricacion}")
    norma_calidad = "ISO 9001"
    st.markdown(f"**Norma de calidad:** {norma_calidad}")

    # 2. CARACTERÍSTICAS TÉCNICAS
    st.markdown("### ⚙️ Características técnicas")
    medio_extincion = st.selectbox("Medio de extinción", ["Vacío", "SF6", "Aceite", "Aire comprimido"])
    num_polos = st.selectbox("Número de polos", [1, 2, 3, 4])
    camaras_por_polo = st.text_input("Número de cámaras por polo")
    tipo_ejecucion = st.selectbox("Tipo de ejecución", ["Exterior", "Interior"])
    altura_instalacion = st.number_input("Altura de instalación (m.s.n.m)", min_value=0, value=1000)

    # 3. TEMPERATURA DE OPERACIÓN
    st.markdown("### 🌡️ Temperatura de operación")
    temp_min = st.number_input("a) Temperatura mínima anual (°C)", value=-5)
    temp_max = st.number_input("b) Temperatura máxima anual (°C)", value=40)
    temp_media = st.number_input("c) Temperatura media (24 h) (°C)", value=25)
    
    # 4. PARÁMETROS AMBIENTALES Y ELÉCTRICOS ADICIONALES
    st.markdown("### 🌍 Parámetros ambientales y eléctricos adicionales")

    categoria_corrosion = st.selectbox(
        "Categoría de corrosión del ambiente (ISO 12944-2 / ISO 9223)",
        options=["C1 - Muy baja", "C2 - Baja", "C3 - Media", "C4 - Alta", "C5 - Muy alta", "CX - Extrema"]
    )

    frecuencia_asignada = st.selectbox("Frecuencia asignada (fr)", options=["50 Hz", "60 Hz"])
    ur = st.selectbox("Tensión asignada (Ur)", options=["145 kV", "245 kV", "550 kV"])
    
    # Asignación automática de Ud según Ur
    ud_por_ur = {
        "145 kV": "275 kV",
        "245 kV": "640 kV",
        "550 kV": "830 kV"
    }
    ud_frecuencia = ud_por_ur.get(ur, "")
    st.markdown(f"**Tensión asignada soportada a frecuencia industrial (Ud):** {ud_frecuencia}")

    # Asignación automática de Us por componente según Ur
    us_por_ur = {
        "145 kV": {"fase_tierra": "N.A.", "entre_fases": "N.A.", "interruptor_abierto": "N.A."},
        "245 kV": {"fase_tierra": "N.A.", "entre_fases": "N.A.", "interruptor_abierto": "N.A."},
        "550 kV": {"fase_tierra": "1175 kV", "entre_fases": "1175 kV", "interruptor_abierto": "1175 kV"}
    }
    us_valores = us_por_ur.get(ur, {"fase_tierra": "", "entre_fases": "", "interruptor_abierto": ""})
    st.markdown("#### Tensión asignada soportada a impulso de maniobra (Us)")
    st.markdown(f"a) Fase-Tierra: **{us_valores['fase_tierra']}**")
    st.markdown(f"b) Entre fases: **{us_valores['entre_fases']}**")
    st.markdown(f"c) A través de interruptor abierto: **{us_valores['interruptor_abierto']}**")

    # Asignación automática de Up según Ur
    up_por_ur = {
        "145 kV": "650 kV",
        "245 kV": "1050 kV",
        "550 kV": "1800 kV"
    }
    up_rayo = up_por_ur.get(ur, "")
    st.markdown(f"**Tensión asignada soportada al impulso tipo rayo (Up):** {up_rayo}")
    # Opciones de corriente asignada según Ur
    ir_por_ur = {
        "145 kV": ["1200 A", "2000 A", "3150 A"],
        "245 kV": ["1200 A", "2000 A", "2500 A", "3000 A", "4000 A"],
        "550 kV": ["3000 A", "4000 A", "5000 A", "6300 A"]
    }
    # Mostrar opciones de Ir según Ur
    opciones_ir = ir_por_ur.get(ur, [])
    ir = st.selectbox("Corriente asignada en servicio continuo (Ir)", opciones_ir)
    
    # Opciones de poder de corte asignado (Ics) según Ur
    ics_por_ur = {
        "145 kV": ["25 kA", "31.5 kA", "40 kA"],
        "245 kV": ["40 kA", "50 kA", "63 kA"],
        "550 kV": ["50 kA", "63 kA"]
    }
    # Mostrar opciones de Ics según Ur
    opciones_ics = ics_por_ur.get(ur, [])
    ics = st.selectbox("Poder de corte asignado en cortocircuito (Ics)", opciones_ics)

    # Duración del cortocircuito asignado (Ics)
    duracion_ics = st.selectbox("Duración del cortocircuito asignado (Ics)", ["1 s", "2 s", "3 s"])

    # Porcentaje de corriente aperiódica (%)
    porcentaje_ap = st.text_input("Porcentaje de corriente aperiódica (%)")

    # Poder de cierre asignado en cortocircuito (Ip)
    st.markdown("**Poder de cierre asignado en cortocircuito (Ip):** 2.6 × Ics")

    # Factor de primer polo según Ur
    factor_primer_polo_por_ur = {
        "145 kV": "1.3",
        "245 kV": "1.5",
        "550 kV": "1.5"
    }
    factor_primer_polo = factor_primer_polo_por_ur.get(ur, "")
    st.markdown(f"**Factor de primer polo:** {factor_primer_polo}")

    # Tensión transitoria de restablecimiento asignada para fallas en bornes
    st.markdown("### ⚡ Tensión transitoria de restablecimiento asignada para fallas en bornes")

    u1 = st.text_input("a) Primera tensión de referencia (u1) kV")
    t1 = st.text_input("b) Tiempo t1 ms")
    uc = st.text_input("c) Valor cresta del TTR (uc) kV")
    t2 = st.text_input("d) Tiempo t2 ms")
    td = st.text_input("e) Retardo td ms")
    u_prima = st.text_input("f) Tensión u’ kV")
    t_prima = st.text_input("g) Tiempo t’ ms")
    vel_crecimiento = st.text_input("h) Velocidad de crecimiento (u1 / t1) kV/ms")

    # Características asignadas para fallas próximas en líneas
    st.markdown("### ⚡ Características asignadas para fallas próximas en líneas")

    # a) Características asignadas del circuito de alimentación
    st.markdown("#### a) Características asignadas del circuito de alimentación")
    u1_alimentacion = st.text_input("• Primera tensión de referencia (u1) [alimentación]")
    t1_alimentacion = st.text_input("• Tiempo t1 [alimentación]")
    uc_alimentacion = st.text_input("• Valor cresta del TTR (uc) [alimentación]")
    t2_alimentacion = st.text_input("• Tiempo t2 [alimentación]")
    td_alimentacion = st.text_input("• Retardo td [alimentación]")
    u_prima_alimentacion = st.text_input("• Tensión u’ [alimentación]")
    t_prima_alimentacion = st.text_input("• Tiempo t’ [alimentación]")
    vel_crecimiento_alimentacion = st.text_input("• Velocidad de crecimiento (u1 / t1) [alimentación]")

    # b) Características asignadas de la línea
    st.markdown("#### b) Características asignadas de la línea")
    z_linea = st.text_input("• Impedancia de onda asignada (Z)")
    k_linea = st.text_input("• Factor de cresta asignada (k)")
    s_linea = st.text_input("• Factor de TCTR (s)")
    tdl_linea = st.text_input("• Retardo (tdl)")
    
    # Característica de TRV de pequeñas corrientes inductivas según IEC 62271-110
    st.markdown("### ⚡ Característica de TRV de pequeñas corrientes inductivas según IEC 62271-110")
    # Rangos de referencia según Ur
    rangos_trv = {
        "145 kV": {
            "Uc": "250–300 kV",
            "t3_1": "100–150 µs",
            "t3_2": "200–250 µs"
        },
        "245 kV": {
            "Uc": "350–400 kV",
            "t3_1": "150–200 µs",
            "t3_2": "250–300 µs"
        },
        "550 kV": {
            "Uc": "600–700 kV",
            "t3_1": "200–300 µs",
            "t3_2": "300–400 µs"
        }
    }
    valores_trv = rangos_trv.get(ur, {"Uc": "", "t3_1": "", "t3_2": ""})
    trv_uc_min = st.text_input(f"a) Valor mínimo pico de TRV Uc (Ej: {valores_trv['Uc']})")
    trv_t3_circuito1 = st.text_input(f"b) Tiempo máximo t₃ Load circuit 1 (Ej: {valores_trv['t3_1']})")
    trv_t3_circuito2 = st.text_input(f"c) Tiempo máximo t₃ Load circuit 2 (Ej: {valores_trv['t3_2']})")

    # Tiempo de arco mínimo ante pequeñas corrientes inductivas
    st.markdown("### ⏱️ Tiempo de arco mínimo ante pequeñas corrientes inductivas")
    tiempo_arco_minimo = st.text_input("Tiempo de arco mínimo (Minimum Arcing Time)")


    # Número de corte λ ("Chopping Number λ")
    st.markdown("### 🔢 Número de corte λ (Chopping Number λ)")
    numero_corte_lambda = st.text_input("Número de corte λ (Chopping Number λ)")

    # Secuencia de maniobras asignada
    st.markdown("### 🔁 Secuencia de maniobras asignada")
    secuencia_maniobras = st.text_input("Secuencia de maniobras asignada")

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
            "Fecha de registro": datetime.now().strftime("%Y-%m-%d"),
            "Categoría de corrosión del ambiente": categoria_corrosion,
            "Frecuencia asignada (fr)": frecuencia_asignada,
            "Tensión asignada (Ur) [kV]": ur,
            "Tensión asignada soportada a frecuencia industrial (Ud)": ud_frecuencia,
            "Us - Fase-Tierra [kV]": us_valores["fase_tierra"],
            "Us - Entre fases [kV]": us_valores["entre_fases"],
            "Us - A través de interruptor abierto [kV]": us_valores["interruptor_abierto"],
            "Tensión asignada soportada al impulso tipo rayo (Up)": up_rayo,
            "Corriente asignada en servicio continuo (Ir)": ir,
            "Poder de corte asignado en cortocircuito (Ics)": ics,
            "Duración del cortocircuito asignado (Ics)": duracion_ics,
            "Porcentaje de corriente aperiódica (%)": porcentaje_ap,
            "Poder de cierre asignado en cortocircuito (Ip)": "2.6 × Ics",
            "Factor de primer polo": factor_primer_polo,
            "TTR - Primera tensión de referencia (u1)": u1,
            "TTR - Tiempo t1": t1,
            "TTR - Valor cresta del TTR (uc)": uc,
            "TTR - Tiempo t2": t2,
            "TTR - Retardo td": td,
            "TTR - Tensión u’": u_prima,
            "TTR - Tiempo t’": t_prima,
            "TTR - Velocidad de crecimiento (u1 / t1)": vel_crecimiento,
            "Fallas próximas - u1 alimentación": u1_alimentacion,
            "Fallas próximas - t1 alimentación": t1_alimentacion,
            "Fallas próximas - uc alimentación": uc_alimentacion,
            "Fallas próximas - t2 alimentación": t2_alimentacion,
            "Fallas próximas - td alimentación": td_alimentacion,
            "Fallas próximas - u’ alimentación": u_prima_alimentacion,
            "Fallas próximas - t’ alimentación": t_prima_alimentacion,
            "Fallas próximas - velocidad crecimiento alimentación": vel_crecimiento_alimentacion,
            "Fallas próximas - impedancia de onda (Z)": z_linea,
            "Fallas próximas - factor de cresta (k)": k_linea,
            "Fallas próximas - factor de TCTR (s)": s_linea,
            "Fallas próximas - retardo (tdl)": tdl_linea,
            "TRV - Valor mínimo pico de TRV Uc": trv_uc_min,
            "TRV - Tiempo máximo t₃ Load circuit 1": trv_t3_circuito1,
            "TRV - Tiempo máximo t₃ Load circuit 2": trv_t3_circuito2,
            "Tiempo de arco mínimo (Minimum Arcing Time)": tiempo_arco_minimo,
            "Número de corte λ (Chopping Number λ)": numero_corte_lambda,
            "Secuencia de maniobras asignada": secuencia_maniobras

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
