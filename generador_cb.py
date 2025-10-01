# empieza codigo
import streamlit as st
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime
import pandas as pd
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import textwrap

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

    # Poder de corte en discordancia de fases (Id)
    st.markdown("### ⚡ Poder de corte en discordancia de fases (Id)")

    id_u1 = st.text_input("a) Primera tensión de referencia (u1) [Id]")
    id_t1 = st.text_input("b) Tiempo t1 [Id]")
    id_uc = st.text_input("c) Valor cresta del TTR (uc) [Id]")
    id_t2 = st.text_input("d) Tiempo t2 [Id]")
    id_vel_crecimiento = st.text_input("e) Velocidad de crecimiento (u1 / t1) [Id]")

    # Apertura de líneas en vacío
    st.markdown("### ⚡ Apertura de líneas en vacío")

    ir_apertura_linea = st.text_input("a) Poder de corte asignado (Ir) [Apertura de líneas en vacío]")
    sobretension_maniobra = st.text_input("b) Sobretensión de maniobra presente")

    # Apertura de corrientes inductivas pequeñas
    st.markdown("### ⚡ Apertura de corrientes inductivas pequeñas")

    apertura_inductiva = st.selectbox("¿Aplica apertura de corrientes inductivas pequeñas?", ["Sí", "No"])
    ir_inductiva = st.text_input("a) Poder de corte asignado [corrientes inductivas pequeñas]")
    sobretension_inductiva = st.text_input("b) Sobretensión de maniobra máxima")

    # Número de operaciones mecánicas
    st.markdown("### ⚙️ Número de operaciones mecánicas")
    num_operaciones_mecanicas = st.selectbox("Número de operaciones mecánicas", ["M1", "M2", "M3"])

    # Probabilidad de reencendido
    st.markdown("### 🔄 Probabilidad de reencendido")
    probabilidad_reencendido = st.selectbox("Probabilidad de reencendido", ["C1", "C2"])

    # Máxima diferencia de tiempo entre contactos de diferente polo
    st.markdown("### ⏱️ Máxima diferencia de tiempo entre contactos de diferente polo")
    diferencia_tiempo_contactos = st.text_input(
        "Máxima diferencia de tiempo entre contactos de diferente polo al tocarse durante un cierre o al separarse durante una apertura"
    )
    
    # Maniobra de apertura
    st.markdown("### 🔧 Maniobra de apertura")

    tiempo_apertura = st.text_input("a) Tiempo de apertura")
    tiempo_arco = st.text_input("b) Tiempo de arco")
    tiempo_max_corte = st.text_input("c) Tiempo máximo de corte asignado")

    # Tiempo muerto
    st.markdown("### ⏳ Tiempo muerto")
    tiempo_muerto = st.text_input("Tiempo muerto")

    # Maniobra de cierre
    st.markdown("### 🔧 Maniobra de cierre")

    tiempo_establecimiento = st.text_input("a) Tiempo de establecimiento")
    tiempo_prearco = st.text_input("b) Tiempo de prearco")
    tiempo_cierre = st.text_input("c) Tiempo de cierre")

    # Gas SF6 - Interruptor
    st.markdown("### 🧪 Gas SF₆ – Interruptor")

    presion_maniobra = st.text_input("a) Presión de gas asignada para maniobra (Pob)")
    presion_corte = st.text_input("b) Presión de gas asignada para el corte (Pcb)")

    # Volumen total de SF6 por polo a 0,1 MPa
    st.markdown("### 🧪 Volumen total de SF₆ por polo a 0,1 MPa")
    volumen_sf6 = st.text_input("Volumen total de SF₆ por polo a 0,1 MPa")

    # Pérdida máxima de SF6 por año (valor fijo)
    st.markdown("### 🧪 Pérdida máxima de SF₆ por año")
    perdida_sf6 = "≤ 0.5%"
    st.markdown(f"**Pérdida máxima de SF₆ por año:** {perdida_sf6}")

    
    # 🧪 Resistencia máxima entre terminales
    st.markdown("### 🧪 Resistencia máxima entre terminales")
    resistencia_max_terminales = st.text_input("Resistencia máxima entre terminales (μΩ)")

    # 🧪 Capacitancia
    st.markdown("### 🧪 Capacitancia")

    cap_entre_contactos_con_resistencia = st.text_input("a) Entre contactos abiertos - Con resistencia de preinserción (pF)")
    cap_entre_contactos_sin_resistencia = st.text_input("a) Entre contactos abiertos - Sin resistencia de preinserción (pF)")
    cap_entre_contactos_tierra = st.text_input("b) Entre contactos y tierra (pF)")
    cap_condensador_gradiente = st.text_input("c) Condensador de gradiente (***) (pF)")

    # 🧪 Material de los empaques
    st.markdown("### 🧪 Material de los empaques")
    material_empaques = st.text_input("Material de los empaques")

    # 🧪 Operación con mando sincronizado
    st.markdown("### 🧪 Operación con mando sincronizado")
    mando_sincronizado = st.radio("¿Operación con mando sincronizado?", ["Sí", "No"])

    # 🧪 Resistencia de preinserción
    st.markdown("### 🧪 Resistencia de preinserción")
    resistencia_preinsercion = st.radio("¿Resistencia de preinserción?", ["Sí", "No"])

    # 🧪 Distancia mínima en aire
    st.markdown("### 🧪 Distancia mínima en aire")

    distancia_entre_polos = st.text_input("a) Entre polos (mm)")
    distancia_a_tierra = st.text_input("b) A tierra (mm)")
    distancia_a_traves_polo = st.text_input("c) A través del polo (mm)")

    # 🧪 Clase de severidad de contaminación del sitio (SPS)
    st.markdown("### 🧪 Clase de severidad de contaminación del sitio (SPS)")
    sps_clase = st.selectbox(
        "Clase de severidad de contaminación del sitio (SPS)",
        ["Ligera", "Media", "Pesada", "Muy pesada"]
    )

    # 🧪 Distancia mínima de fuga
    st.markdown("### 🧪 Distancia mínima de fuga")
    distancia_minima_fuga = st.text_input("Distancia mínima de fuga (mm)")

    # 🧪 Datos sísmicos
    st.markdown("### 🧪 Datos sísmicos")
    desempeno_sismico_ieee = st.text_input("Desempeño sísmico según IEEE-693-Vigente (**)")
    frecuencia_natural_vibracion = st.text_input("a) Frecuencia natural de vibración (Hz)")
    coef_amortiguamiento_critico = st.text_input("b) Coeficiente de amortiguamiento crítico (%)")

    # 🧪 Cargas admisibles en bornes
    st.markdown("### 🧪 Cargas admisibles en bornes")
    carga_estatica_admisible = st.text_input("a) Carga estática admisible (N)")
    carga_dinamica_admisible = st.text_input("b) Carga dinámica admisible (N)")
    
    # 🧪 Fuerzas asociadas a la operación del equipo
    st.markdown("### 🧪 Fuerzas asociadas a la operación del equipo")

    fuerza_vertical = st.text_input("a) Fuerza vertical (N)")
    fuerza_horizontal = st.text_input("b) Fuerza horizontal (N)")

    # 🧪 Masa neta de un polo completo con estructura
    st.markdown("### 🧪 Masa neta de un polo completo con estructura")
    masa_neta_polo = st.text_input("Masa neta de un polo completo con estructura (kg)")

    # 🧪 Dimensiones para transporte
    st.markdown("### 🧪 Dimensiones para transporte")
    dimensiones_transporte = st.text_input("Dimensiones para transporte (Alto x Ancho x Largo) [mm]")

    # 🧪 Datos adicionales para transporte y campo eléctrico
    st.markdown("### 🧪 Datos adicionales")

    masa_neta_transporte = st.text_input("Masa neta para transporte (kg)")
    volumen_total_transporte = st.text_input("Volumen total para transporte (m³)")
    campo_electrico_1m = st.text_input("Campo eléctrico a 1 metro de separación del piso (kV/m)")

    # BOTÓN PARA GENERAR FICHA
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
        "Secuencia de maniobras asignada": secuencia_maniobras,
        "Poder de corte en discordancia de fases - u1": id_u1,
        "Poder de corte en discordancia de fases - t1": id_t1,
        "Poder de corte en discordancia de fases - uc": id_uc,
        "Poder de corte en discordancia de fases - t2": id_t2,
        "Poder de corte en discordancia de fases - velocidad de crecimiento (u1 / t1)": id_vel_crecimiento,
        "Apertura de líneas en vacío - Poder de corte asignado (Ir)": ir_apertura_linea,
        "Apertura de líneas en vacío - Sobretensión de maniobra presente": sobretension_maniobra,
        "Apertura de corrientes inductivas pequeñas": apertura_inductiva,
        "Apertura inductiva - Poder de corte asignado": ir_inductiva,
        "Apertura inductiva - Sobretensión de maniobra máxima": sobretension_inductiva,
        "Número de operaciones mecánicas": num_operaciones_mecanicas,
        "Probabilidad de reencendido": probabilidad_reencendido,
        "Máxima diferencia de tiempo entre contactos de diferente polo": diferencia_tiempo_contactos,
        "Maniobra de apertura - Tiempo de apertura": tiempo_apertura,
        "Maniobra de apertura - Tiempo de arco": tiempo_arco,
        "Maniobra de apertura - Tiempo máximo de corte asignado": tiempo_max_corte,
        "Tiempo muerto": tiempo_muerto,
        "Maniobra de cierre - Tiempo de establecimiento": tiempo_establecimiento,
        "Maniobra de cierre - Tiempo de prearco": tiempo_prearco,
        "Maniobra de cierre - Tiempo de cierre": tiempo_cierre,
        "Gas SF6 - Presión de maniobra (Pob)": presion_maniobra,
        "Gas SF6 - Presión de corte (Pcb)": presion_corte,
        "Volumen total de SF6 por polo a 0,1 MPa": volumen_sf6,
        "Pérdida máxima de SF6 por año": perdida_sf6,
        "Resistencia máxima entre terminales (μΩ)": resistencia_max_terminales,
        "Capacitancia - Entre contactos abiertos con resistencia de preinserción (pF)": cap_entre_contactos_con_resistencia,
        "Capacitancia - Entre contactos abiertos sin resistencia de preinserción (pF)": cap_entre_contactos_sin_resistencia,
        "Capacitancia - Entre contactos y tierra (pF)": cap_entre_contactos_tierra,
        "Capacitancia - Condensador de gradiente (***) (pF)": cap_condensador_gradiente,
        "Material de los empaques": material_empaques,
        "Operación con mando sincronizado": mando_sincronizado,
        "Resistencia de preinserción": resistencia_preinsercion,
        "Distancia mínima en aire - Entre polos (mm)": distancia_entre_polos,
        "Distancia mínima en aire - A tierra (mm)": distancia_a_tierra,
        "Distancia mínima en aire - A través del polo (mm)": distancia_a_traves_polo,
        "Clase de severidad de contaminación del sitio (SPS)": sps_clase,
        "Distancia mínima de fuga (mm)": distancia_minima_fuga,
        "Desempeño sísmico según IEEE-693-Vigente (**)": desempeno_sismico_ieee,
        "Frecuencia natural de vibración (Hz)": frecuencia_natural_vibracion,
        "Coeficiente de amortiguamiento crítico (%)": coef_amortiguamiento_critico,
        "Cargas admisibles en bornes - Carga estática admisible (N)": carga_estatica_admisible,
        "Cargas admisibles en bornes - Carga dinámica admisible (N)": carga_dinamica_admisible,
        "Fuerzas asociadas a la operación del equipo - Vertical (N)": fuerza_vertical,
        "Fuerzas asociadas a la operación del equipo - Horizontal (N)": fuerza_horizontal,
        "Masa neta de un polo completo con estructura (kg)": masa_neta_polo,
        "Dimensiones para transporte (Alto x Ancho x Largo) [mm]": dimensiones_transporte,
        "Masa neta para transporte (kg)": masa_neta_transporte,
        "Volumen total para transporte (m³)": volumen_total_transporte,
        "Campo eléctrico a 1 metro de separación del piso (kV/m)": campo_electrico_1m
            
    }
    
    # 📤 Función para exportar Excel con estilo personalizado
    def exportar_excel(datos, fuente="Calibri", tamaño=9):
        # Diccionario de unidades (puedes ampliarlo según tus campos)
        unidades = {
            "Tensión asignada (Ur) [kV]": "kV",
            "Altura de instalación (m.s.n.m)": "m.s.n.m",
            "Temperatura mínima anual (°C)": "°C",
            "Temperatura máxima anual (°C)": "°C",
            "Temperatura media (24 h) (°C)": "°C",
            "Frecuencia asignada (fr)": "Hz",
            "Corriente asignada en servicio continuo (Ir)": "A",
            "Poder de corte asignado en cortocircuito (Ics)": "kA",
            "Duración del cortocircuito asignado (Ics)": "s",
            "Porcentaje de corriente aperiódica (%)": "%",
            "Distancia mínima en aire - Entre polos (mm)": "mm",
            "Distancia mínima de fuga (mm)": "mm",
            "Campo eléctrico a 1 metro de separación del piso (kV/m)": "kV/m",
            "Masa neta para transporte (kg)": "kg",
            "Volumen total para transporte (m³)": "m³",
            "Dimensiones para transporte (Alto x Ancho x Largo) [mm]": "mm",
            "Masa neta de un polo completo con estructura (kg)": "kg"
            # Añade más unidades según tus campos
        }
    
        # Crear DataFrame con estructura personalizada
        df = pd.DataFrame([
            {
                "ÍTEM": i + 1,
                "DESCRIPCIÓN": campo,
                "UNIDAD": unidades.get(campo, ""),
                "REQUERIDO": valor,
                "OFRECIDO": ""  # Columna vacía para completar manualmente
            }
            for i, (campo, valor) in enumerate(datos.items())
        ])
    
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="CTG", startrow=6)
            wb = writer.book
            ws = writer.sheets["CTG"]
    
            # 🖼️ Insertar imagen del logo (opcional)
            logo_path = "siemens_logo.png"
            try:
                img = Image(logo_path)
                img.width = 300
                img.height = 100
                ws.add_image(img, "C1")
            except FileNotFoundError:
                st.warning("⚠️ No se encontró el logo 'siemens_logo.png'. Asegúrate de subirlo al repositorio.")
    
            # 🟪 Caja de título
            ws.merge_cells("A2:E4")
            cell = ws.cell(row=2, column=1)
            cell.value = "FICHA TÉCNICA INTERRUPTOR DE POTENCIA"
            cell.font = Font(name=fuente, bold=True, size=14, color="000000")
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
            # 🏷️ Subtítulo técnico
            ws.merge_cells("A5:D5")
            ws["A5"] = f"CARACTERÍSTICAS GARANTIZADAS"
            ws["A5"].font = Font(name=fuente, bold=True, size=12)
            ws["A5"].alignment = Alignment(horizontal="center")
    
            # 🎨 Encabezados con estilo
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
            ws.column_dimensions["A"].width = 5
            ws.column_dimensions["B"].width = 55
            ws.column_dimensions["C"].width = 12
            ws.column_dimensions["D"].width = 15
            ws.column_dimensions["E"].width = 15
    
            # 📋 Formato de filas con fuente personalizada
            for row in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=5):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.font = Font(name=fuente, size=tamaño)
                row[0].alignment = Alignment(horizontal="center", vertical="center")
                row[2].alignment = Alignment(horizontal="center", vertical="center")
                row[3].alignment = Alignment(horizontal="center", vertical="center")
                row[4].alignment = Alignment(horizontal="center", vertical="center")
                
            # Ajuste dinámico de altura de filas según contenido
            for row in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=5):
                max_lines = 1  # Mínimo una línea
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        # Estimar número de líneas según ancho de columna (por ejemplo, 55 caracteres en columna B)
                        wrapped = textwrap.wrap(cell.value, width=55)
                        max_lines = max(max_lines, len(wrapped))
                # Ajustar altura de la fila (aprox. 15 puntos por línea)
                ws.row_dimensions[cell.row].height = max_lines * 15

    
        output.seek(0)
        return output
    
    # 📥 Botón para generar y descargar
    fuente = "Calibri"
    tamaño = 9
    if st.button("📊 Generar archivo CTG"):
        archivo_excel = exportar_excel(ficha_cb, fuente=fuente, tamaño=tamaño)
        nivel_tension = ficha_cb.get("Nivel de tensión (kV)", "XX")
        st.download_button(
            label="📥 Descargar archivo CTG en Excel",
            data=archivo_excel,
            file_name=f"CTG_{nivel_tension}kV.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
