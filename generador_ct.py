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

    # 4. CONFIGURACIÓN DE NÚCLEOS
    st.markdown("### 🧲 Configuración de núcleos")
    
    # 4. CONFIGURACIÓN DE NÚCLEOS
    st.markdown("### 🧲 Configuración de núcleos")
    
    num_nucleos = st.number_input("Número total de núcleos", min_value=1, max_value=10, step=1)
    
    nucleos = []
    
    for i in range(int(num_nucleos)):
        st.markdown(f"#### Núcleo {i+1}")
        tipo_nucleo = st.selectbox(f"Tipo de núcleo {i+1}", options=["Medida", "Protección"], key=f"tipo_{i}")
        clase = st.text_input(f"Clase del núcleo {i+1}", key=f"clase_{i}")
        relacion_transformacion = st.text_input(f"Relación de transformación {i+1} (Ej: 1000/1)", key=f"relacion_{i}")
        carga = st.text_input(f"Carga (VA) del núcleo {i+1}", key=f"carga_{i}")
        precision = st.text_input(f"Precisión del núcleo {i+1}", key=f"precision_{i}")
    
        nucleos.append({
            "Número": i + 1,
            "Tipo": tipo_nucleo,
            "Clase": clase,
            "Relación": relacion_transformacion,
            "Carga (VA)": carga,
            "Precisión": precision
        })
    
    # Mostrar resumen de núcleos
    st.markdown("### 📋 Resumen de núcleos configurados")
    for nucleo in nucleos:
        st.write(f"Núcleo {nucleo['Número']}: Tipo: {nucleo['Tipo']}, Clase: {nucleo['Clase']}, "
                 f"Relación: {nucleo['Relación']}, Carga: {nucleo['Carga (VA)']} VA, Precisión: {nucleo['Precisión']}")
    
    # BOTÓN PARA GENERAR FICHA
    ficha_cb = {
        "Fabricante": fabricante,
        "País": pais,
        "Referencia": referencia,
        "Norma de fabricación": norma_fabricacion,
        "Norma de calidad": norma_calidad,
        "Medio de extinción": medio_extincion,
        "Número de polos": num_polos,
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
            ws.print_title_rows = '1:7'
            ws.print_area = f"A1:E{ws.max_row}"

            
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
            ws.column_dimensions["A"].width = 4
            ws.column_dimensions["B"].width = 50
            ws.column_dimensions["C"].width = 10
            ws.column_dimensions["D"].width = 12
            ws.column_dimensions["E"].width = 12
    
            
            
            # 📋 Formato de filas con fuente personalizada y ajuste dinámico de altura
            for row in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=5):
                max_lines = 1  # Mínimo una línea por celda
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.font = Font(name=fuente, size=tamaño)
            
                    # Estimar número de líneas necesarias si el contenido es texto
                    if cell.value and isinstance(cell.value, str):
                        # Ajusta el ancho según la columna (por ejemplo, columna B tiene 55 caracteres de ancho)
                        if cell.column_letter == "B":
                            wrapped = textwrap.wrap(cell.value, width=55)
                            max_lines = max(max_lines, len(wrapped))
            
                # Ajustar altura de la fila según el contenido más largo
                ws.row_dimensions[row[0].row].height = max_lines * 15  # 15 puntos por línea aprox.
            
                # Alineación horizontal para columnas específicas
                row[0].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                row[2].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                row[3].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                row[4].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
                
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
            
