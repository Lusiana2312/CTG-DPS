# empieza codigo
import streamlit as st
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime
import pandas as pd
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import textwrap


################## CTG DISCONNECTOW SWITCH// SECCIONADOR
def mostrar_app():
    st.set_page_config(page_title="Generador CTG - Interruptor de Potencia", layout="wide")

    st.title("📄 Generador de Ficha CTG")
    st.subheader("Seccionador")
    # 1. Fabricante
    fabricante = "Indicar"
    st.text("### 🏢 Fabricante: " + fabricante)
    # 2. País
    pais = "Indicar"
    st.text("### País: " + pais)
    # 3. Referencia
    referencia = "Indicar"
    st.text("### Referencia: " + referencia)
    # 4. Norma de fabricación
    norma_fabricacion = "IEC 62271-102"
    st.markdown(f"**Norma de fabricación:** {norma_fabricacion}")
    # 5. Norma de calidad
    norma_calidad = "ISO 9001"
    st.markdown(f"**Norma de calidad:** {norma_calidad}")
    #6. Número de polos
    num_polos = "3"
    st.text("### Número de polos: " + num_polos)

    # 7. Instalación
    instalacion = st.selectbox("Tipo de ejecución", ["Exterior", "Interior"])

    # 8. Tipo de accionamiento
    accionamiento = st.selectbox("Tipo de accionamiento", ["Monopolar", "Tripolar"])

    # 9. Tipo de construcción para seccionador de conexión
    conexion = st.selectbox("Tipo de construcción para seccionador de conexión", ["Pantógrafo", "Semi-pantógrafo", "Rotación Central"])
    
    # 10. Altura 
    altura_instalacion = st.number_input("Altura de instalación (m.s.n.m)", min_value=0, value=1000)

    # 11. Temperatura de operación
    st.markdown("### 🌡️ Temperatura de operación")
    temp_min = -5
    st.text(f"### Temperatura mínima anual (°C): {temp_min}")
    temp_max = +40
    st.text(f"### Temperatura máxima anual (°C): {temp_max}")
    temp_media = +35
    st.text(f"### Temperatura media (24 h) (°C): {temp_media}")

    # 12. Frecuencia
    frecuencia_asignada = "60 Hz"
    st.text(f"### Frecuencia asignada (fr): " + frecuencia_asignada)

    #13. Clafisicación ambiente sitio de instalación para corrosión según ISO 12944
    corrosion ="Indicar"
    st.text("### Clafisicación ambiente sitio de instalación para corrosión según ISO 12944: " + corrosion)

    #14. Nivel de polución sitio de instalación según IEC 60815
    polucion = "Indicar"
    st.text("### Nivel de polución sitio de instalación según IEC 60815: " + polucion)

    # 15. Tensión asignada Ur
    ur = st.selectbox("Tensión asignada (Ur)", options=["145 kV", "245 kV", "550 kV"])

     # 16. Tensión asignada a frecuencia industrial
    # Asignación automática de Ud según Ur
    ud_por_ur = {
        "145 kV": {"fase_tierra_ud": "275", "distancia_seccionamiento": "315"},
        "245 kV": {"fase_tierra_ud": "460", "distancia_seccionamiento": "530"},
        "550 kV": {"fase_tierra_ud": "620 kV", "distancia_seccionamiento": "800 kV"}
    }
    ud_valores = ud_por_ur.get(ur,{"fase_tierra_ud": "", "distancia_seccionamiento": ""})
    st.markdown("#### Tensión asignada soportada a frecuencia industrial (Ud)")
    st.markdown(f"a) A tierra y entre polos: **{ud_valores['fase_tierra_ud']}**")
    st.markdown(f"b) A través de la distancia de seccionamiento: **{ud_valores['distancia_seccionamiento']}**")

    # BOTÓN PARA GENERAR FICHA
    ficha_cb = {
        "Fabricante": fabricante,
        "País": pais,
        "Referencia": referencia,
        "Norma de fabricación": norma_fabricacion,
        "Norma de calidad": norma_calidad,
        "Número de polos": num_polos,
        "Instalación": instalacion,
        "Tipo de construcción para seccionador de conexión": conexion,
        "Tipo de accionamiento": accionamiento,
        "Temperatura mínima anual (°C)": temp_min,
        "Temperatura máxima anual (°C)": temp_max,
        "Temperatura media (24 h) (°C)": temp_media,
        "Frecuencia asignada": frecuencia_asignada,
        "Clafisicación ambiente sitio de instalación para corrosión según ISO 12944": corrosion,
        "Nivel de polución sitio de instalación según IEC 60815": polucion,
        "Tensión asignada (Ur)": ur
        
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
            



