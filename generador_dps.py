
#empieza codigo
import streamlit as st
import pandas as pd 		
from io import BytesIO
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
    
def mostrar_app():
    st.title("Generador CTG DPS")
    st.set_page_config(page_title="Generador CTG", layout="wide")
    st.title("📄 Generador de Archivos CTG")
    
    # 1. Altura de instalación
    altura_instalacion = st.number_input("### 🧱 Altura sobre el nivel del mar (m.s.n.m)", min_value=0, value=1000, step=10)
    
    # 2. Fabricante
    fabricante = st.text_input("### 🏢 Fabricante")
    
    # 3. Referencia
    referencia = st.text_input("### 🏷️ Referencia")
    
    # 4. Norma de fabricación (fijo)
    norma_fabricacion = "IEC 60099-4"
    st.text("### 📘 Norma de fabricación: " + norma_fabricacion)
    
    # 5. Norma de calidad
    norma_calidad = st.selectbox("### 📘 Norma de calidad", ["IEC 9001", "ISO 9001"])
    
    # 6. Tipo de ejecución
    tipo_ejecucion = st.selectbox("### 🏗️ Tipo de ejecución", ["Exterior", "Interior"])
    
    # 7. Frecuencia asignada (fijo)
    frecuencia_asignada = "60 Hz"
    st.text("### ⚡ Frecuencia asignada: " + frecuencia_asignada)
    
    # 8. Material cubierta
    material_cubierta = st.selectbox("### 🧩 Material de la cubierta", ["Polimérico", "Porcelana"])
    
    # 9. Número de columnas
    numero_columnas = st.selectbox("### 🔢 Número de columnas", [1, 2])
    
    # 10. Número de cuerpos
    numero_cuerpos = st.text_input("### 🔢 Número de cuerpos")
    
    # 11. Tensión más elevada para el material (Um)
    um = st.selectbox("### ⚡ Tensión más elevada para el material (Um)", ["145 kV", "245 kV", "550 kV"])
    
    # 12. Tensión asignada (Ur)
    ur_por_um = {"145 kV": "110 kV", "245 kV": "198 kV", "550 kV": "210 kV"}
    ur = ur_por_um.get(um, "")
    st.text(f"### ⚡ Tensión asignada (Ur): {ur}")
    
    # 13. Tensión continua de operación (Uc)
    uc = st.text_input("### ⚡ Tensión continua de operación (Uc)")
    
    # 14. Corriente de descarga asignada (In)
    in_por_um = {"145 kV": "10 kA", "245 kV": "20 kA", "550 kV": "30 kA"}
    in_corriente = in_por_um.get(um, "")
    st.text(f"### ⚡ Corriente de descarga asignada (In): {in_corriente}")
    
    # 15. Corriente asignada del dispositivo de alivio de presión (0.2 seg)
    alivio_por_um = {"245 kV": "20 kA", "550 kV": "30 kA"}
    corriente_alivio = alivio_por_um.get(um, "")
    st.text(f"### ⚡ Corriente asignada del dispositivo de alivio de presión (0.2 seg): {corriente_alivio if corriente_alivio else 'No aplica'}")
    
    # 16. Tensión residual al impulso de corriente de escalón (10 kA)
    ures_escalon = st.text_input("### ⚡ Tensión residual al impulso de corriente de escalón (10 kA)")
    
    # 17. Tensión residual al impulso tipo maniobra (Ures)
    st.markdown("### ⚡ Tensión residual al impulso tipo maniobra (Ures)")
    ures_maniobra_250 = st.text_input("Ures - Para 250 A", value="")
    ures_maniobra_500 = st.text_input("Ures - Para 500 A", value="")
    ures_maniobra_1000 = st.text_input("Ures - Para 1000 A", value="")
    ures_maniobra_2000 = st.text_input("Ures - Para 2000 A", value="")
    
    # 18. Tensión residual al impulso tipo rayo (Ures)
    st.markdown("### ⚡ Tensión residual al impulso tipo rayo (Ures)")
    ures_rayo_5ka = st.text_input("Ures - 5 kA", value="")
    ures_rayo_10ka = st.text_input("Ures - 10 kA", value="")
    ures_rayo_20ka = st.text_input("Ures - 20 kA", value="")
    
    # 19. Clase de descarga de línea
    clase_descarga = st.selectbox("### ⚡ Clase de descarga de línea", [1, 2, 3, 4, 5])
    
    # 20. Capacidad mínima de disipación de energía
    capacidad_duracion = "≥10 kJ/kV"
    st.text(f"### ⚡ Capacidad mínima de disipación de energía (2 impulsos largos): {capacidad_duracion}")
    
    # 21. Transferencia de carga repetitiva Qrs
    qrs = "≥2.4"
    st.text(f"### ⚡ Transferencia de carga repetitiva Qrs: {qrs}")
    
    # 22. Mínima sobretensión temporal soportada
    st.markdown("### ⚡ Mínima sobretensión temporal soportada luego de absorber la energía asignada")
    sobretension_1s = st.text_input("Durante 1s", value="")
    sobretension_10s = st.text_input("Durante 10s", value="")
    
    # 23. Capacitancia fase-tierra
    capacitancia = st.text_input("### ⚡ Capacitancia fase-tierra", value="")
    
    # 24. Distancia de arco
    distancia_arco = st.text_input("### ⚡ Distancia de arco (con anillos anticorona si aplica)", value="")
    
    # 25. Clase de severidad de contaminación del sitio (SPS)
    st.markdown("### 🌫️ Clase de severidad de contaminación del sitio (SPS)")
    sps_opciones = {"Bajo": 16, "Medio": 20, "Pesado": 25, "Muy Pesado": 31}
    sps_seleccion = st.selectbox("Selecciona la clase SPS", list(sps_opciones.keys()))
    valor_sps = sps_opciones[sps_seleccion]
    
    # 26. Distancia mínima de fuga = Um * SPS
    st.markdown("### 📏 Distancia mínima de fuga requerida")
    um_valores = {"145 kV": 145, "245 kV": 245, "550 kV": 550}
    um_num = um_valores.get(um, 0)
    distancia_fuga = um_num * valor_sps
    st.text(f"Distancia mínima de fuga: {distancia_fuga} mm")
    
    # 27. Aislamiento de la envolvente
    st.markdown("### 🧪 Aislamiento de la envolvente (con anillos anticorona si aplica)")
    ud = st.text_input("Tensión asignada soportada a la frecuencia industrial (Ud)", value="")
    up = st.text_input("Tensión asignada soportada al impulso tipo rayo (Up)", value="")
    us = st.text_input("Tensión asignada soportada al impulso tipo maniobra (Us)", value="")
    
    # 28. Datos sísmicos
    st.markdown("### 🌍 Datos sísmicos según IEEE-693 vigente")
    desempeno_sismico = st.selectbox("Desempeño sísmico", ["Alto", "Moderado", "Bajo"])
    frecuencia_natural = st.text_input("Frecuencia natural de vibración", value="")
    amortiguamiento_critico = st.text_input("Coeficiente de amortiguamiento crítico", value="")
    
    # 29. Cargas admisibles en bornes
    st.markdown("### 🧱 Cargas admisibles en bornes")
    carga_estatica = st.text_input("Carga estática admisible", value="")
    carga_dinamica = st.text_input("Carga dinámica admisible", value="")
    
    # 30. Altura total
    altura_total = st.text_input("### 📏 Altura total", value="")
    
    # 31. Dimensiones para transporte
    dimensiones_transporte = st.text_input("### 📦 Dimensiones para transporte (Alto x Ancho x Largo)", value="")
    
    # 32. Masa neta para transporte
    masa_transporte = st.text_input("### ⚖️ Masa neta para transporte", value="")
    
    # 33. Volumen total
    volumen_total = st.text_input("### 📦 Volumen total", value="")
    
    # 34. Anillo corona y de distribución de campo
    anillo_corona = st.text_input("### 🧲 Anillo corona y de distribución de campo", value="")
    
    # 35. Contador de descargas
    contador_descargas = st.selectbox("### 🔌 Contador de descargas", ["Sí", "No"])
    
    # 36. Accesorios
    accesorios = st.text_input("### 🧰 Accesorios", value="")

    
    
    # Consolidar todos los datos en un solo diccionario
    datos = {
        # 1–3
        "Altura sobre el nivel del mar (m.s.n.m)": altura_instalacion,
        "Fabricante": fabricante,
        "Referencia": referencia,
    
        # 4–7
        "Norma de fabricación": norma_fabricacion,
        "Norma de calidad": norma_calidad,
        "Tipo de ejecución": tipo_ejecucion,
        "Frecuencia asignada (Hz)": frecuencia_asignada,
    
        # 8–10
        "Material de la cubierta": material_cubierta,
        "Número de columnas": numero_columnas,
        "Número de cuerpos": numero_cuerpos,
    
        # 11–12
        "Tensión más elevada para el material (Um)": um,
        "Tensión asignada (Ur)": ur,
    
        # 13–15
        "Tensión continua de operación (Uc)": uc,
        "Corriente de descarga asignada (In)": in_corriente,
        "Corriente asignada del dispositivo de alivio de presión (0.2 seg)": corriente_alivio,
    
        # 16
        "Tensión residual al impulso de corriente de escalón (10 kA)": ures_escalon,
    
        # 17
        "Tensión residual al impulso tipo maniobra (Ures) - 250 A": ures_maniobra_250,
        "Tensión residual al impulso tipo maniobra (Ures) - 500 A": ures_maniobra_500,
        "Tensión residual al impulso tipo maniobra (Ures) - 1000 A": ures_maniobra_1000,
        "Tensión residual al impulso tipo maniobra (Ures) - 2000 A": ures_maniobra_2000,
    
        # 18
        "Tensión residual al impulso tipo rayo (Ures) - 5 kA": ures_rayo_5ka,
        "Tensión residual al impulso tipo rayo (Ures) - 10 kA": ures_rayo_10ka,
        "Tensión residual al impulso tipo rayo (Ures) - 20 kA": ures_rayo_20ka,
    
        # 19–21
        "Clase de descarga de línea": clase_descarga,
        "Capacidad mínima de disipación de energía (kJ/kV)": capacidad_duracion,
        "Transferencia de carga repetitiva Qrs": qrs,
    
        # 22
        "Sobretensión temporal soportada - 1s": sobretension_1s,
        "Sobretensión temporal soportada - 10s": sobretension_10s,
    
        # 23–24
        "Capacitancia fase-tierra": capacitancia,
        "Distancia de arco (con anillos anticorona)": distancia_arco,
    
        # 25–26
        "Clase SPS": sps_seleccion,
        "Valor SPS": valor_sps,
        "Distancia mínima de fuga (mm)": distancia_fuga,
    
        # 27
        "Tensión soportada a frecuencia industrial (Ud)": ud,
        "Tensión soportada al impulso tipo rayo (Up)": up,
        "Tensión soportada al impulso tipo maniobra (Us)": us,
    
        # 28
        "Desempeño sísmico": desempeno_sismico,
        "Frecuencia natural de vibración": frecuencia_natural,
        "Coeficiente de amortiguamiento crítico": amortiguamiento_critico,
    
        # 29
        "Carga estática admisible": carga_estatica,
        "Carga dinámica admisible": carga_dinamica,
    
        # 30–34
        "Altura total": altura_total,
        "Dimensiones para transporte (Alto x Ancho x Largo)": dimensiones_transporte,
        "Masa neta para transporte": masa_transporte,
        "Volumen total": volumen_total,
        "Anillo corona y de distribución de campo": anillo_corona,
    
        # 35–36
        "Contador de descargas": contador_descargas,
        "Accesorios": accesorios
    }
        
        
    # 📤 Función para exportar Excel con estilo personalizado
    def exportar_excel(datos, fuente="Calibri", tamaño=9):
        unidades = {
            "Nivel de tensión (kV)": "kV",
            "Tensión asignada (Ur)": "kV",
            "Altura de instalación (m.s.n.m)": "m.s.n.m",
            "Coeficiente Ka": "",
            "Coeficiente Km": "",
            "Distancia mínima de fuga (mm)": "mm",
            "Tensión residual al impulso de corriente de escalón (10 kA)": "kV",
            "Tensión residual al impulso tipo maniobra (250 A)": "kV",
            "Tensión residual al impulso tipo maniobra (500 A)": "kV",
            "Tensión residual al impulso tipo maniobra (1000 A)": "kV",
            "Tensión residual al impulso tipo maniobra (2000 A)": "kV",
            "Tensión residual al impulso tipo rayo (5 kA)": "kV",
            "Tensión residual al impulso tipo rayo (10 kA)": "kV",
            "Tensión residual al impulso tipo rayo (20 kA)": "kV",
            "Tensión asignada soportada a la frecuencia industrial (Ud)": "kV",
            "Tensión asignada soportada al impulso tipo rayo (Up)": "kV",
            "Tensión asignada soportada al impulso tipo maniobra (Us)": "kV"
        }
        
        df = pd.DataFrame([
            {
                "ÍTEM": i + 1,
                "DESCRIPCIÓN": campo,
                "UNIDAD": unidades.get(campo, ""),
                "REQUERIDO": valor,
                "OFRECIDO": "" #Columna vacía
            }
            for i, (campo, valor) in enumerate(datos.items())
        ])
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="CTG", startrow=6)
            wb = writer.book
            ws = writer.sheets["CTG"]
        
            # 🖼️ Insertar imagen del logo
            logo_path = "siemens_logo.png"
            try:
                img = Image(logo_path)
                img.width = 300
                img.height = 100
                ws.add_image(img, "C1")
            except FileNotFoundError:
                st.warning("⚠️ No se encontró el logo 'siemens_logo.png'. Asegúrate de subirlo al repositorio.")
                
            #🧱 Crear borde negro alrededor de A2:E4
            black_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )
        
            for row in ws.iter_rows(min_row=2, max_row=4, min_col=1, max_col=5):
                for cell in row:
                    cell.border = black_border
        
            # 🟪 Caja de título
            ws.merge_cells("A2:E4")
            cell = ws.cell(row=2, column=1)
            cell.value = "CARACTERÍSTICAS GARANTIZADAS"
            cell.font = Font(name=fuente, bold=True, size=14, color="000000")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
            # 🏷️ Título técnico
            ws.merge_cells("A5:D5")
            ws["A5"] = f"DESCARGADORES DE SOBRETENSIÓN {datos['Nivel de tensión (kV)']} kV"
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
        
            # 📋 Formato de filas con fuente personalizada
            for row in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=5):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.font = Font(name=fuente, size=tamaño)
                row[0].alignment = Alignment(horizontal="center", vertical="center")
                row[2].alignment = Alignment(horizontal="center", vertical="center")
                row[3].alignment = Alignment(horizontal="center", vertical="center")
                row[4].alignment = Alignment(horizontal="center", vertical="center")  # Alineación para OFRECIDO
        
        
        output.seek(0)
        return output
        
        
    # 📥 Botón para generar y descargar
    fuente = "Calibri"
    tamaño = 9
    if st.button("📊 Generar archivo CTG"):
        archivo_excel = exportar_excel(datos, fuente=fuente, tamaño=tamaño)
        nivel_tension = datos.get("Nivel de tensión (kV)", "XX")
        st.download_button(
            label="📥 Descargar archivo CTG en Excel",
            data=archivo_excel,
            file_name=f"CTG_{nivel_tension}kV.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        
        
        
        
        
        
    
    
    
    
    








