import streamlit as st

# 🧭 Selector de equipo
equipo = st.selectbox("Selecciona el tipo de equipo", ["CTG DPS", "Otro equipo"])

# 📦 Importar función según el equipo seleccionado
if equipo == "CTG DPS":
    from generador_ctg import exportar_excel, obtener_nombre_archivo
elif equipo == "Otro equipo":
    from generador_otro_equipo import exportar_excel, obtener_nombre_archivo

# 🎨 Parámetros de estilo
fuente = "Calibri"
tamaño = 9

# 📥 Botón para generar y descargar
if st.button("📊 Generar archivo Excel"):
    archivo_excel = exportar_excel(fuente=fuente, tamaño=tamaño)
    nombre_archivo = obtener_nombre_archivo()
    st.download_button(
        label="📥 Descargar archivo Excel",
        data=archivo_excel,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
