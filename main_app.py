import streamlit as st

# 🧭 Selector de equipo
equipo = st.selectbox("Selecciona el tipo de equipo", ["CTG DPS", "Otro equipo"])

# 📦 Importar función según el equipo seleccionado
if equipo == "Descargador de sobretensiones":
    from generador_ctg import exportar_excel, obtener_nombre_archivo
elif equipo == "Otro equipo":
    from generador_otro_equipo import exportar_excel, obtener_nombre_archivo
