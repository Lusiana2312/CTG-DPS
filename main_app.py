import streamlit as st
streamlit run main_app.py

# 🔐 Login de usuario
usuarios_autorizados = {
    "lusiana": "clave123",
    "fer": "hola6"
}

st.set_page_config(page_title="Generador CTG", layout="wide")
st.title("🔐 Acceso privado")

usuario = st.text_input("Usuario")
clave = st.text_input("Contraseña", type="password")

if usuario not in usuarios_autorizados or usuarios_autorizados[usuario] != clave:
    st.warning("🔒 Ingresa tus credenciales para continuar")
    st.stop()

st.success("✅ Acceso concedido")

# 🧭 Selector de equipo
equipo = st.selectbox("Selecciona el tipo de equipo", ["CTG DPS", "CT"])

# ▶️ Ejecutar el código correspondiente según el equipo
if equipo == "CTG DPS":
    import generador_ctg  # Ejecuta el código de CTG DPS
elif equipo == "CT":
    import generador_ct  # Ejecuta el código de CT

