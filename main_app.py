import streamlit as st

# 🔐 Login
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

# ▶️ Ejecutar solo la función correspondiente

    if equipo == "CTG DPS":
        import generador_ctg
        generador_ctg.mostrar_app()

    elif equipo == "CT":
        import generador_ct
        generador_ct.mostrar_app()

    elif equipo == "PT":
        import generador_pt
        generador_pt.mostrar_app()

