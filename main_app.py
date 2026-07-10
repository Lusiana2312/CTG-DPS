import streamlit as st

# 1️⃣ PRIMERA LÍNEA DE STREAMLIT (SIEMPRE ABAJO DE LOS IMPORTS)
st.set_page_config(page_title="Generador CTG", layout="wide")

# 🔐 Configuración de Login
usuarios_autorizados = {
    "lusiana": "clave123",
    "fer": "hola6"
}

st.title("🔐 Acceso privado")

# Usamos session_state para que no pida login a cada rato
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    usuario = st.text_input("Usuario")
    clave = st.text_input("Contraseña", type="password")
    
    if st.button("Ingresar"):
        if usuario in usuarios_autorizados and usuarios_autorizados[usuario] == clave:
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("❌ Credenciales incorrectas")
    st.stop() # Detiene la ejecución aquí si no está autenticado

# --- Si llega aquí, es que el usuario ya está logueado ---

st.success(f"✅ Acceso concedido")

# 🧭 Selector de equipo
equipo = st.selectbox("Selecciona el tipo de equipo", 
                     ["Descargador de sobretensiones", 
                      "Transformador de corriente", 
                      "Transformador de tensión", 
                      "Interruptor", 
                      "Seccionador"])

# ▶️ Ejecutar la función correspondiente
try:
    if equipo == "Descargador de sobretensiones":
        import generador_dps
        generador_dps.mostrar_app()

    elif equipo == "Transformador de corriente":
        import generador_ct
        generador_ct.mostrar_app()

    elif equipo == "Transformador de tensión":
        import generador_pt
        generador_pt.mostrar_app()
        
    elif equipo == "Interruptor":
        import generador_cb
        generador_cb.mostrar_app()
        
    elif equipo == "Seccionador":
        import generador_ds
        generador_ds.mostrar_app()

except ModuleNotFoundError as e:
    st.error(f"❌ No se encontró el módulo: {e.name}")
except Exception as e:
    st.error(f"⚠️ Ocurrió un error inesperado: {e}")






