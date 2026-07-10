import streamlit as st

st.set_page_config(page_title="Generador CTG", layout="wide")
st.title("🔐 Acceso privado")


# 🔐 Login
usuarios_autorizados = {
    "lusiana": "clave123",
    "fer": "hola6"
}

usuario = st.text_input("Usuario")
clave = st.text_input("Contraseña", type="password")

if usuario not in usuarios_autorizados or usuarios_autorizados[usuario] != clave:
    st.warning("🔒 Ingresa tus credenciales para continuar")
    st.stop()

st.success("✅ Acceso concedido")

# 🧭 Selector de equipo
equipo = st.selectbox("Selecciona el tipo de equipo", ["Descargador de sobretensiones", "Transformador de corriente", "Transformador de tensión", "Interruptor", "Seccionador"])

# ▶️ Ejecutar solo la función correspondiente
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
except AttributeError:
    st.error("⚠️ El módulo existe pero no tiene la función 'mostrar_app()'. Verifica que esté correctamente definida.")
except Exception as e:
    st.error(f"⚠️ Ocurrió un error inesperado: {e}")







