"""Identificación por usuario conocido, sin contraseña ni verificación de identidad."""
import streamlit as st


def exigir_ingreso(responsables):
    usuarios = {
        perfil['correo'].split('@')[0].strip().lower(): nombre
        for nombre, perfil in responsables.items()
    }
    usuario_actual = st.session_state.get('_usuario_identificado')
    if usuario_actual in usuarios:
        st.sidebar.write(f'Responsable: {usuarios[usuario_actual]}')
        st.sidebar.caption(f'Usuario: {usuario_actual}')
        if st.sidebar.button('Cambiar usuario'):
            st.session_state.clear()
            st.rerun()
        return usuario_actual, usuarios[usuario_actual]

    st.subheader('Ingreso de usuario')
    st.caption('Escribe la parte anterior al @ de tu correo. No se solicita contraseña.')
    with st.form('identificar_usuario'):
        usuario = st.text_input('Usuario', placeholder='Ej.: desparza').strip().lower()
        if usuario in usuarios:
            st.caption(f'Responsable: {usuarios[usuario]}')
        entrar = st.form_submit_button('Ingresar', type='primary')
    if entrar:
        if usuario in usuarios:
            st.session_state.clear()
            st.session_state['_usuario_identificado'] = usuario
            st.rerun()
        else:
            st.error('Usuario no reconocido. Ingresa solo la parte anterior al @ de tu correo registrado.')
    st.stop()
