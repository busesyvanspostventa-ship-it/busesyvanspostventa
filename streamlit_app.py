from datetime import datetime
from zoneinfo import ZoneInfo
from hashlib import sha256
import re
import streamlit as st
from generador import generar_informe

MODELOS = ['U9', 'U10', 'U12', 'D9', 'D10']
RESPONSABLES = {
    'Alejandro Vasquez Andrews': {'cargo': 'Jefe Asistencia Técnica Buses y Vans (iveco)', 'correo': 'avasqueza@andesmotor.cl'},
    'Ignacio Gutiérrez Gaete': {'cargo': 'Supervisor de Flotas Buses', 'correo': 'ignacio.gutierrez@andesmotor.cl'},
    'Patricio Gormaz Pardo': {'cargo': 'Ingeniero de Operaciones y Desarrollo Técnico', 'correo': 'pgormaz@andesmotor.cl'},
    'Dimas Toro Echeverria': {'cargo': 'Ingeniero de Operaciones y Desarrollo Técnico', 'correo': 'dimas.toro@andesmotor.cl'},
    'David Esparza': {'cargo': 'Supervisor de Flotas Buses', 'correo': 'desparza@andesmotor.cl'},
    'Danilo Flores Burgos': {'cargo': 'Ingeniero de Operaciones y Desarrollo Técnico', 'correo': 'danilo.flores@andesmotor.cl'},
}

st.set_page_config(page_title='Informes técnicos Foton', page_icon='🚌', layout='centered')
st.title('Informes técnicos Foton')
st.caption('Postventa · Andes Motor · Plantilla Word')
st.write('Completa los antecedentes, agrega las fotos y descarga el informe editable.')
st.info('Marca fija: FOTON. Los datos se usan en esta sesión; descarga el informe para conservarlo.')
a,b = st.columns(2)
modelo = a.selectbox('Modelo *', ['Selecciona un modelo'] + MODELOS)
fecha = b.date_input('Fecha de evaluación', datetime.now(ZoneInfo('America/Santiago')).date())
vin = a.text_input('VIN', max_chars=17).strip().upper()
patente = b.text_input('Patente').strip().upper()
sistema = st.text_input('Sistema afectado *', placeholder='Ej.: puertas, carrocería, motor o baterías')
st.subheader('Motivo de admisión')
motivo = st.text_area('Síntoma o reclamo informado *', height=110)
st.subheader('Entidades')
participantes = st.text_area('Cliente, taller y otras personas participantes', placeholder='Cliente / operador:\nTaller / sucursal:')
responsable = st.selectbox('Responsable del informe *', ['Selecciona un responsable'] + list(RESPONSABLES))
perfil = RESPONSABLES.get(responsable)
entidades = participantes.strip()
if perfil:
    st.text_input('Cargo', value=perfil['cargo'], disabled=True)
    st.text_input('Correo', value=perfil['correo'], disabled=True)
    entidades = '\n'.join(filter(None, [entidades, f'Responsable: {responsable}', f'Cargo: {perfil["cargo"]}', f'Correo: {perfil["correo"]}']))
st.subheader('Análisis')
analisis = st.text_area('Inspección, diagnóstico, intervención y resultado *', height=190, placeholder='Inspección realizada:\nDiagnóstico confirmado o pendiente:\nIntervención:\nResultado y pendientes:')
st.subheader('Evidencia')
archivos = st.file_uploader('Fotografías JPG o PNG (máximo 12; 10 MB por foto)', type=['jpg','jpeg','png'], accept_multiple_files=True)
fotos=[]
errores=[]
if len(archivos)>12: errores.append('Adjunta un máximo de 12 fotografías.')
for i,f in enumerate(archivos[:12]):
    contenido = f.getvalue()
    if len(contenido)>10*1024*1024:
        errores.append(f'{f.name}: supera 10 MB.'); continue
    with st.expander(f'Fotografía {i+1}: {f.name}', expanded=True):
        st.image(contenido, width=280)
        desc = st.text_input('Descripción del hallazgo', key=f'foto_{i}_{sha256(contenido).hexdigest()[:16]}')
        fotos.append((contenido,desc))
datos=dict(modelo=modelo,fecha=fecha.strftime('%d-%m-%Y'),vin=vin,patente=patente,sistema=sistema,motivo=motivo,entidades=entidades,analisis=analisis)
firma=sha256(repr((datos,[(sha256(x).hexdigest(),y) for x,y in fotos])).encode()).hexdigest()
confirmado=st.checkbox('Revisé los antecedentes técnicos y las descripciones de las fotos.')
if st.button('Generar informe Word', type='primary'):
    st.session_state.pop('informe', None)
    if modelo not in MODELOS: errores.append('Selecciona un modelo Foton.')
    if perfil is None: errores.append('Selecciona al responsable del informe.')
    if not all(x.strip() for x in (modelo,sistema,motivo,analisis)): errores.append('Completa los campos obligatorios (*).')
    if not vin and not patente: errores.append('Ingresa al menos VIN o patente.')
    if vin and (len(vin)!=17 or not re.fullmatch('[A-HJ-NPR-Z0-9]{17}',vin)): errores.append('Revisa el VIN: debe tener 17 caracteres y no contener I, O ni Q.')
    if any(not d.strip() for _,d in fotos): errores.append('Agrega la descripción de cada fotografía.')
    if not confirmado: errores.append('Confirma la revisión de los antecedentes.')
    if errores:
        for e in errores: st.error(e)
    else:
        try:
            documento=generar_informe(datos,fotos)
            identificador=re.sub(r'[^A-Za-z0-9_-]','',patente or vin)
            nombre=f'Informe_Foton_{identificador}_{fecha:%Y%m%d}.docx'
            st.session_state.informe=(firma,documento,nombre)
        except Exception:
            st.error('No se pudo generar el informe. Revisa que las fotos sean JPG/PNG válidos y que plantilla_foton.docx esté junto a app.py.')
if 'informe' in st.session_state:
    firma_guardada,documento,nombre=st.session_state.informe
    if firma_guardada==firma and confirmado and not errores:
        st.success('Informe listo para descargar.')
        st.download_button('Descargar Word', documento, nombre, mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    else:
        st.warning('Los datos cambiaron. Genera nuevamente el informe para descargar la versión actualizada.')
