from datetime import datetime
from zoneinfo import ZoneInfo
from hashlib import sha256
from io import BytesIO, StringIO
import csv
import re

from PIL import Image
import streamlit as st

from generador import generar_informe
from base_datos import (
    guardar_informe,
    listar_informes,
    respaldo_base,
    actualizar_informe,
    eliminar_informe,
    estado_conexion,
)


MODELOS = ["U9", "U10", "U12", "D9", "D10"]

RESPONSABLES = {
    "Alejandro Vasquez Andrews": {
        "prefijo": "AV",
        "cargo": "Jefe Asistencia Técnica Buses y Vans (iveco)",
        "correo": "avasqueza@andesmotor.cl",
    },
    "Ignacio Gutiérrez Gaete": {
        "prefijo": "IG",
        "cargo": "Supervisor de Flotas Buses",
        "correo": "ignacio.gutierrez@andesmotor.cl",
    },
    "Patricio Gormaz Pardo": {
        "prefijo": "PG",
        "cargo": "Ingeniero de Operaciones y Desarrollo Técnico",
        "correo": "pgormaz@andesmotor.cl",
    },
    "Dimas Toro Echeverria": {
        "prefijo": "DT",
        "cargo": "Ingeniero de Operaciones y Desarrollo Técnico",
        "correo": "dimas.toro@andesmotor.cl",
    },
    "David Esparza": {
        "prefijo": "DE",
        "cargo": "Supervisor de Flotas Buses",
        "correo": "desparza@andesmotor.cl",
    },
    "Danilo Flores Burgos": {
        "prefijo": "DF",
        "cargo": "Ingeniero de Operaciones y Desarrollo Técnico",
        "correo": "danilo.flores@andesmotor.cl",
    },
}


st.set_page_config(
    page_title="Informes técnicos Foton",
    page_icon="🚌",
    layout="centered",
)

st.title("Informes técnicos Foton")
st.caption("Postventa · Andes Motor · Plantilla Word")

aviso = st.session_state.pop("aviso", None)
if aviso:
    st.success(aviso)

st.write(
    "Completa los antecedentes, agrega las fotos "
    "y descarga el informe editable."
)

a, b = st.columns(2)

modelo = a.selectbox(
    "Modelo *",
    ["Selecciona un modelo"] + MODELOS,
)
fecha = b.date_input(
    "Fecha de evaluación",
    datetime.now(ZoneInfo("America/Santiago")).date(),
)
vin = a.text_input("VIN", max_chars=17).strip().upper()
patente = b.text_input("Patente").strip().upper()

sistema = st.text_input(
    "Sistema afectado *",
    placeholder="Ej.: puertas, carrocería, motor o baterías",
)

st.subheader("Motivo de admisión")
motivo = st.text_area(
    "Síntoma o reclamo informado *",
    height=110,
)

st.subheader("Entidades")
entidades = st.text_area(
    "Áreas anexas o externos que participaron ",
    placeholder="Proveedor ;Fábrica",
).strip()

st.subheader("Análisis")
analisis = st.text_area(
    "Inspección, diagnóstico, intervención y resultado *",
    height=190,
    placeholder=(
        "Inspección realizada:\n"
        "Diagnóstico confirmado o pendiente:\n"
        "Intervención:\n"
        "Resultado y pendientes:"
    ),
)

st.subheader("Evidencia")
archivos = st.file_uploader(
    "Fotografías JPG o PNG (máximo 12; 10 MB por foto)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

fotos = []
errores = []

if len(archivos) > 12:
    errores.append("Adjunta un máximo de 12 fotografías.")

for i, archivo in enumerate(archivos[:12]):
    contenido = archivo.getvalue()

    if len(contenido) > 10 * 1024 * 1024:
        errores.append(f"{archivo.name}: supera 10 MB.")
        continue

    try:
        with Image.open(BytesIO(contenido)) as imagen:
            imagen.verify()
    except Exception:
        errores.append(f"{archivo.name}: no es una imagen válida.")
        continue

    with st.expander(
        f"Fotografía {i + 1}: {archivo.name}",
        expanded=True,
    ):
        st.image(contenido, width=280)
        descripcion = st.text_input(
            "Descripción del hallazgo",
            key=f"foto_{i}_{sha256(contenido).hexdigest()[:16]}",
        )
        fotos.append((contenido, descripcion))

st.subheader("Responsable del informe")
responsable = st.selectbox(
    "Responsable *",
    ["Selecciona un responsable"] + list(RESPONSABLES),
)
perfil = RESPONSABLES.get(responsable)

if perfil:
    st.text(f"{perfil['cargo']}\n{perfil['correo']}")
    st.caption(
        f"Prefijo: {perfil['prefijo']}. "
        "El número se asigna al registrar el informe."
    )

datos = {
    "modelo": modelo,
    "fecha": fecha.strftime("%d-%m-%Y"),
    "vin": vin,
    "patente": patente,
    "sistema": sistema,
    "motivo": motivo,
    "entidades": entidades,
    "analisis": analisis,
    "responsable": responsable,
    "cargo": perfil["cargo"] if perfil else "",
    "correo": perfil["correo"] if perfil else "",
}

firma = sha256(
    repr((
        datos,
        [(sha256(imagen).hexdigest(), texto) for imagen, texto in fotos],
    )).encode()
).hexdigest()

confirmado = st.checkbox(
    "Revisé los antecedentes técnicos y las descripciones de las fotos."
)

if st.button("Generar Word y registrar correlativo", type="primary"):
    st.session_state.pop("informe", None)

    if modelo not in MODELOS:
        errores.append("Selecciona un modelo Foton.")

    if perfil is None:
        errores.append("Selecciona al responsable.")

    if not all(x.strip() for x in (sistema, motivo, analisis)):
        errores.append("Completa los campos obligatorios (*).")

    if not vin and not patente:
        errores.append("Ingresa al menos VIN o patente.")

    if vin and not re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", vin):
        errores.append(
            "El VIN debe tener 17 caracteres y no contener I, O ni Q."
        )

    if any(not texto.strip() for _, texto in fotos):
        errores.append("Agrega una descripción a cada fotografía.")

    if not confirmado:
        errores.append("Confirma la revisión de los antecedentes.")

    if errores:
        for error in errores:
            st.error(error)
    else:
        try:
            resultado, nuevo = guardar_informe(
                datos,
                fotos,
                perfil["prefijo"],
                firma,
                generar_informe,
            )
            st.session_state.informe = (
                firma,
                resultado["documento"],
                resultado["nombre_archivo"],
                resultado["numero"],
            )
            st.session_state.pop("respaldo", None)

            if not nuevo:
                st.info(
                    "Este contenido ya tiene correlativo. "
                    "Se generó el Word con el mismo número."
                )
        except Exception as exc:
            st.error(f"No se pudo registrar el informe: {exc}")

if "informe" in st.session_state:
    firma_guardada, documento, nombre, numero = st.session_state.informe

    if firma_guardada == firma and confirmado and not errores:
        st.success(
            f"Informe {numero} registrado. "
            "Descarga el Word antes de cerrar la sesión."
        )
        st.download_button(
            "Descargar Word",
            documento,
            nombre,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )
    else:
        st.warning(
            "Los datos cambiaron. Genera nuevamente el informe."
        )

st.divider()
st.subheader("Historial de informes")
st.caption(
    "Google Sheets · Datos y textos. "
    "Las eliminaciones borran la fila de la planilla."
)

try:
    titulo, identificador = estado_conexion()
    st.success(f"Conectado a Google Sheets: {titulo}")
    st.caption(f"ID de la planilla: …{identificador[-6:]}")

    mostrar_eliminados = st.checkbox(
        "Mostrar también registros marcados como eliminados "
        "por la versión anterior"
    )

    registros = listar_informes(
        incluir_eliminados=mostrar_eliminados
    )

except Exception as exc:
    st.error(f"No se pudo cargar Google Sheets: {exc}")
    st.stop()

st.caption(f"Registros encontrados: {len(registros)}")

filtro = st.text_input(
    "Buscar por número, patente, VIN, modelo o responsable"
)

visibles = [
    registro
    for registro in registros
    if not filtro
    or filtro.casefold() in " ".join(
        str(valor) for valor in registro.values()
    ).casefold()
]

columnas_visibles = [
    "numero", "fecha", "modelo", "vin",
    "patente", "sistema", "responsable", "estado",
]

if visibles:
    st.dataframe(
        [
            {campo: registro.get(campo, "") for campo in columnas_visibles}
            for registro in visibles
        ],
        use_container_width=True,
        hide_index=True,
    )
else:
    st.info("No hay registros para mostrar.")

if registros:
    salida = StringIO()
    escritor = csv.DictWriter(
        salida,
        fieldnames=list(registros[0]),
    )
    escritor.writeheader()

    for registro in registros:
        fila_segura = {}
        for campo, valor in registro.items():
            texto = str(valor)
            if texto.lstrip().startswith(("=", "+", "-", "@")):
                texto = "'" + texto
            fila_segura[campo] = texto
        escritor.writerow(fila_segura)

    st.download_button(
        "Exportar historial CSV",
        salida.getvalue().encode("utf-8-sig"),
        "historial_foton.csv",
        "text/csv",
    )

if visibles:
    st.subheader("Consultar o corregir un informe")

    numeros = list(dict.fromkeys(
        registro["numero"] for registro in visibles
    ))

    elegido = st.selectbox(
        "Seleccionar número de informe",
        ["Selecciona un informe"] + numeros,
    )

    if elegido != "Selecciona un informe":
        registro = next(
            r for r in visibles if r["numero"] == elegido
        )

        st.caption(
            f"Responsable: {registro['responsable']} · "
            f"Registrado: {registro['creado_en']}"
        )

        eliminado_anterior = (
            registro["estado"].strip().lower() == "eliminado"
        )

        if eliminado_anterior:
            st.info(
                "Este registro fue marcado como eliminado por "
                "la versión anterior. Puedes borrarlo definitivamente abajo."
            )
        else:
            with st.form(f"editar_{elegido}"):
                cambios = {}

                for campo in [
                    "fecha", "modelo", "vin", "patente",
                    "sistema", "motivo", "entidades", "analisis",
                ]:
                    cambios[campo] = st.text_area(
                        campo.capitalize(),
                        value=registro.get(campo, ""),
                        height=(
                            120 if campo in ("motivo", "entidades", "analisis")
                            else 68
                        ),
                    )

                guardar = st.form_submit_button("Guardar correcciones")

            if guardar:
                try:
                    actualizar_informe(elegido, cambios)
                except Exception as exc:
                    st.error(f"No se pudieron guardar los cambios: {exc}")
                else:
                    # Invalida descargas que correspondan a la versión anterior.
                    actual = st.session_state.get("informe")
                    if actual and actual[3] == elegido:
                        st.session_state.pop("informe", None)

                    st.session_state.pop("respaldo", None)
                    st.session_state.aviso = (
                        f"Textos de {elegido} actualizados en Google Sheets. "
                        "Los Word descargados anteriormente no se modifican."
                    )
                    st.rerun()

        with st.expander("Eliminar informe de Google Sheets"):
            st.warning(
                "Se eliminará definitivamente la fila y sus textos "
                "de la pestaña Informes. El correlativo no se reutilizará. "
                "Los Word que ya descargaste no se borran."
            )

            confirmar_borrado = st.checkbox(
                f"Confirmo que deseo borrar definitivamente {elegido}",
                key=f"confirmar_borrado_{elegido}",
            )

            if st.button(
                "Eliminar definitivamente",
                disabled=not confirmar_borrado,
                key=f"borrar_{elegido}",
            ):
                try:
                    eliminar_informe(elegido)
                except Exception as exc:
                    st.error(f"No se pudo completar la eliminación: {exc}")
                else:
                    actual = st.session_state.get("informe")
                    if actual and actual[3] == elegido:
                        st.session_state.pop("informe", None)

                    st.session_state.pop("respaldo", None)
                    st.session_state.aviso = (
                        f"Registro {elegido} eliminado de Google Sheets."
                    )
                    st.rerun()

if st.button("Preparar respaldo de textos CSV"):
    try:
        st.session_state.respaldo = respaldo_base()
    except Exception as exc:
        st.error(f"No se pudo preparar el respaldo: {exc}")

if "respaldo" in st.session_state:
    st.download_button(
        "Descargar respaldo de textos",
        st.session_state.respaldo,
        "respaldo_informes_foton.csv",
        "text/csv",
    )