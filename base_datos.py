"""Google Sheets: guarda datos y textos, sin fotos ni archivos Word."""

import csv
import json
import re
from datetime import datetime
from io import StringIO
from threading import RLock
from zoneinfo import ZoneInfo

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials


LOCK = RLock()

COLUMNAS = [
    "numero", "fecha", "modelo", "vin", "patente", "marca",
    "sistema", "motivo", "entidades", "analisis",
    "responsable", "cargo", "correo", "firma",
    "creado_en", "actualizado_en", "estado",
]

CONTADORES = ["prefijo", "ultimo_numero"]
PREFIJOS_VALIDOS = {"AV", "IG", "PG", "DT", "DE", "DF"}


def _ahora():
    return datetime.now(
        ZoneInfo("America/Santiago")
    ).strftime("%Y-%m-%d %H:%M:%S")


def _texto(valor):
    return "" if valor is None else str(valor)


def _spreadsheet_id():
    try:
        valor = str(
            st.secrets["google_sheets"]["spreadsheet_id"]
        ).strip()
    except KeyError:
        raise RuntimeError(
            "Falta [google_sheets].spreadsheet_id en Secrets."
        ) from None

    if valor.startswith(("https://", "http://")):
        coincidencia = re.search(
            r"/spreadsheets/d/([A-Za-z0-9_-]+)", valor
        )
        if not coincidencia:
            raise RuntimeError("El enlace de Google Sheets no es válido.")
        valor = coincidencia.group(1)

    if not re.fullmatch(r"[A-Za-z0-9_-]+", valor):
        raise RuntimeError("El ID de Google Sheets no es válido.")

    return valor


@st.cache_resource(show_spinner=False)
def _hojas():
    if "gcp_service_account" in st.secrets:
        credenciales = dict(st.secrets["gcp_service_account"])

    elif (
        "gcp" in st.secrets
        and "credentials_json" in st.secrets["gcp"]
    ):
        try:
            credenciales = json.loads(
                st.secrets["gcp"]["credentials_json"]
            )
        except (json.JSONDecodeError, TypeError):
            raise RuntimeError(
                "credentials_json no contiene un JSON válido."
            ) from None
    else:
        raise RuntimeError(
            "Faltan credenciales [gcp].credentials_json "
            "o [gcp_service_account] en Secrets."
        )

    if not isinstance(credenciales, dict):
        raise RuntimeError("Las credenciales deben ser un objeto JSON.")

    faltantes = [
        campo
        for campo in ("private_key", "client_email", "token_uri")
        if not credenciales.get(campo)
    ]
    if faltantes:
        raise RuntimeError(
            "Faltan campos en las credenciales: " + ", ".join(faltantes)
        )

    credenciales["private_key"] = str(
        credenciales["private_key"]
    ).replace("\\n", "\n")

    try:
        credenciales_google = Credentials.from_service_account_info(
            credenciales,
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
    except (ValueError, TypeError):
        raise RuntimeError(
            "No se pudieron interpretar las credenciales de Google."
        ) from None

    cliente = gspread.authorize(credenciales_google)

    try:
        libro = cliente.open_by_key(_spreadsheet_id())
    except gspread.exceptions.SpreadsheetNotFound:
        raise RuntimeError(
            "No se encontró la planilla. Revisa su ID y "
            "compártela con la cuenta de servicio."
        ) from None
    except PermissionError:
        raise RuntimeError(
            "La cuenta de servicio no tiene acceso. "
            "Comparte la planilla como Editor con su client_email "
            "y verifica que Google Sheets API esté habilitada."
        ) from None

    try:
        informes = libro.worksheet("Informes")
        correlativos = libro.worksheet("Correlativos")
    except gspread.exceptions.WorksheetNotFound:
        raise RuntimeError(
            "Deben existir las pestañas Informes y Correlativos."
        ) from None

    return informes, correlativos


def _asegurar_encabezados(hoja, esperados):
    valores = hoja.get_all_values()

    if not valores:
        ultima = gspread.utils.rowcol_to_a1(1, len(esperados))
        hoja.update(
            range_name=f"A1:{ultima}",
            values=[esperados],
            value_input_option="RAW",
        )
        return

    encabezados = valores[0]
    if encabezados == esperados:
        return

    tiene_registros = any(
        any(str(celda).strip() for celda in fila)
        for fila in valores[1:]
    )

    if tiene_registros:
        raise RuntimeError(
            f"La pestaña {hoja.title} contiene registros y tiene "
            "encabezados diferentes. No se modificó su contenido. "
            "Se esperan: " + ", ".join(esperados)
        )

    ancho = max(len(encabezados), len(esperados))
    ultima = gspread.utils.rowcol_to_a1(1, ancho)
    hoja.update(
        range_name=f"A1:{ultima}",
        values=[esperados + [""] * (ancho - len(esperados))],
        value_input_option="RAW",
    )


def _preparar():
    with LOCK:
        informes, correlativos = _hojas()
        _asegurar_encabezados(informes, COLUMNAS)

        encabezados = [
            str(x).strip().lower()
            for x in correlativos.row_values(1)
        ]
        if encabezados not in (
            ["prefijo", "ultimo_numero"],
            ["prefijo", "último número"],
        ):
            _asegurar_encabezados(correlativos, CONTADORES)

        return informes, correlativos


def _leer_registros(hoja):
    filas = hoja.get_all_records(
        expected_headers=COLUMNAS,
        numericise_ignore=["all"],
    )
    return [
        {columna: _texto(fila.get(columna, "")) for columna in COLUMNAS}
        for fila in filas
        if _texto(fila.get("numero", "")).strip()
    ]


def listar_informes(incluir_eliminados=False):
    informes, _ = _preparar()
    registros = _leer_registros(informes)

    if not incluir_eliminados:
        registros = [
            r for r in registros
            if r["estado"].strip().lower() != "eliminado"
        ]

    return list(reversed(registros))


def _siguiente(prefijo, registros, correlativos):
    valores = correlativos.get_all_values()
    fila_contador = None
    ultimo = 0

    for indice, fila in enumerate(valores[1:], start=2):
        if not fila or str(fila[0]).strip().upper() != prefijo:
            continue

        if fila_contador is not None:
            raise RuntimeError(
                f"El prefijo {prefijo} está repetido en Correlativos."
            )

        fila_contador = indice
        contenido = str(fila[1]).strip() if len(fila) > 1 else ""

        if contenido:
            try:
                ultimo = int(contenido)
            except ValueError:
                raise RuntimeError(
                    f"El contador de {prefijo} debe ser un entero."
                ) from None

            if ultimo < 0:
                raise RuntimeError("Los contadores no pueden ser negativos.")

    patron = re.compile(rf"{re.escape(prefijo)}-(\d+)")
    for registro in registros:
        coincidencia = patron.fullmatch(registro["numero"])
        if coincidencia:
            ultimo = max(ultimo, int(coincidencia.group(1)))

    return ultimo + 1, fila_contador


def _guardar_contador(correlativos, prefijo, ultimo, fila_contador):
    if fila_contador is not None:
        correlativos.update_cell(fila_contador, 2, ultimo)
    else:
        correlativos.append_row(
            [prefijo, ultimo],
            value_input_option="RAW",
        )


def _crear_documento(datos, fotos, numero, generar_informe):
    datos_doc = dict(datos, numero=numero, numero_informe=numero)
    resultado = generar_informe(datos_doc, fotos)
    nombre_predeterminado = f"Informe_Foton_{numero}.docx"

    if isinstance(resultado, dict):
        documento = resultado.get("documento")
        nombre = resultado.get("nombre_archivo") or nombre_predeterminado
    elif isinstance(resultado, (tuple, list)) and len(resultado) >= 2:
        documento, nombre = resultado[:2]
        nombre = nombre or nombre_predeterminado
    else:
        documento = resultado
        nombre = nombre_predeterminado

    if hasattr(documento, "getvalue"):
        documento = documento.getvalue()
    elif hasattr(documento, "read"):
        if hasattr(documento, "seek"):
            documento.seek(0)
        documento = documento.read()

    if not isinstance(documento, bytes) or not documento:
        raise TypeError("El generador no devolvió un Word válido.")

    return documento, nombre


def guardar_informe(datos, fotos, prefijo, firma, generar_informe):
    """Genera el Word en memoria y guarda únicamente datos y textos."""
    prefijo = str(prefijo).strip().upper()

    if prefijo not in PREFIJOS_VALIDOS:
        raise ValueError("El prefijo del responsable no es válido.")
    if not firma:
        raise ValueError("Falta la identificación del contenido.")

    with LOCK:
        informes, correlativos = _preparar()
        registros = _leer_registros(informes)

        existente = next(
            (
                r for r in registros
                if r["firma"] == firma
                and r["estado"].strip().lower() != "eliminado"
            ),
            None,
        )

        if existente:
            numero = existente["numero"]
            documento, nombre = _crear_documento(
                datos, fotos, numero, generar_informe
            )
            return {
                "documento": documento,
                "nombre_archivo": nombre,
                "numero": numero,
            }, False

        siguiente, fila_contador = _siguiente(
            prefijo, registros, correlativos
        )
        numero = f"{prefijo}-{siguiente:04d}"

        documento, nombre = _crear_documento(
            datos, fotos, numero, generar_informe
        )

        if any(
            r["numero"] == numero
            for r in _leer_registros(informes)
        ):
            raise RuntimeError(
                "Ese correlativo acaba de ser utilizado. Reintenta."
            )

        momento = _ahora()
        registro = dict(
            datos,
            numero=numero,
            marca="FOTON",
            firma=firma,
            creado_en=momento,
            actualizado_en=momento,
            estado="Activo",
        )

        informes.append_row(
            [_texto(registro.get(c, "")) for c in COLUMNAS],
            value_input_option="RAW",
        )

        try:
            _guardar_contador(
                correlativos, prefijo, siguiente, fila_contador
            )
        except Exception:
            st.warning(
                f"El informe {numero} quedó registrado, pero no se "
                "actualizó Correlativos. El siguiente cálculo también "
                "revisa el historial."
            )

        return {
            "documento": documento,
            "nombre_archivo": nombre,
            "numero": numero,
        }, True


def actualizar_informe(numero, cambios):
    permitidos = {
        "fecha", "modelo", "vin", "patente",
        "sistema", "motivo", "entidades", "analisis",
    }

    with LOCK:
        hoja, _ = _preparar()
        filas = hoja.get_all_values()

        coincidencias = [
            (indice, fila)
            for indice, fila in enumerate(filas[1:], start=2)
            if fila and fila[0] == numero
        ]

        if not coincidencias:
            raise ValueError("No se encontró el informe.")
        if len(coincidencias) > 1:
            raise ValueError(
                "El número está duplicado. Revisa la hoja antes de editar."
            )

        indice, fila = coincidencias[0]
        actual = dict(zip(COLUMNAS, fila + [""] * len(COLUMNAS)))

        if actual["estado"].strip().lower() == "eliminado":
            raise ValueError("No se puede editar un informe eliminado.")

        actual.update({
            k: _texto(v)
            for k, v in cambios.items()
            if k in permitidos
        })
        actual["actualizado_en"] = _ahora()
        actual["firma"] = ""

        hoja.update(
            range_name=f"A{indice}:Q{indice}",
            values=[[actual[c] for c in COLUMNAS]],
            value_input_option="RAW",
        )


def eliminar_informe(numero):
    """Borra la fila de Informes y conserva el correlativo utilizado."""
    with LOCK:
        informes, correlativos = _preparar()
        numero = str(numero).strip()

        coincidencia = re.fullmatch(r"([A-Z]{2})-(\d+)", numero)
        if not coincidencia:
            raise ValueError("El número de informe no es válido.")

        prefijo = coincidencia.group(1)
        filas = informes.get_all_values()

        coincidencias = [
            indice
            for indice, fila in enumerate(filas[1:], start=2)
            if fila and str(fila[0]).strip() == numero
        ]

        if not coincidencias:
            raise ValueError("No se encontró el informe.")
        if len(coincidencias) > 1:
            raise ValueError(
                "Ese número está duplicado. Revisa la hoja antes de borrar."
            )

        # Primero asegura el contador para no reutilizar números borrados.
        registros = _leer_registros(informes)
        siguiente, fila_contador = _siguiente(
            prefijo, registros, correlativos
        )

        _guardar_contador(
            correlativos,
            prefijo,
            siguiente - 1,
            fila_contador,
        )

        # Solo se elimina si fue posible conservar el contador.
        informes.delete_rows(coincidencias[0])


def respaldo_base():
    """CSV de datos y textos, incluidos registros eliminados antiguos."""
    salida = StringIO()
    escritor = csv.DictWriter(salida, fieldnames=COLUMNAS)
    escritor.writeheader()

    for fila in listar_informes(incluir_eliminados=True):
        segura = {}
        for campo, valor in fila.items():
            texto = _texto(valor)
            if texto.lstrip().startswith(("=", "+", "-", "@")):
                texto = "'" + texto
            segura[campo] = texto
        escritor.writerow(segura)

    return salida.getvalue().encode("utf-8-sig")


def estado_conexion():
    informes, _ = _preparar()
    return informes.spreadsheet.title, _spreadsheet_id()

# ==========================================================
# REDUCIR CONSULTAS A GOOGLE SHEETS
# Este bloque debe quedar al final de base_datos.py.
# ==========================================================

from functools import wraps


# Revisa los encabezados como máximo cada cinco minutos,
# en lugar de hacerlo con cada cambio del formulario.
_preparar = st.cache_resource(
    ttl=300,
    show_spinner=False,
)(_preparar)


# Reutiliza el historial durante 60 segundos.
listar_informes = st.cache_data(
    ttl=60,
    show_spinner=False,
)(listar_informes)


# Reutiliza los datos de conexión durante cinco minutos.
estado_conexion = st.cache_data(
    ttl=300,
    show_spinner=False,
)(estado_conexion)


def _actualizar_historial_despues(funcion):
    """Invalida el historial después de una operación de escritura."""
    @wraps(funcion)
    def ejecutar(*args, **kwargs):
        try:
            return funcion(*args, **kwargs)
        finally:
            # También se limpia si hubo un error parcial de escritura.
            listar_informes.clear()

    return ejecutar


guardar_informe = _actualizar_historial_despues(
    guardar_informe
)

actualizar_informe = _actualizar_historial_despues(
    actualizar_informe
)

eliminar_informe = _actualizar_historial_despues(
    eliminar_informe
)