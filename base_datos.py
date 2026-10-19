"""Base de datos de informes Foton en Google Sheets.

Guarda datos y textos.
No guarda fotografías ni archivos Word.
"""

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


# Protege operaciones dentro de una misma instancia de la aplicación.
LOCK = RLock()

COLUMNAS = [
    "numero",
    "fecha",
    "modelo",
    "vin",
    "patente",
    "marca",
    "sistema",
    "motivo",
    "entidades",
    "analisis",
    "responsable",
    "cargo",
    "correo",
    "firma",
    "creado_en",
    "actualizado_en",
    "estado",
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
    """Acepta el ID o un enlace completo de Google Sheets."""
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
            r"/spreadsheets/d/([A-Za-z0-9_-]+)",
            valor,
        )

        if not coincidencia:
            raise RuntimeError(
                "El enlace configurado no corresponde a Google Sheets."
            )

        valor = coincidencia.group(1)

    if not re.fullmatch(r"[A-Za-z0-9_-]+", valor):
        raise RuntimeError(
            "El ID de Google Sheets está vacío o contiene "
            "caracteres incorrectos."
        )

    return valor


@st.cache_resource(show_spinner=False)
def _hojas():
    """Conecta con las pestañas Informes y Correlativos."""
    try:
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
                    "credentials_json no contiene un JSON válido. "
                    "Copia el contenido del archivo JSON de la "
                    "cuenta de servicio dentro de las comillas triples."
                ) from None

        else:
            raise RuntimeError(
                "Faltan credenciales en Secrets. Configura "
                "[gcp].credentials_json o [gcp_service_account]."
            )

    except FileNotFoundError:
        raise RuntimeError(
            "No se encontró la configuración de Secrets."
        ) from None

    if not isinstance(credenciales, dict):
        raise RuntimeError(
            "Las credenciales deben ser un objeto JSON."
        )

    campos_requeridos = (
        "private_key",
        "client_email",
        "token_uri",
    )

    faltantes = [
        campo
        for campo in campos_requeridos
        if not credenciales.get(campo)
    ]

    if faltantes:
        raise RuntimeError(
            "Faltan campos en las credenciales: "
            + ", ".join(faltantes)
        )

    credenciales["private_key"] = str(
        credenciales["private_key"]
    ).replace("\\n", "\n")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
    ]

    try:
        credenciales_google = Credentials.from_service_account_info(
            credenciales,
            scopes=scopes,
        )
    except (ValueError, TypeError):
        raise RuntimeError(
            "No se pudieron interpretar las credenciales. "
            "Revisa el JSON y el formato de la clave privada."
        ) from None

    cliente = gspread.authorize(credenciales_google)
    identificador = _spreadsheet_id()

    try:
        libro = cliente.open_by_key(identificador)

    except gspread.exceptions.SpreadsheetNotFound:
        raise RuntimeError(
            "No se encontró la planilla. Revisa su ID y compártela "
            "con el correo client_email de la cuenta de servicio."
        ) from None

    except PermissionError:
        raise RuntimeError(
            "La cuenta de servicio no tiene acceso a la planilla. "
            "Compártela como Editor con su correo client_email "
            "y verifica que Google Sheets API esté habilitada."
        ) from None

    try:
        informes = libro.worksheet("Informes")
        correlativos = libro.worksheet("Correlativos")

    except gspread.exceptions.WorksheetNotFound:
        raise RuntimeError(
            "La planilla debe tener dos pestañas llamadas "
            "exactamente Informes y Correlativos."
        ) from None

    return informes, correlativos


def _asegurar_encabezados(hoja, esperados, rango):
    """Prepara hojas vacías sin sobrescribir registros existentes."""
    valores = hoja.get_all_values()

    if not valores:
        hoja.update(
            range_name=rango,
            values=[esperados],
            value_input_option="RAW",
        )
        return

    encabezados = valores[0]
    tiene_registros = any(
        any(str(celda).strip() for celda in fila)
        for fila in valores[1:]
    )

    if encabezados == esperados:
        return

    if tiene_registros:
        raise RuntimeError(
            f"La pestaña {hoja.title} contiene registros y sus "
            "encabezados no coinciden con los esperados. "
            "No se modificó su contenido. Encabezados requeridos: "
            + ", ".join(esperados)
        )

    # Solo hay encabezados, sin registros: se normaliza la primera fila.
    ancho = max(len(encabezados), len(esperados))
    ultima_celda = gspread.utils.rowcol_to_a1(1, ancho)

    hoja.update(
        range_name=f"A1:{ultima_celda}",
        values=[
            esperados + [""] * (ancho - len(esperados))
        ],
        value_input_option="RAW",
    )


def _preparar():
    with LOCK:
        informes, correlativos = _hojas()

        _asegurar_encabezados(
            informes,
            COLUMNAS,
            "A1:Q1",
        )

        # Acepta también la cabecera alternativa usada anteriormente.
        encabezados_c = correlativos.row_values(1)
        normalizados = [
            str(valor).strip().lower()
            for valor in encabezados_c
        ]

        variantes_validas = [
            ["prefijo", "ultimo_numero"],
            ["prefijo", "último número"],
        ]

        if normalizados not in variantes_validas:
            _asegurar_encabezados(
                correlativos,
                CONTADORES,
                "A1:B1",
            )

        return informes, correlativos


def _leer_registros(hoja):
    filas = hoja.get_all_records(
        expected_headers=COLUMNAS,
        numericise_ignore=["all"],
    )

    return [
        {
            columna: _texto(fila.get(columna, ""))
            for columna in COLUMNAS
        }
        for fila in filas
        if _texto(fila.get("numero", "")).strip()
    ]


def listar_informes(incluir_eliminados=False):
    informes, _ = _preparar()
    registros = _leer_registros(informes)

    if not incluir_eliminados:
        registros = [
            registro
            for registro in registros
            if registro["estado"].strip().lower() != "eliminado"
        ]

    return list(reversed(registros))


def _siguiente(prefijo, registros, correlativos):
    valores = correlativos.get_all_values()
    fila_contador = None
    ultimo_numero = 0

    for indice, fila in enumerate(valores[1:], start=2):
        if not fila:
            continue

        if str(fila[0]).strip().upper() != prefijo:
            continue

        if fila_contador is not None:
            raise RuntimeError(
                f"El prefijo {prefijo} está repetido en Correlativos. "
                "Debe existir una sola fila por prefijo."
            )

        fila_contador = indice
        contenido = str(fila[1]).strip() if len(fila) > 1 else ""

        if contenido:
            try:
                ultimo_numero = int(contenido)
            except ValueError:
                raise RuntimeError(
                    f"El contador de {prefijo} debe ser un número entero."
                ) from None

            if ultimo_numero < 0:
                raise RuntimeError(
                    f"El contador de {prefijo} no puede ser negativo."
                )

    # Considera también informes eliminados para no reutilizar números.
    patron = re.compile(rf"{re.escape(prefijo)}-(\d+)")

    for registro in registros:
        coincidencia = patron.fullmatch(registro["numero"])

        if coincidencia:
            ultimo_numero = max(
                ultimo_numero,
                int(coincidencia.group(1)),
            )

    return ultimo_numero + 1, fila_contador


def _crear_documento(datos, fotos, numero, generar_informe):
    datos_doc = dict(
        datos,
        numero=numero,
        numero_informe=numero,
    )

    resultado = generar_informe(datos_doc, fotos)
    nombre_predeterminado = f"Informe_Foton_{numero}.docx"

    if isinstance(resultado, dict):
        documento = resultado.get("documento")
        nombre_archivo = (
            resultado.get("nombre_archivo")
            or nombre_predeterminado
        )

    elif isinstance(resultado, (tuple, list)) and len(resultado) >= 2:
        documento, nombre_archivo = resultado[:2]
        nombre_archivo = nombre_archivo or nombre_predeterminado

    else:
        documento = resultado
        nombre_archivo = nombre_predeterminado

    if hasattr(documento, "getvalue"):
        documento = documento.getvalue()

    elif hasattr(documento, "read"):
        if hasattr(documento, "seek"):
            documento.seek(0)
        documento = documento.read()

    if not isinstance(documento, bytes) or not documento:
        raise TypeError(
            "El generador no devolvió un documento Word válido."
        )

    return documento, nombre_archivo


def guardar_informe(datos, fotos, prefijo, firma, generar_informe):
    """Genera Word en memoria y registra únicamente textos."""
    prefijo = str(prefijo).strip().upper()

    if prefijo not in PREFIJOS_VALIDOS:
        raise ValueError("El prefijo del responsable no es válido.")

    if not firma:
        raise ValueError("Falta la identificación del contenido del informe.")

    with LOCK:
        informes, correlativos = _preparar()
        registros = _leer_registros(informes)

        existente = next(
            (
                registro
                for registro in registros
                if registro["firma"] == firma
                and registro["estado"].strip().lower() != "eliminado"
            ),
            None,
        )

        if existente:
            numero = existente["numero"]

            documento, nombre_archivo = _crear_documento(
                datos,
                fotos,
                numero,
                generar_informe,
            )

            return {
                "documento": documento,
                "nombre_archivo": nombre_archivo,
                "numero": numero,
            }, False

        siguiente, fila_contador = _siguiente(
            prefijo,
            registros,
            correlativos,
        )

        numero = f"{prefijo}-{siguiente:04d}"

        documento, nombre_archivo = _crear_documento(
            datos,
            fotos,
            numero,
            generar_informe,
        )

        # Comprobación adicional; no es un bloqueo entre varios servidores.
        registros_actuales = _leer_registros(informes)

        if any(
            registro["numero"] == numero
            for registro in registros_actuales
        ):
            raise RuntimeError(
                "Ese correlativo acaba de ser utilizado. "
                "Vuelve a generar el informe."
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
            [
                _texto(registro.get(columna, ""))
                for columna in COLUMNAS
            ],
            value_input_option="RAW",
        )

        # El informe ya está registrado.
        # Si falla el contador, el siguiente cálculo usa también el historial.
        try:
            if fila_contador is not None:
                correlativos.update_cell(
                    fila_contador,
                    2,
                    siguiente,
                )
            else:
                correlativos.append_row(
                    [prefijo, siguiente],
                    value_input_option="RAW",
                )

        except Exception:
            st.warning(
                f"El informe {numero} quedó registrado, pero no se pudo "
                "actualizar la pestaña Correlativos. El siguiente número "
                "también se calculará desde el historial."
            )

        return {
            "documento": documento,
            "nombre_archivo": nombre_archivo,
            "numero": numero,
        }, True


def actualizar_informe(numero, cambios):
    permitidos = {
        "fecha",
        "modelo",
        "vin",
        "patente",
        "sistema",
        "motivo",
        "entidades",
        "analisis",
    }

    with LOCK:
        hoja, _ = _preparar()
        filas = hoja.get_all_values()

        for indice, fila in enumerate(filas[1:], start=2):
            if not fila or fila[0] != numero:
                continue

            actual = dict(
                zip(COLUMNAS, fila + [""] * len(COLUMNAS))
            )

            if actual["estado"].strip().lower() == "eliminado":
                raise ValueError(
                    "No se puede editar un informe eliminado."
                )

            actual.update(
                {
                    campo: _texto(valor)
                    for campo, valor in cambios.items()
                    if campo in permitidos
                }
            )

            actual["actualizado_en"] = _ahora()
            actual["firma"] = ""

            hoja.update(
                range_name=f"A{indice}:Q{indice}",
                values=[
                    [actual[columna] for columna in COLUMNAS]
                ],
                value_input_option="RAW",
            )

            return

        raise ValueError("No se encontró el informe solicitado.")


def eliminar_informe(numero):
    """Marca el registro como eliminado sin borrar su correlativo."""
    with LOCK:
        hoja, _ = _preparar()
        filas = hoja.get_all_values()

        for indice, fila in enumerate(filas[1:], start=2):
            if fila and fila[0] == numero:
                hoja.update(
                    range_name=f"P{indice}:Q{indice}",
                    values=[[_ahora(), "Eliminado"]],
                    value_input_option="RAW",
                )
                return

        raise ValueError("No se encontró el informe solicitado.")


def respaldo_base():
    """Exporta todos los registros y textos como CSV."""
    salida = StringIO()

    escritor = csv.DictWriter(
        salida,
        fieldnames=COLUMNAS,
    )

    escritor.writeheader()

    for fila in listar_informes(incluir_eliminados=True):
        fila_segura = {}

        for campo, valor in fila.items():
            texto = _texto(valor)

            if texto.lstrip().startswith(("=", "+", "-", "@")):
                texto = "'" + texto

            fila_segura[campo] = texto

        escritor.writerow(fila_segura)

    return salida.getvalue().encode("utf-8-sig")


def estado_conexion():
    """Devuelve el título y el ID de la planilla conectada."""
    informes, _ = _preparar()

    return (
        informes.spreadsheet.title,
        _spreadsheet_id(),
    )