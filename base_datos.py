"""Google Sheets: únicamente textos. Usar una sola instancia de la aplicación."""
import csv
import json
import re
import unicodedata
from io import StringIO
from datetime import datetime
from threading import RLock
from zoneinfo import ZoneInfo
import gspread
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

LOCK = RLock()
COLUMNAS_ANTERIORES = ['numero','fecha','modelo','vin','patente','marca','sistema','motivo','entidades','analisis','responsable','cargo','correo','firma','creado_en','actualizado_en','estado']
COLUMNAS = COLUMNAS_ANTERIORES + ['cliente','km','sucursal','tipo_atencion','foto_vin_adjunta','foto_vin_nombre','anexos_pdf','anexos_paginas','usuario_ingreso']
CONTADORES = ['prefijo','ultimo_numero']
EDITABLES = ['fecha','modelo','vin','patente','sistema','motivo','entidades','analisis','cliente','km','sucursal','tipo_atencion']


def _ahora():
    return datetime.now(ZoneInfo('America/Santiago')).strftime('%Y-%m-%d %H:%M:%S')


@st.cache_resource(show_spinner=False)
def _hojas():
    if 'gcp_service_account' in st.secrets:
        credenciales = dict(st.secrets['gcp_service_account'])
    else:
        credenciales = json.loads(st.secrets['gcp']['credentials_json'])
    credenciales['private_key'] = credenciales['private_key'].replace('\\n', '\n')
    cliente = gspread.authorize(Credentials.from_service_account_info(
        credenciales, scopes=['https://www.googleapis.com/auth/spreadsheets']))
    identificador = str(st.secrets['google_sheets']['spreadsheet_id']).strip()
    match = re.search(r'/spreadsheets/d/([A-Za-z0-9_-]+)', identificador)
    if match:
        identificador = match.group(1)
    libro = cliente.open_by_key(identificador)
    hojas = []
    for nombre in ('Informes', 'Correlativos'):
        try:
            hoja = libro.worksheet(nombre)
        except gspread.WorksheetNotFound:
            hoja = libro.add_worksheet(title=nombre, rows=1000, cols=26)
        hojas.append(hoja)
    return tuple(hojas)


def _encabezados(hoja, columnas):
    actuales = hoja.row_values(1)
    # Solo migrar encabezados compatibles, agregando columnas a la derecha.
    if actuales and actuales != columnas[:len(actuales)]:
        raise ValueError(f'Columnas incompatibles en {hoja.title}. No se modificaron sus datos.')
    if actuales == columnas:
        return
    if not actuales and any(any(c for c in fila) for fila in hoja.get_all_values()[1:]):
        raise ValueError(f'{hoja.title}: hay datos sin encabezados; revisa la primera fila.')
    if actuales:
        # No sobrescribir datos antiguos que pudieran existir bajo columnas sin título.
        for fila in hoja.get_all_values()[1:]:
            if any(c for c in fila[len(actuales):]):
                raise ValueError(f'{hoja.title}: hay datos en columnas sin encabezado.')
    if hoja.col_count < len(columnas):
        hoja.add_cols(len(columnas) - hoja.col_count)
    inicio = len(actuales) + 1
    hoja.update(range_name=f'{rowcol_to_a1(1,inicio)}:{rowcol_to_a1(1,len(columnas))}',
                values=[columnas[len(actuales):]], value_input_option='RAW')


def _normalizar_encabezado(valor):
    texto = unicodedata.normalize('NFKD', str(valor).strip().lower())
    return ''.join(c for c in texto if c.isalnum() and not unicodedata.combining(c))


def _preparar_correlativos(hoja):
    cabecera = hoja.row_values(1)
    if not cabecera:
        _encabezados(hoja, CONTADORES)
        return
    normalizados = [_normalizar_encabezado(c) for c in cabecera]
    nombres_prefijo = {'prefijo', 'prefix', 'prefijousuario', 'prefijoresponsable'}
    nombres_numero = {
        'ultimonumero', 'ultimo', 'ultimocorrelativo', 'correlativo',
        'contador', 'numerocorrelativo', 'ultimonumerousado',
    }
    if (len(normalizados) >= 2 and normalizados[0] in nombres_prefijo
            and normalizados[1] in nombres_numero):
        # Solo normaliza A1:B1. No toca registros ni columnas adicionales.
        if cabecera[:2] != CONTADORES:
            hoja.update(range_name='A1:B1', values=[CONTADORES], value_input_option='RAW')
        return
    raise ValueError(
        'No se reconocen los encabezados de Correlativos. '
        f'A1={cabecera[0]!r}; B1={cabecera[1] if len(cabecera) > 1 else "(vacío)"!r}. '
        'Se espera prefijo en A1 y ultimo_numero en B1. '
        'No se modificaron los valores. Comparte los encabezados antes de cambiar su orden.'
    )


@st.cache_resource(ttl=300, show_spinner=False)
def _preparar():
    with LOCK:
        informes, correlativos = _hojas()
        # Revisar Correlativos primero para detectar incompatibilidades antes de migrar Informes.
        _preparar_correlativos(correlativos)
        _encabezados(informes, COLUMNAS)
        return informes, correlativos


def _leer(hoja):
    return [(i, dict(zip(COLUMNAS, fila + [''] * len(COLUMNAS))))
            for i, fila in enumerate(hoja.get_all_values()[1:], 2) if fila and fila[0]]


@st.cache_data(ttl=60, show_spinner=False)
def listar_informes(incluir_eliminados=False):
    hoja, _ = _preparar()
    return [r for _, r in reversed(_leer(hoja))
            if incluir_eliminados or r['estado'].lower() != 'eliminado']


def _contador(prefijo, registros, hoja):
    encontrados = [(i, f) for i, f in enumerate(hoja.get_all_values()[1:], 2)
                   if f and f[0].strip().upper() == prefijo]
    if len(encontrados) > 1:
        raise ValueError(f'Hay contadores duplicados para {prefijo}.')
    fila = encontrados[0][0] if encontrados else None
    valor = 0
    if encontrados:
        try:
            valor = int(encontrados[0][1][1])
        except (ValueError, IndexError) as exc:
            raise ValueError(f'Contador inválido para {prefijo}.') from exc
    for _, registro in registros:
        numero = registro['numero']
        if re.fullmatch(rf'{re.escape(prefijo)}-\d+', numero):
            valor = max(valor, int(numero.rsplit('-', 1)[1]))
    return valor, fila


def _escribir_contador(hoja, prefijo, valor, fila):
    if fila:
        hoja.update(range_name=f'B{fila}', values=[[valor]], value_input_option='RAW')
    else:
        hoja.append_row([prefijo, valor], value_input_option='RAW')


def guardar_informe(datos, fotos, prefijo, firma, generar_informe):
    try:
        with LOCK:
            if not re.fullmatch('[A-Z]{2,4}', prefijo):
                raise ValueError('Prefijo inválido.')
            hoja, contadores = _preparar()
            registros = _leer(hoja)  # Escrituras siempre usan lectura fresca.
            existente = next((r for _, r in registros if r['firma'] == firma and r['estado'].lower() != 'eliminado'), None)
            nuevo = existente is None
            if nuevo:
                ultimo, fila = _contador(prefijo, registros, contadores)
                numero = f'{prefijo}-{ultimo + 1:04d}'
            else:
                numero = existente['numero']
            documento = generar_informe(dict(datos, numero=numero, numero_informe=numero), fotos)
            if hasattr(documento, 'getvalue'):
                documento = documento.getvalue()
            if not isinstance(documento, bytes):
                raise TypeError('El generador debe devolver bytes o BytesIO.')
            if nuevo:
                # Reservar antes de insertar: un fallo puede dejar un salto, ajustable desde la aplicación.
                _escribir_contador(contadores, prefijo, ultimo + 1, fila)
                registro = dict(datos, numero=numero, marca='FOTON', firma=firma,
                                creado_en=_ahora(), actualizado_en=_ahora(), estado='Activo')
                hoja.append_row([str(registro.get(k, '')) for k in COLUMNAS], value_input_option='RAW')
            return {'documento': documento, 'numero': numero,
                    'nombre_archivo': f'Informe_Foton_{numero}.docx'}, nuevo
    finally:
        listar_informes.clear()


def _unico(registros, numero):
    encontrados = [(i, r) for i, r in registros if r['numero'] == numero]
    if len(encontrados) != 1:
        raise ValueError('No se encontró un registro único con ese número.')
    return encontrados[0]


def actualizar_informe(numero, cambios):
    try:
        with LOCK:
            hoja, _ = _preparar()
            indice, actual = _unico(_leer(hoja), numero)
            if actual['estado'].lower() == 'eliminado':
                raise ValueError('No se puede editar un informe eliminado.')
            actual.update({k: str(v).strip() for k, v in cambios.items() if k in EDITABLES})
            if actual['modelo'] not in ('U9','U10','U12','D9','D10'):
                raise ValueError('Modelo no válido.')
            if actual['vin'] and not re.fullmatch('[A-HJ-NPR-Z0-9]{17}', actual['vin']):
                raise ValueError('VIN no válido.')
            if not actual['vin'] and not actual['patente']:
                raise ValueError('Ingresa VIN o patente.')
            if actual['km'] and not actual['km'].isdigit():
                raise ValueError('Km debe ser un número entero no negativo.')
            if actual['tipo_atencion'] not in ('', 'Postventa', 'Preventa'):
                raise ValueError('Selecciona Postventa o Preventa.')
            actual.update(actualizado_en=_ahora(), firma='')
            hoja.update(range_name=f'A{indice}:{rowcol_to_a1(indice,len(COLUMNAS))}',
                        values=[[actual[k] for k in COLUMNAS]], value_input_option='RAW')
    finally:
        listar_informes.clear()


def eliminar_informe(numero):
    """Borra la fila; al borrar el último informe, retrocede el contador."""
    try:
        with LOCK:
            hoja, contadores = _preparar()
            registros = _leer(hoja)
            indice, _ = _unico(registros, numero)
            match = re.fullmatch(r'([A-Z]{2,4})-(\d+)', numero)
            if not match:
                raise ValueError('Correlativo no válido; no se eliminó la fila.')
            prefijo = match.group(1)
            ultimo, fila = _contador(prefijo, registros, contadores)
            # Preservar el máximo antes de eliminar por si la segunda escritura falla.
            _escribir_contador(contadores, prefijo, ultimo, fila)
            if fila is None:
                _, fila = _contador(prefijo, registros, contadores)
            maximo_actual = _maximo(prefijo, registros)
            hoja.delete_rows(indice)
            if int(match.group(2)) == maximo_actual:
                restantes = [(i, r) for i, r in registros if i != indice]
                try:
                    _escribir_contador(contadores, prefijo, _maximo(prefijo, restantes), fila)
                except Exception as exc:
                    raise RuntimeError(
                        'La fila fue eliminada, pero el contador no pudo retroceder. '
                        'Usa Ajustar mi correlativo al historial con el usuario del responsable.'
                    ) from exc
    finally:
        listar_informes.clear()


def respaldo_base():
    salida = StringIO()
    escritor = csv.DictWriter(salida, fieldnames=COLUMNAS)
    escritor.writeheader()
    for fila in listar_informes(incluir_eliminados=True):
        escritor.writerow({k: "'" + v if v.lstrip().startswith(('=','+','-','@')) else v for k,v in fila.items()})
    return salida.getvalue().encode('utf-8-sig')


@st.cache_data(ttl=300, show_spinner=False)
def estado_conexion():
    hoja, _ = _preparar()
    return hoja.spreadsheet.title, hoja.spreadsheet.id


def _maximo(prefijo, registros):
    valores = [int(r['numero'].rsplit('-', 1)[1]) for _, r in registros
               if re.fullmatch(rf'{re.escape(prefijo)}-\d+', r['numero'])]
    return max(valores, default=0)


def recalcular_correlativo(prefijo):
    """Repara contadores de eliminaciones anteriores sin renumerar registros."""
    with LOCK:
        if not re.fullmatch('[A-Z]{2,4}', prefijo):
            raise ValueError('Prefijo inválido.')
        hoja, contadores = _preparar()
        registros = _leer(hoja)
        _, fila = _contador(prefijo, registros, contadores)
        ultimo = _maximo(prefijo, registros)
        _escribir_contador(contadores, prefijo, ultimo, fila)
        return f'{prefijo}-{ultimo + 1:04d}'
