"""Historial SQLite: datos y textos del informe, sin documentos ni fotografías."""
from contextlib import closing
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
import os
import re
import sqlite3
import tempfile

RUTA_DB = Path(os.environ.get('FOTON_DB_PATH', str(Path(__file__).parent / 'datos' / 'informes_foton.db')))

def conectar(ruta=None):
    ruta = Path(ruta or RUTA_DB)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(str(ruta), timeout=60)
    c.row_factory = sqlite3.Row
    return c

def inicializar(ruta=None):
    with closing(conectar(ruta)) as c, c:
        c.execute('BEGIN IMMEDIATE')
        c.execute('''CREATE TABLE IF NOT EXISTS informes (
            id INTEGER PRIMARY KEY, prefijo TEXT NOT NULL, secuencia INTEGER NOT NULL,
            numero TEXT NOT NULL UNIQUE, huella TEXT NOT NULL UNIQUE,
            creado_en TEXT NOT NULL, fecha TEXT NOT NULL, modelo TEXT NOT NULL,
            vin TEXT NOT NULL, patente TEXT NOT NULL, sistema TEXT NOT NULL,
            responsable TEXT NOT NULL, correo TEXT NOT NULL,
            UNIQUE(prefijo, secuencia))''')
        existentes = {fila[1] for fila in c.execute('PRAGMA table_info(informes)')}
        for campo in ('motivo', 'entidades', 'analisis', 'cargo'):
            if campo not in existentes:
                c.execute(f"ALTER TABLE informes ADD COLUMN {campo} TEXT NOT NULL DEFAULT ''")

def guardar_informe(datos, fotos, prefijo, huella, generador, ruta=None):
    """Reintentar el mismo contenido devuelve el informe existente sin consumir número."""
    if not re.fullmatch(r'[A-Z]{2}', prefijo):
        raise ValueError('Prefijo inválido')
    inicializar(ruta)
    with closing(conectar(ruta)) as c:
        try:
            c.execute('BEGIN IMMEDIATE')
            existente = c.execute('SELECT numero FROM informes WHERE huella=?', (huella,)).fetchone()
            if existente:
                documento = generador(dict(datos, numero_informe=existente['numero']), fotos)
                c.commit()
                return dict(numero=existente['numero'], documento=documento,
                            nombre_archivo=f"Informe_Foton_{existente['numero']}.docx"), False
            siguiente = c.execute('SELECT COALESCE(MAX(secuencia),0)+1 FROM informes WHERE prefijo=?', (prefijo,)).fetchone()[0]
            numero = f'{prefijo}-{siguiente:04d}'
            datos = dict(datos, numero_informe=numero)
            documento = generador(datos, fotos)
            if not documento:
                raise ValueError('Documento vacío')
            identificador = re.sub(r'[^A-Za-z0-9_-]', '', datos.get('patente') or datos.get('vin') or 'unidad')
            nombre = f'Informe_Foton_{numero}_{identificador}.docx'
            creado = datetime.now(ZoneInfo('America/Santiago')).isoformat(timespec='seconds')
            c.execute('''INSERT INTO informes
                (prefijo,secuencia,numero,huella,creado_en,fecha,modelo,vin,patente,sistema,responsable,correo,motivo,entidades,analisis,cargo)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (prefijo,siguiente,numero,huella,creado,datos['fecha'],datos['modelo'],datos['vin'],datos['patente'],datos['sistema'],datos['responsable'],datos['correo'],datos.get('motivo',''),datos.get('entidades',''),datos.get('analisis',''),datos.get('cargo','')))
            c.commit()
            return dict(numero=numero,nombre_archivo=nombre,documento=documento), True
        except Exception:
            c.rollback()
            raise

def listar_informes(ruta=None):
    inicializar(ruta)
    with closing(conectar(ruta)) as c:
        return [dict(r) for r in c.execute('''SELECT numero,fecha,modelo,vin,patente,sistema,responsable,cargo,correo,motivo,entidades,analisis,creado_en
            FROM informes ORDER BY id DESC''')]

def respaldo_base(ruta=None):
    """Copia consistente usando la API backup, incluso con escrituras concurrentes."""
    with tempfile.TemporaryDirectory() as carpeta:
        destino = Path(carpeta) / 'respaldo.db'
        with closing(conectar(ruta)) as origen, closing(sqlite3.connect(destino)) as copia:
            origen.backup(copia)
        return destino.read_bytes()
