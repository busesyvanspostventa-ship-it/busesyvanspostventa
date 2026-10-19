from io import BytesIO
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from PIL import Image, ImageOps

CAMPOS = ('modelo', 'fecha', 'vin', 'patente', 'sistema', 'motivo',
          'entidades', 'analisis', 'numero_informe')


def generar_informe(datos, fotos=(), plantilla=None):
    """Genera el Word en memoria. Las fotografías no se almacenan."""
    doc = Document(plantilla or Path(__file__).with_name('plantilla_foton.docx'))
    valores = {k: str(datos.get(k, '') or '').strip() or 'No informado' for k in CAMPOS}

    # Reemplaza marcadores de texto del documento y de sus tablas.
    parrafos = list(doc.paragraphs)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                parrafos.extend(celda.paragraphs)

    tiene_numero = any('{{numero_informe}}' in p.text for p in parrafos)
    for p in parrafos:
        texto = p.text
        for k, v in valores.items():
            texto = texto.replace('{{' + k + '}}', v)
        # Evita duplicar responsable en otra posición: siempre va al final.
        for k in ('responsable', 'cargo', 'correo'):
            texto = texto.replace('{{' + k + '}}', '')
        if texto != p.text:
            if p.runs:
                p.runs[0].text = texto
                for r in p.runs[1:]:
                    r.text = ''
            else:
                p.add_run(texto)

    ancla = next((p for p in doc.paragraphs if '{{fotos}}' in p.text), None)
    if ancla is None:
        # Permite trabajar con plantillas sin el marcador de fotografías.
        ancla = doc.add_paragraph('{{fotos}}')

    def insertar_antes_ancla():
        nodo = OxmlElement('w:p')
        ancla._p.addprevious(nodo)
        p = Paragraph(nodo, ancla._parent)
        p.paragraph_format.space_after = Pt(6)
        return p

    if not fotos:
        insertar_antes_ancla().add_run('Sin evidencia fotográfica adjunta.')

    for i, (contenido, descripcion) in enumerate(fotos, 1):
        with Image.open(BytesIO(contenido)) as im:
            im = ImageOps.exif_transpose(im).convert('RGB')
            im.thumbnail((1800, 1800))
            w, h = im.size
            buffer = BytesIO()
            im.save(buffer, format='JPEG', quality=88)
        buffer.seek(0)
        p = insertar_antes_ancla()
        p.alignment = 1
        p.paragraph_format.keep_with_next = True
        ancho = min(14.0, 8.5 * w / max(h, 1))
        p.add_run().add_picture(buffer, width=Cm(ancho))
        p = insertar_antes_ancla()
        p.paragraph_format.keep_together = True
        p.add_run(f'Fotografía {i}. ').bold = True
        p.add_run(str(descripcion).strip() or 'Sin descripción.')

    ancla._p.getparent().remove(ancla._p)

    numero = str(datos.get('numero_informe', '') or '').strip()
    if numero and not tiene_numero:
        nodo = OxmlElement('w:p')
        doc.paragraphs[0]._p.addnext(nodo)
        p = Paragraph(nodo, doc._body)
        p.alignment = 2
        p.add_run(f'Informe N.º {numero}').bold = True

    # Responsable siempre al final, incluso si la plantilla tenía marcadores.
    if str(datos.get('responsable', '') or '').strip():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.keep_with_next = True
        p.add_run('RESPONSABLE DEL INFORME').bold = True
        for clave in ('responsable', 'cargo', 'correo'):
            p = doc.add_paragraph(str(datos.get(clave, '') or '').strip())
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.keep_with_next = clave != 'correo'
            for r in p.runs:
                r.font.size = Pt(10)

    salida = BytesIO()
    doc.save(salida)
    return salida.getvalue()
