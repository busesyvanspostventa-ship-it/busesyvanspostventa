from io import BytesIO
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from PIL import Image, ImageOps

CAMPOS = ('modelo','fecha','vin','patente','sistema','motivo','entidades','analisis')

def generar_informe(datos, fotos=(), plantilla=None):
    """fotos: secuencia de (bytes de imagen, descripción). No guarda datos en disco."""
    doc = Document(plantilla or Path(__file__).with_name('plantilla_foton.docx'))
    valores = {k: str(datos.get(k, '')).strip() or 'No informado' for k in CAMPOS}
    parrafos = list(doc.paragraphs)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                parrafos.extend(celda.paragraphs)
    for p in parrafos:
        texto = p.text
        for k,v in valores.items():
            texto = texto.replace('{{'+k+'}}', v)
        if texto != p.text:
            # Conserva propiedades del primer run y del párrafo.
            if p.runs:
                p.runs[0].text = texto
                for r in p.runs[1:]: r.text = ''
            else: p.add_run(texto)
    ancla = next(p for p in doc.paragraphs if '{{fotos}}' in p.text)
    def insertar():
        nodo = OxmlElement('w:p')
        ancla._p.addprevious(nodo)
        p = Paragraph(nodo, ancla._parent)
        p.paragraph_format.space_after = Pt(6)
        return p
    if not fotos:
        insertar().add_run('Sin evidencia fotográfica adjunta.')
    for i,(contenido, descripcion) in enumerate(fotos,1):
        with Image.open(BytesIO(contenido)) as im:
            im = ImageOps.exif_transpose(im).convert('RGB')
            im.thumbnail((1800,1800))
            w,h = im.size
            buffer = BytesIO()
            im.save(buffer, format='JPEG', quality=88)
        buffer.seek(0)
        p = insertar()
        p.alignment = 1
        p.paragraph_format.keep_with_next = True
        # Acota ambos ejes; conserva proporciones en fotografías verticales.
        ancho = min(14.0, 8.5*w/h)
        p.add_run().add_picture(buffer, width=Cm(ancho))
        p = insertar()
        p.paragraph_format.keep_together = True
        p.add_run(f'Fotografía {i}. ').bold = True
        p.add_run(descripcion.strip() or 'Sin descripción.')
    ancla._p.getparent().remove(ancla._p)
    # El número se muestra bajo el título; la plantilla original sigue siendo compatible.
    numero = str(datos.get('numero_informe', '')).strip()
    if numero:
        nodo = OxmlElement('w:p')
        doc.paragraphs[0]._p.addnext(nodo)
        p = Paragraph(nodo, doc._body)
        p.alignment = 2
        p.add_run(f'Informe N.º {numero}').bold = True
    # Responsable al final del cuerpo, después de toda la evidencia.
    if datos.get('responsable'):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.keep_with_next = True
        p.add_run('RESPONSABLE DEL INFORME').bold = True
        for clave in ('responsable', 'cargo', 'correo'):
            p = doc.add_paragraph(str(datos.get(clave, '')).strip())
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.keep_with_next = clave != 'correo'
            for r in p.runs:
                r.font.size = Pt(10)
    salida = BytesIO()
    doc.save(salida)
    return salida.getvalue()
