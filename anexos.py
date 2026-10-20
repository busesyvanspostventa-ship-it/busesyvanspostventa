"""Validación y conversión local de anexos; no sube archivos a servicios externos."""
import fitz

MAX_PDFS = 3
MAX_PAGINAS = 20
MAX_BYTES = 20 * 1024 * 1024


def validar_anexos(anexos):
    if len(anexos) > MAX_PDFS:
        raise ValueError('Adjunta como máximo 3 PDF.')
    total = 0
    for nombre, contenido in anexos:
        if len(contenido) > MAX_BYTES:
            raise ValueError(f'{nombre}: supera 20 MB.')
        try:
            with fitz.open(stream=contenido, filetype='pdf') as pdf:
                if pdf.needs_pass:
                    raise ValueError(f'{nombre}: utiliza un PDF sin contraseña.')
                if not pdf.page_count:
                    raise ValueError(f'{nombre}: el PDF no tiene páginas.')
                total += pdf.page_count
        except ValueError:
            raise
        except Exception as exc:
            raise ValueError(f'{nombre}: no se pudo leer como PDF.') from exc
        if total > MAX_PAGINAS:
            raise ValueError('Los anexos pueden sumar como máximo 20 páginas.')
    return total


def paginas_pdf(contenido):
    with fitz.open(stream=contenido, filetype='pdf') as pdf:
        for indice, pagina in enumerate(pdf):
            rect = pagina.rect
            escala = min(2.0, 1800 / max(rect.width, rect.height))
            pix = pagina.get_pixmap(matrix=fitz.Matrix(escala, escala), alpha=False)
            yield indice + 1, pdf.page_count, pix.tobytes('png'), pix.width, pix.height
