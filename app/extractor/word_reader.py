"""
word_reader.py
==============
Funciones para extraer texto y datos clave de documentos Word (.docx).
"""

import re
import docx


def extractor_text_document_word(document_path: str) -> str:
    """
    Extrae el texto completo de un documento Word bajando al XML subyacente.

    Esto permite leer Controles de Contenido, Campos de Formulario y Textboxes
    que la API de alto nivel de python-docx ignora.

    Args:
        document_path: Ruta absoluta o relativa al archivo .docx.

    Returns:
        Texto completo del documento. Devuelve cadena vacía si ocurre un error.
    """
    try:
        doc = docx.Document(document_path)
        fragmentos: list[str] = []

        for elemento in doc.element.body.iter():
            if elemento.tag.endswith("}t"):
                if elemento.text:
                    fragmentos.append(elemento.text)
            elif elemento.tag.endswith("}p"):
                fragmentos.append("\n")

        return "".join(fragmentos)

    except Exception as exc:
        print(f"[word_reader] Error al leer '{document_path}': {exc}")
        return ""


def search_autoridad(text: str) -> str:
    """
    Extrae el nombre completo de la autoridad solicitante a partir del texto del oficio.

    Busca el patrón: "emitido por <NOMBRE>, a efecto"

    Args:
        text: Texto completo del documento Word.

    Returns:
        Nombre de la autoridad o 'No encontrado' si el patrón no está presente.
    """
    patron = r"emitido por\s+(.+?),\s*a efecto"
    coincidencia = re.search(patron, text, re.IGNORECASE)

    if coincidencia:
        return coincidencia.group(1).strip()

    return "No encontrado"
