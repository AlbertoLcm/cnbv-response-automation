"""
xml_parser.py
=============
Funciones para parsear y manipular archivos XML del formato CNBV.
"""

import xml.etree.ElementTree as ET


# Namespace estándar de los XML de CNBV
_NS = {"ns": "http://www.cnbv.gob.mx"}


def generar_estructura_vacia_xml() -> dict:
    """
    Devuelve la estructura base del XML con todos los campos vacíos.
    Se usa como fallback cuando el archivo XML no existe o su estado no es 'OK'.

    Returns:
        Diccionario con la estructura esperada pero sin datos.
    """
    return {
        "DatosGenerales": {
            "Cnbv_NumeroOficio": "",
            "Cnbv_NumeroExpediente": "",
            "Cnbv_SolicitudSiara": "",
            "Cnbv_Folio": "",
            "Cnbv_OficioYear": "",
            "Cnbv_AreaClave": "",
            "Cnbv_AreaDescripcion": "",
            "Cnbv_FechaPublicacion": "",
            "Cnbv_DiasPlazo": "",
            "AutoridadNombre": "",
            "NombreSolicitante": "",
            "Referencia": "",
            "Referencia1": "",
            "Referencia2": "",
            "TieneAseguramiento": "",
        },
        "SolicitudPartes": [],
        "SolicitudesEspecificas": [],
    }


def parsear_expediente_cnbv(xml_source: str, es_archivo: bool = False) -> dict:
    """
    Parsea un XML del expediente CNBV y devuelve un diccionario estructurado.

    Args:
        xml_source: Ruta al archivo XML (si `es_archivo=True`) o string XML crudo.
        es_archivo:  Si es True, `xml_source` se trata como ruta de archivo en disco.

    Returns:
        Diccionario con DatosGenerales, SolicitudPartes y SolicitudesEspecificas.

    Raises:
        ET.ParseError: Si el XML está malformado.
        FileNotFoundError: Si `es_archivo=True` y la ruta no existe.
    """
    if es_archivo:
        tree = ET.parse(xml_source)
        root = tree.getroot()
    else:
        root = ET.fromstring(xml_source)

    def _texto(elemento, tag: str) -> str:
        """Devuelve el texto limpio del nodo o cadena vacía si no existe."""
        nodo = elemento.find(f"ns:{tag}", _NS)
        if nodo is not None and nodo.text is not None:
            return nodo.text.strip()
        return ""

    datos = {
        "DatosGenerales": {
            "Cnbv_NumeroOficio":       _texto(root, "Cnbv_NumeroOficio"),
            "Cnbv_NumeroExpediente":   _texto(root, "Cnbv_NumeroExpediente"),
            "Cnbv_SolicitudSiara":     _texto(root, "Cnbv_SolicitudSiara"),
            "Cnbv_Folio":              _texto(root, "Cnbv_Folio"),
            "Cnbv_OficioYear":         _texto(root, "Cnbv_OficioYear"),
            "Cnbv_AreaClave":          _texto(root, "Cnbv_AreaClave"),
            "Cnbv_AreaDescripcion":    _texto(root, "Cnbv_AreaDescripcion"),
            "Cnbv_FechaPublicacion":   _texto(root, "Cnbv_FechaPublicacion"),
            "Cnbv_DiasPlazo":          _texto(root, "Cnbv_DiasPlazo"),
            "AutoridadNombre":         _texto(root, "AutoridadNombre"),
            "NombreSolicitante":       _texto(root, "NombreSolicitante"),
            "Referencia":              _texto(root, "Referencia"),
            "Referencia1":             _texto(root, "Referencia1"),
            "Referencia2":             _texto(root, "Referencia2"),
            "TieneAseguramiento":      _texto(root, "TieneAseguramiento"),
        },
        "SolicitudPartes": [],
        "SolicitudesEspecificas": [],
    }

    for parte in root.findall("ns:SolicitudPartes", _NS):
        datos["SolicitudPartes"].append(
            {
                "ParteId":  _texto(parte, "ParteId"),
                "Caracter": _texto(parte, "Caracter"),
                "Persona":  _texto(parte, "Persona"),
                "Paterno":  _texto(parte, "Paterno"),
                "Materno":  _texto(parte, "Materno"),
                "Nombre":   _texto(parte, "Nombre"),
                "Rfc":      _texto(parte, "Rfc"),
            }
        )

    for sol in root.findall("ns:SolicitudEspecifica", _NS):
        personas: list[dict] = []
        for pers in sol.findall("ns:PersonasSolicitud", _NS):
            personas.append(
                {
                    "PersonaId":       _texto(pers, "PersonaId"),
                    "Caracter":        _texto(pers, "Caracter"),
                    "Persona":         _texto(pers, "Persona"),
                    "Paterno":         _texto(pers, "Paterno"),
                    "Materno":         _texto(pers, "Materno"),
                    "Nombre":          _texto(pers, "Nombre"),
                    "Rfc":             _texto(pers, "Rfc"),
                    "Relacion":        _texto(pers, "Relacion"),
                    "Domicilio":       _texto(pers, "Domicilio"),
                    "Complementarios": _texto(pers, "Complementarios"),
                }
            )

        datos["SolicitudesEspecificas"].append(
            {
                "SolicitudEspecificaId":              _texto(sol, "SolicitudEspecificaId"),
                "InstruccionesCuentasPorConocer":     _texto(sol, "InstruccionesCuentasPorConocer"),
                "PersonasSolicitud":                  personas,
            }
        )

    return datos
