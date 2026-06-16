"""
utils.py
========
Funciones de utilidad compartidas entre todos los módulos del proyecto.
"""


def split_list(lista: list, n: int):
    """
    Divide `lista` en `n` sublistas de tamaño lo más uniforme posible.

    Args:
        lista: Lista de elementos a dividir.
        n: Número de partes.

    Yields:
        Sublistas del tamaño correspondiente.
    """
    k, m = divmod(len(lista), n)
    for i in range(n):
        yield lista[i * k + min(i, m) : (i + 1) * k + min(i + 1, m)]


def print_banner_ascii() -> None:
    """Imprime el banner ASCII de bienvenida del proyecto."""
    banner = """
    ██╗███╗   ██╗███████╗ ██████╗ ██████╗ ███╗   ███╗███████╗███████╗
    ██║████╗  ██║██╔════╝██╔═══██╗██╔══██╗████╗ ████║██╔════╝██╔════╝
    ██║██╔██╗ ██║█████╗  ██║   ██║██████╔╝██╔████╔██║█████╗  ███████╗
    ██║██║╚██╗██║██╔══╝  ██║   ██║██╔══██╗██║╚██╔╝██║██╔══╝  ╚════██║
    ██║██║ ╚████║██║     ╚██████╔╝██║  ██║██║ ╚═╝ ██║███████╗███████║
    ╚═╝╚═╝  ╚═══╝╚═╝      ╚═════╝ ╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝╚══════╝
                            v2.0.0 | Generador de Informes
    """
    print(banner)
    print("-" * 65)
