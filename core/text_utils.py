import re
import unicodedata

import pandas as pd

from config.settings import MONTH_ORDER


def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(ch for ch in texto if not unicodedata.combining(ch)).upper()


def chave_coluna(valor):
    texto = limpar_texto(valor)
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return " ".join(texto.split())


def normalizar_mes(valor):
    texto = limpar_texto(valor)
    if not texto:
        return texto
    if texto in MONTH_ORDER:
        return texto
    base = texto.replace(" ", "")
    for mes in MONTH_ORDER:
        if base == mes.replace(" ", "") or texto.startswith(mes[:3]):
            return mes
    return texto
