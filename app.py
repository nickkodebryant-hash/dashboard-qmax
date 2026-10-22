import io
import hashlib
import logging
import math
import re
import sqlite3
import unicodedata
import zipfile
from difflib import get_close_matches

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from sqlalchemy.exc import SQLAlchemyError
from openpyxl.utils.exceptions import InvalidFileException
from pypdf import PdfReader
from pypdf.errors import PdfReadError

from config.settings import (
    MAX_CERTIFICATE_BATCH_MB,
    MAX_CERTIFICATE_FILES,
    MAX_CERTIFICATE_FILE_MB,
    MAX_CERTIFICATE_PAGES,
    MAX_CERTIFICATE_TEXT_CHARS,
    MONTH_ORDER,
)
from core.calibration_management import (
    render_gestao_calibracao,
)
from core.administration import render_administracao_global
from core.home_navigation import (
    CARDS_CENTRAL,
    DESTINOS_VALIDOS,
    GRUPOS_FERRAMENTAS,
    PAINEIS,
    PAINEIS_POR_NOME,
)
from core.text_utils import chave_coluna, limpar_texto, normalizar_mes
from core.pdf_batch import analisar_lote_certificados
from core.quality_actions_dashboard import (
    exibir_dashboard_operacional,
    exibir_relatorios_acoes,
    exibir_painel_target_gerencia,
    exibir_visao_executiva_acoes,
    ler_target_gerencia_excel,
)
from database.calibration_store import (
    carregar_analise_certificado,
    listar_analises_certificados,
    salvar_analise_certificado,
)
from database.connection import carregar_dataframe, salvar_dataframe

st.set_page_config(page_title="Gestão avançada de indicadores da qualidade", layout="wide", page_icon="🎯")

logger = logging.getLogger(__name__)

CRITERIOS_CALIBRACAO = (
    ("Aquecedor", 10, 10, -10, "± 10% do ponto medido"),
    ("Aquecedor arapuca", 10, 10, -10, "± 10% do ponto medido"),
    ("Aquecedor Toaster Oven", 10, 10, -10, "± 10% do ponto medido"),
    ("Balança", 1, 1, -1, "± 1% do ponto medido"),
    ("Balança Ambulatorial", 5, 5, -5, "± 5% do ponto medido"),
    ("Bip", "Visual/sonoro", "Visual/sonoro", None, "Visual/sonoro"),
    ("Bloco Padrão", 0.002, 0.002, -0.002, "± 0,002mm"),
    ("Cabine Audiométrica", 2, 2, -2, "± 2mm"),
    ("Calibrador de Dosímetro", 3, 3, -3, "± 3db"),
    ("Cálibre", 0.2, 0.2, -0.2, "± 0,2mm"),
    ("Controlador de temperatura", 5, 5, -5, "± 5% do ponto medido ou 5ºC"),
    ("Cronômetro", 10, 10, -10, "± 10s"),
    ("Decibelímetro", 5, 5, -5, "± 5% Ponto medido"),
    ("Dip Solder", 2, 2, -2, "± 2mm"),
    ("Dispositivo Estanqueidade", 10, 10, -10, "± 10% Ponto medido"),
    ("Dispositivo Medição", 0.01, 0.01, -0.01, "± 0,01mm"),
    ("Dispositivo Padrão", 0.1, 0.1, -0.1, "± 0,1mm"),
    ("Dosímetro", 3, 3, -3, "± 3db"),
    ("Durometro HRC", 10, 10, -10, "± 10% do ponto medido"),
    ("Escala Metálica", 1.5, 1.5, -1.5, "± 1,5mm"),
    ("Esfigmomanômetro", 5, 5, -5, "+/- 5mmHg"),
    ("Estufa", 5, 5, -5, "± 5°C"),
    ("Extensiômetro", 1, 1, -1, "± 1 mm"),
    ("Fonte Alta Tensão (Hi pot)", 10, 10, -10, "± 10% do ponto medido"),
    ("HYPOT", 20, 20, -20, "± 20% do ponto medido"),
    ("HYPOT III", 20, 20, -20, "± 20% do ponto medido"),
    ("leak master", 0.33, 0.33, None, "0,33Sccm e KPA 5% do ponto medido"),
    ("Lump Neck (Detector de Falha)", 5, 5, -5, "± 5% Do ponto medido"),
    ("Lupa de medição", 0.1, 0.1, -0.1, "± 0.1mm"),
    ("Luximetro", 5, 5, -5, "± 5% do ponto medido"),
    ("Manômetro", 0.5, 0.5, -0.5, "± 0,5bar,  0,51Kgf, 5% do ponto medido"),
    ("Manômetro (Padrão)", 5, 5, -5, "± 5% do ponto medido"),
    ("Máquina de Tração", 5, 5, -5, "± 5% ponto medido / < 50mm/min."),
    ("Máquina de Tração", 5, 5, -5, "± 5% Final de escala / < 50mm/min."),
    ("Massa Padrão", 1, 1, -1, "± 1% da massa"),
    ("Medidor de Diametro", 5, 5, -5, "± 5% Do ponto medido"),
    ("Medidor de Diametro", 5, 5, -5, "± 5% final de escala"),
    ("Medidor de Gás", 1, 1, -1, "± 1% do ponto medido"),
    ("Medidor de Vazão", 0.33, 0.33, None, "0,33Sccm e KPA 5% do ponto medido"),
    ("Megohmetro", 10, 10, 10, "± 10% do ponto medido"),
    ("Metrador", 5, 5, -5, "± 5% do ponto medido"),
    ("Micrômetro", 0.01, 0.01, -0.01, "± 0,01mm"),
    ("Micrômetro de Face", 0.01, 0.01, -0.01, "± 0,01mm"),
    ("Micrômetro Digital", 0.01, 0.01, -0.01, "± 0,01mm"),
    ("Microscópio", 0.02, 0.02, -0.02, "± 0,02mm"),
    ("Mufla", 10, 10, -10, "± 10°C"),
    ("Multimetro", 10, 10, -10, "± 10%  do ponto medido"),
    ("Ohmimetro", 10, 10, -10, "± 10% do ponto medido"),
    ("Paquímetro", 0.03, 0.03, -0.03, "± 0,03mm"),
    ("Parafusadeira", 10, 10, -10, "± 10 % do Ponto medido."),
    ("Phmetro", 0.15, 0.15, -0.15, "± 0,15pH"),
    ("Pipeta Graduada", 2, 2, -2, "± 2% do ponto medido"),
    ("Ponte Kelvin", 3, 3, -3, "± 3mΩ e 0,3ohms "),
    ("Projetor de Perfil", 0.005, 0.005, -0.005, "± 0,005mm / 0°30' / 5% de X e de Y"),
    ("Relógio comparador (medidor de altura)", 0.015, 0.015, -0.015, "± 0.015mm"),
    ("Revolver p/ Presilha", "Visual", "Visual", None, "Visual"),
    ("Spark Tester", 5, 5, -5, "± 5% do ponto medido"),
    ("Suporte Dinamometro", 1, 1, "<1", "< 1 seg."),
    ("Tensiômetro", 3, 3, -3, "± 3% do ponto medido"),
    ("Termohigrometro Digital", 5, 5, -5, "± 5ºC Do Ponto Medido"),
    ("Termômetro", 5, 5, -5, "± 5ºC Do Ponto Medido"),
    ("Termômetro de Imersão", 5, 5, -5, "± 5ºC Do Ponto Medido"),
    ("Termômetro de Imersão (Padrão)", 5, 5, -5, "± 5ºC Do Ponto Medido"),
    ("Termopar J", 5, 5, -5, "± 5ºC Do Ponto Medido"),
    ("Torquímetro", 10, 10, -10, "± 10% do ponto medido"),
    ("Trena Digital (Padrão)", 1, 1, -1, "± 1mm"),
    ("Trena Metálica", 1.5, 1.5, -1.5, "± 1,5mm"),
    ("Válvula de Segurança", 0.5, 0.5, -0.5, "± 0,5bar ou 0,51Kgf"),
    ("Dinamometro Ambulatorial", 5, 5, -5, "± 5% do ponto medido"),
)

st.markdown(
    """
    <style>
    :root {
        --qmax-navy: #172b43;
        --qmax-navy-light: #203a58;
        --qmax-gold: #e99a2b;
        --qmax-white: #f7f9fc;
    }
    .stApp {
        color: var(--qmax-white);
        background-color: var(--qmax-navy);
        background-image:
            radial-gradient(ellipse 68% 95% at 108% -8%, transparent 63%, rgba(233,154,43,.40) 63.2%, transparent 63.6%),
            radial-gradient(ellipse 70% 98% at 108% -8%, transparent 67%, rgba(233,154,43,.30) 67.2%, transparent 67.5%),
            radial-gradient(ellipse 72% 101% at 108% -8%, transparent 71%, rgba(233,154,43,.22) 71.2%, transparent 71.5%);
        background-attachment: fixed;
    }
    .stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3,
    .stApp [data-testid="stCaptionContainer"], .stApp [data-testid="stMetricLabel"] {
        color: var(--qmax-white);
    }
    section[data-testid="stMain"] .block-container {
        padding-top: 5rem !important;
        padding-bottom: 2rem;
    }
    [data-testid="stSidebar"] {
        background-color: rgba(12, 28, 47, .92);
        border-right: 1px solid rgba(233, 154, 43, .25);
    }
    [data-testid="stFileUploader"] section {
        background-color: rgba(23, 43, 67, .8);
        border: 1px dashed var(--qmax-gold);
        border-radius: 0.8rem;
    }
    [data-testid="stFileUploader"] button {
        color: var(--qmax-white);
        border-color: var(--qmax-gold);
        background-color: rgba(233, 154, 43, .16);
    }
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: rgba(23, 43, 67, .72);
        border-color: rgba(233, 154, 43, .45);
        border-radius: 1rem;
    }
    .welcome-hero {
        min-height: 22rem;
        padding: 2.2rem;
        border: 1px solid rgba(233, 154, 43, .5);
        border-radius: 1rem;
        background:
            radial-gradient(ellipse 80% 150% at 110% -25%, transparent 60%, rgba(233,154,43,.34) 60.2%, transparent 60.6%),
            radial-gradient(ellipse 82% 155% at 110% -25%, transparent 65%, rgba(233,154,43,.24) 65.2%, transparent 65.5%),
            linear-gradient(120deg, rgba(12, 28, 47, .94), rgba(23, 43, 67, .78));
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .welcome-kicker {
        color: var(--qmax-gold);
        font-size: .82rem;
        font-weight: 750;
        letter-spacing: .16em;
        text-transform: uppercase;
    }
    .welcome-title {
        color: var(--qmax-white);
        font-size: clamp(2.1rem, 5vw, 3.7rem);
        font-weight: 800;
        line-height: 1.05;
        margin: .7rem 0;
        max-width: 12em;
    }
    .welcome-copy {
        color: rgba(247, 249, 252, .82);
        font-size: 1.05rem;
        max-width: 36rem;
    }
    .welcome-accent {
        width: 5rem;
        height: .25rem;
        background: var(--qmax-gold);
        border-radius: 999px;
        margin: 1.1rem 0;
    }
    .hub-card {
        min-height: 7.2rem;
        padding: 1rem;
        border: 1px solid rgba(233, 154, 43, .42);
        border-radius: 1rem;
        background: linear-gradient(145deg, rgba(12, 28, 47, .92), rgba(32, 58, 88, .84));
        box-shadow: 0 8px 24px rgba(0, 0, 0, .18);
        display: flex;
        align-items: flex-start;
        gap: .9rem;
        transition: transform 160ms ease, border-color 160ms ease, box-shadow 160ms ease;
    }
    .hub-card:hover {
        transform: translateY(-3px);
        border-color: var(--qmax-gold);
        box-shadow: 0 0 22px rgba(233, 154, 43, .24);
    }
    .hub-icon { font-size: 1.8rem; line-height: 1; flex: 0 0 2.2rem; text-align: center; padding-top: .15rem; }
    .hub-content { flex: 1; min-width: 0; }
    .hub-title { color: var(--qmax-gold); font-size: 1.1rem; font-weight: 800; margin: 0 0 .35rem; }
    .hub-copy { color: rgba(247, 249, 252, .82); line-height: 1.4; }
    .hub-metrics { display: flex; flex-wrap: wrap; gap: .4rem 1rem; margin-top: .65rem; }
    .hub-metric { color: rgba(247, 249, 252, .94); font-size: .84rem; font-weight: 650; }
    @media (max-width: 640px) {
        .hub-card { min-height: 0; padding: .85rem; gap: .6rem; }
        .hub-icon { flex-basis: 1.8rem; font-size: 1.5rem; }
        .hub-title { font-size: 1rem; }
        .hub-copy { font-size: .9rem; }
    }
    div.stButton > button {
        border: 1px solid rgba(233, 154, 43, .65);
        background: rgba(233, 154, 43, .12);
        color: var(--qmax-white);
        font-weight: 700;
        transition: background 160ms ease, box-shadow 160ms ease;
        height: auto;
        min-height: 2.5rem;
        max-width: 100%;
        white-space: normal;
        line-height: 1.25;
    }
    div.stButton > button:hover {
        border-color: var(--qmax-gold);
        background: rgba(233, 154, 43, .24);
        box-shadow: 0 0 14px rgba(233, 154, 43, .22);
        color: var(--qmax-white);
    }
    .header-blue {
        background-color: rgba(12, 28, 47, .88); color: white; padding: 15px; text-align: center;
        font-size: 26px; font-weight: bold; border: 1px solid rgba(233, 154, 43, .5);
        border-radius: 8px; margin-bottom: 20px;
        text-transform: uppercase; letter-spacing: 1px;
    }
    .kpi-card {
        background-color: rgba(12, 28, 47, .88); padding: 20px; border-radius: 15px;
        box-shadow: 0px 8px 22px rgba(0, 0, 0, 0.18); border: 1px solid rgba(233, 154, 43, .28);
        text-align: center; min-height: 140px; display: flex; flex-direction: column; justify-content: center;
    }
    .kpi-title { font-size: 11px; font-weight: bold; margin-bottom: 10px; text-transform: uppercase; }
    .kpi-value { font-size: 36px; font-weight: bold; }
    .title-programado { color: #5dd69a; }
    .val-programado { color: #5dd69a; }
    .title-realizado { color: #82c7ff; }
    .val-realizado { color: #82c7ff; }
    .title-porcentagem { color: #ffc45c; }
    .val-porcentagem { color: #ffc45c; }
    .banner-critico {
        background-color: #fdf2f2; border: 1px solid #f8b4b4; padding: 10px 20px;
        border-radius: 6px; color: #9b1c1c; font-weight: bold; font-size: 13px;
        display: flex; justify-content: space-between; align-items: center; margin-top: 15px;
    }
    .chart-box-header {
        background-color: rgba(12, 28, 47, .92); color: white; font-size: 11px; font-weight: bold;
        text-align: center; padding: 4px; border-radius: 4px 4px 0 0; text-transform: uppercase;
        border: 1px solid rgba(233, 154, 43, .4);
    }
    .meta-box {
        background-color: rgba(12, 28, 47, .88); border: 1px dashed #e99a2b; padding: 8px;
        border-radius: 6px; text-align: center; margin-top: 10px;
    }
    div[data-baseweb="tab-list"] {
        gap: 0.65rem;
        border-bottom: 1px solid rgba(128, 128, 128, 0.25);
        padding: 0.35rem 0 0.7rem 0;
    }
    button[data-baseweb="tab"] {
        height: 2.8rem;
        padding: 0 1.25rem;
        color: var(--qmax-white);
        border: 1px solid rgba(233, 154, 43, 0.35);
        border-radius: 0.7rem;
        background-color: rgba(12, 28, 47, 0.72);
        font-weight: 650;
        transition: background-color 120ms ease, border-color 120ms ease;
    }
    button[data-baseweb="tab"]:hover {
        background-color: rgba(233, 154, 43, 0.14);
        border-color: #e99a2b;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #172b43;
        background-color: #e99a2b;
        border-color: #e99a2b;
    }
    div[data-baseweb="tab-highlight"] {
        display: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def localizar_coluna(df, aliases, fallback_positions=None):
    lookup = {chave_coluna(col): col for col in df.columns}
    for alias in aliases:
        key = chave_coluna(alias)
        if key in lookup:
            return lookup[key]

    for col in df.columns:
        chave = chave_coluna(col)
        for alias in aliases:
            if chave_coluna(alias) in chave or chave in chave_coluna(alias):
                return col

    if fallback_positions:
        for pos in fallback_positions:
            if 0 <= pos < len(df.columns):
                return df.columns[pos]

    first_row = df.iloc[0].fillna("").astype(str) if not df.empty else pd.Series([])
    for idx, valor in enumerate(first_row):
        if chave_coluna(valor) in {chave_coluna(alias) for alias in aliases}:
            return df.columns[idx]

    raise KeyError(f"Coluna necessária não encontrada. Esperado um destes nomes: {aliases}")


def tentar_ler_conteudo_do_arquivo(file_uploader):
    file_uploader.seek(0)
    nome = str(file_uploader.name or "").lower()

    if nome.endswith((".csv", ".txt")):
        conteudo = file_uploader.getvalue()
        try:
            texto = conteudo.decode("utf-8-sig")
        except UnicodeDecodeError:
            texto = conteudo.decode("latin-1")
        linhas = texto.splitlines()
        if not linhas:
            raise ValueError("O arquivo carregado está vazio.")

        linhas_validas = []
        numeros_linha_origem = []
        linhas_invalidas = 0
        header = None
        for numero_linha, linha in enumerate(linhas, start=1):
            if not linha.strip():
                continue
            if header is None:
                header = linha
                linhas_validas.append(linha)
                numeros_linha_origem.append(numero_linha)
                continue
            if linha.startswith(",,,,"):
                continue
            colunas = [c.strip() for c in linha.split(",")]
            if len(colunas) < 2:
                linhas_invalidas += 1
                continue
            linhas_validas.append(linha)
            numeros_linha_origem.append(numero_linha)

        if len(linhas_validas) < 2:
            raise ValueError("Não foi possível localizar a tabela FORMS no arquivo CSV/TXT.")

        df = pd.read_csv(io.StringIO("\n".join(linhas_validas)))
        df.columns = [str(c).strip() for c in df.columns]
        numeros_linhas_dados = numeros_linha_origem[1:]
        if len(numeros_linhas_dados) != len(df):
            raise ValueError(
                "Não foi possível mapear com segurança as linhas de origem do arquivo CSV/TXT."
            )
        df.index = numeros_linhas_dados
        return df, linhas_invalidas

    if nome.endswith(".xls") and not nome.endswith((".xlsx", ".xlsm")):
        raise ValueError(
            "O formato .xls não permite ignorar fórmulas com segurança. "
            "Salve a planilha como .xlsx ou .xlsm e carregue novamente."
        )

    if nome.endswith((".xlsx", ".xlsm")):
        file_uploader.seek(0)
        workbook = openpyxl.load_workbook(
            file_uploader,
            read_only=True,
            data_only=True,
        )
        try:
            planilha_forms = next(
                (sheet for sheet in workbook.sheetnames if chave_coluna(sheet) == "FORMS"),
                None,
            )
            if planilha_forms is None:
                raise ValueError("A aba FORMS não foi encontrada no arquivo Excel.")

            worksheet = workbook[planilha_forms]
            linhas = []
            for linha in worksheet.iter_rows():
                linhas.append([celula.value for celula in linha])
        finally:
            workbook.close()

        header_row = None
        for idx, valores in enumerate(linhas[:30]):
            normalizados = [chave_coluna(valor) for valor in valores]
            tem_nome = any("NOME" in valor for valor in normalizados)
            tem_registro = any("REGISTRO" in valor for valor in normalizados)
            tem_status = any("STATUS" in valor for valor in normalizados)
            tem_mes = any("MES" in valor for valor in normalizados)
            if tem_nome and tem_registro and (tem_status or tem_mes):
                header_row = idx
                break

        if header_row is None:
            raise ValueError("Não foi possível localizar o cabeçalho da aba FORMS.")

        cabecalhos = [
            str(valor).strip() if valor is not None else f"Unnamed: {idx}"
            for idx, valor in enumerate(linhas[header_row])
        ]
        df = pd.DataFrame(linhas[header_row + 1 :], columns=cabecalhos)
        df.index = range(header_row + 2, header_row + 2 + len(df))

        if df.empty:
            raise ValueError("A planilha FORMS está vazia.")
        return df, 0

    raise ValueError("Formato de arquivo não suportado. Envie CSV, TXT ou Excel (.xlsx/.xlsm).")


def parece_criterio_nao_numerico(*valores):
    for valor in valores:
        if valor is None:
            continue
        texto = chave_coluna(valor)
        if any(tipo in texto for tipo in ("VISUAL", "SONOR", "AUDIT", "ACUST")):
            return True
        if re.search(r"\b(?:BIP|SOM)\b", texto):
            return True
    return False


def limite_ausente(valor):
    return (
        valor is None
        or (isinstance(valor, str) and not valor.strip())
        or (not isinstance(valor, str) and pd.isna(valor))
    )


def ler_criterios_calibracao():
    criterios = []
    avisos = []

    def limite_numerico(valor):
        if isinstance(valor, bool) or valor is None:
            return None
        if isinstance(valor, (int, float)):
            return float(valor) if math.isfinite(float(valor)) else None
        texto = str(valor).strip().replace(",", ".")
        try:
            numero = float(texto)
            return numero if math.isfinite(numero) else None
        except ValueError:
            return None

    def limite_comparador_valido(valor):
        if not isinstance(valor, str):
            return False
        return re.fullmatch(
            r"(?:<=|>=|<|>|≤|≥)\s*-?\d+(?:[.,]\d+)?",
            valor.strip(),
        ) is not None

    for numero_linha, registro in enumerate(CRITERIOS_CALIBRACAO, start=3):
        equipamento, criterio, limite_maximo, limite_minimo, observacoes = registro
        avaliacao_visual = parece_criterio_nao_numerico(
            criterio,
            limite_maximo,
            limite_minimo,
            observacoes,
        )
        if avaliacao_visual:
            avisos.append(
                f"Linha {numero_linha} — {equipamento}: critério visual/sonoro; "
                "sugiro avaliação do responsável."
            )
        elif limite_ausente(limite_maximo):
            avisos.append(
                f"Linha {numero_linha} — {equipamento}: máximo (C) não informado; "
                "não será presumido."
            )
        else:
            limites_textuais = [
                nome_limite
                for nome_limite, valor in [
                    ("máximo (C)", limite_maximo),
                    ("mínimo (D)", limite_minimo),
                ]
                if not limite_ausente(valor)
                and limite_numerico(valor) is None
                and not limite_comparador_valido(valor)
            ]
            if limites_textuais:
                avisos.append(
                    f"Linha {numero_linha} — {equipamento}: conteúdo não numérico "
                    f"em {' e '.join(limites_textuais)}; requer revisão."
                )
            else:
                maximo = limite_numerico(limite_maximo)
                minimo = limite_numerico(limite_minimo)
                if maximo is not None and minimo is not None and minimo > maximo:
                    avisos.append(
                        f"Linha {numero_linha} — {equipamento}: o mínimo excede o "
                        "máximo; confirme os limites antes de avaliar."
                    )
                elif maximo is not None and limite_ausente(limite_minimo):
                    avisos.append(
                        f"Linha {numero_linha} — {equipamento}: mínimo (D) não "
                        "informado; será considerado somente o máximo (C)."
                    )

        criterios.append(
            {
                "Linha no Excel": numero_linha,
                "Equipamento": equipamento,
                "Critério": criterio,
                "Máximo (C)": limite_maximo,
                "Mínimo (D)": limite_minimo,
                "Observações": observacoes,
                "Tipo de avaliação": (
                    "Visual/sonoro / revisão humana"
                    if avaliacao_visual
                    else "Limites / confirmar no PDF"
                ),
            }
        )

    return pd.DataFrame(criterios), avisos


def converter_numero_certificado(texto):
    valor = str(texto).strip().replace(" ", "")
    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")
    return float(valor)


def limite_numerico_criterio(valor):
    if valor is None or isinstance(valor, bool) or pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor) if math.isfinite(float(valor)) else None
    try:
        return converter_numero_certificado(valor)
    except ValueError:
        return None


def criterios_para_selecao_manual(criterios):
    opcoes = {}
    for _, criterio in criterios.iterrows():
        if criterio["Tipo de avaliação"] != "Limites / confirmar no PDF":
            continue
        maximo = limite_numerico_criterio(criterio["Máximo (C)"])
        minimo = limite_numerico_criterio(criterio["Mínimo (D)"])
        if maximo is None or (
            not limite_ausente(criterio["Mínimo (D)"]) and minimo is None
        ):
            continue
        if minimo is not None and minimo > maximo:
            continue
        linha = int(criterio["Linha no Excel"])
        minimo_texto = f"{minimo:g}" if minimo is not None else "não definido"
        rotulo = (
            f"{criterio['Equipamento']} — mín. {minimo_texto}, máx. {maximo:g} "
            f"(linha {linha})"
        )
        opcoes[rotulo] = linha
    return opcoes


UNIDADES_FATOR_BASE = {
    "%": ("percentual", 1.0),
    "m": ("comprimento", 1.0),
    "km": ("comprimento", 1000.0),
    "cm": ("comprimento", 0.01),
    "mm": ("comprimento", 0.001),
    "um": ("comprimento", 0.000001),
    "µm": ("comprimento", 0.000001),
    "μm": ("comprimento", 0.000001),
    "in": ("comprimento", 0.0254),
    "inch": ("comprimento", 0.0254),
    "inches": ("comprimento", 0.0254),
    "kg": ("massa", 1.0),
    "g": ("massa", 0.001),
    "mg": ("massa", 0.000001),
    "n": ("força", 1.0),
    "kn": ("força", 1000.0),
    "kgf": ("força", 9.80665),
    "s": ("tempo", 1.0),
    "ms": ("tempo", 0.001),
    "min": ("tempo", 60.0),
    "h": ("tempo", 3600.0),
    "pa": ("pressão", 1.0),
    "kpa": ("pressão", 1000.0),
    "mpa": ("pressão", 1_000_000.0),
    "bar": ("pressão", 100_000.0),
    "mmhg": ("pressão", 133.322387415),
    "db": ("nível sonoro", 1.0),
    "hz": ("frequência", 1.0),
    "khz": ("frequência", 1000.0),
    "mhz": ("frequência", 1_000_000.0),
}

PADRAO_UNIDADE = re.compile(
    r"(?<![^\W\d_])(?<!_)"
    r"(%|inches|inch|mmhg|kgf|µm|μm|um|km|cm|mm|m|kg|mg|g|kn|n|"
    r"mpa|kpa|pa|bar|db|khz|mhz|hz|ms|min|h|s|°c|ºc|°f|ºf|rpm|"
    r"mω|ω|ohms?)"
    r"(?![^\W\d_])(?!_)",
    re.IGNORECASE,
)


def normalizar_unidade(unidade):
    valor = str(unidade or "").strip().casefold().replace("º", "°")
    if valor in {"μm"}:
        return "µm"
    if valor in {"ohm", "ohms", "ω"}:
        return "ω"
    if valor == "mω":
        return "mω"
    return valor


def converter_unidade_valor(valor, unidade_origem, unidade_destino):
    origem = normalizar_unidade(unidade_origem)
    destino = normalizar_unidade(unidade_destino)
    if not origem or not destino:
        return None
    if origem == destino:
        return valor
    fator_origem = UNIDADES_FATOR_BASE.get(origem)
    fator_destino = UNIDADES_FATOR_BASE.get(destino)
    if (
        fator_origem is None
        or fator_destino is None
        or fator_origem[0] != fator_destino[0]
    ):
        return None
    return valor * fator_origem[1] / fator_destino[1]


def extrair_unidade(texto):
    if texto is None:
        return ""
    correspondencia = PADRAO_UNIDADE.search(str(texto))
    return (
        normalizar_unidade(correspondencia.group(1))
        if correspondencia
        else ""
    )


def extrair_unidade_resultados(texto):
    inicio = re.search(
        r"(?im)^\s*(?:Resultados|Results)(?:\s*\([^)]*\))?\s*$",
        texto,
    )
    if inicio is None:
        return ""
    secao = texto[inicio.start() : inicio.start() + 5000]
    unidade_resultados = re.search(
        r"(?im)^\s*(?:Resultados|Results)\s*\(\s*([^()\r\n]+)\s*\)\s*$",
        texto,
    )
    if unidade_resultados:
        return extrair_unidade(unidade_resultados.group(1))
    return extrair_unidade(secao)


def extrair_valores_numericos_linha(linha):
    valores = []
    for token in linha.split():
        if token in {"-", "–", "—"}:
            continue
        try:
            valores.append(converter_numero_certificado(token))
        except ValueError:
            continue
    return valores


def parsear_linha_medicao_certificado(linha, contexto=""):
    valores = linha.split()
    valores_numericos = extrair_valores_numericos_linha(linha)
    cabecalho = chave_coluna(contexto)

    if (
        "VALOR MEDIDO NO CENTRO" in cabecalho
        and "DESVIO" in cabecalho
        and len(valores_numericos) == 6
    ):
        nominal, medido, _, fator_k, incerteza_um, veff = valores_numericos
        unidade_medicao = extrair_unidade(contexto)
        unidade_incerteza = extrair_unidade(
            re.search(r"(?is)\bU\s*\(([^)]+)\)", contexto).group(1)
            if re.search(r"(?is)\bU\s*\(([^)]+)\)", contexto)
            else unidade_medicao
        )
        correcao = medido - nominal
        incerteza = converter_unidade_valor(
            incerteza_um,
            unidade_incerteza,
            unidade_medicao,
        )
        if incerteza is None:
            return None
        return {
            "Valor indicado": medido,
            "Valor de referência": nominal,
            "Correção": correcao,
            "Fator k": fator_k,
            "Incerteza": incerteza,
            "Veff": veff,
            "Unidade da correção": unidade_medicao,
            "Unidade da incerteza": unidade_medicao,
        }

    if (
        "ERRO RELATIVO DE REPETITIVIDADE" in cabecalho
        and "VALOR DA REFERENCIA" in cabecalho
        and len(valores_numericos) in {10, 11}
    ):
        valor_indicado = valores_numericos[1]
        valor_referencia = valores_numericos[3]
        correcao = valor_referencia - valor_indicado
        fator_k = valores_numericos[-2]
        incerteza = valores_numericos[-3]
        unidade = extrair_unidade(contexto)
        return {
            "Valor indicado": valor_indicado,
            "Valor de referência": valor_referencia,
            "Correção": correcao,
            "Fator k": fator_k,
            "Incerteza": incerteza,
            "Veff": float("inf"),
            "Unidade da correção": unidade,
            "Unidade da incerteza": unidade,
        }

    if (
        "VALOR INDICADO APOS AJUSTE" in cabecalho
        and "CORRECAO" in cabecalho
        and len(valores_numericos) in {7, 8}
    ):
        deslocamento = 1 if len(valores_numericos) == 8 else 0
        unidade = extrair_unidade(contexto)
        return {
            "Valor indicado": valores_numericos[deslocamento + 2],
            "Valor de referência": valores_numericos[deslocamento],
            "Correção": valores_numericos[deslocamento + 3],
            "Fator k": valores_numericos[deslocamento + 4],
            "Incerteza": valores_numericos[deslocamento + 5],
            "Veff": valores_numericos[deslocamento + 6],
            "Unidade da correção": unidade,
            "Unidade da incerteza": unidade,
        }

    if (
        "ERRO RELATIVO DE INDICACAO" in cabecalho
        and "VALOR REFERENCIA" in cabecalho
        and "S I" in cabecalho
        and len(valores_numericos) in {6, 7}
    ):
        unidade = extrair_unidade(contexto)
        return {
            "Valor indicado": valores_numericos[2],
            "Valor de referência": valores_numericos[1],
            "Correção": valores_numericos[1] - valores_numericos[2],
            "Fator k": valores_numericos[5],
            "Incerteza": valores_numericos[4],
            "Veff": (
                valores_numericos[6]
                if len(valores_numericos) == 7
                else float("inf")
            ),
            "Unidade da correção": unidade,
            "Unidade da incerteza": unidade,
        }

    if (
        "INSTRUMENTO CORRECAO" in cabecalho
        and len(valores) == 6
        and valores[1].casefold() in {"s", "min", "h"}
        and len(valores_numericos) == 5
    ):
        unidade = extrair_unidade(contexto) or valores[1].casefold()
        escalas_tempo = {"s": 1.0, "min": 60.0, "h": 3600.0}
        referencia = valores_numericos[0] * escalas_tempo[valores[1].casefold()]
        if unidade != "s":
            referencia = valores_numericos[0]
        return {
            "Valor indicado": referencia,
            "Valor de referência": referencia,
            "Correção": valores_numericos[1],
            "Fator k": valores_numericos[2],
            "Incerteza": valores_numericos[3],
            "Veff": valores_numericos[4],
            "Unidade da correção": unidade,
            "Unidade da incerteza": unidade,
        }

    valores_sem_separadores = [
        token for token in valores if token not in {"-", "–", "—"}
    ]
    if len(valores_sem_separadores) != 6 or len(valores_numericos) != 5:
        return None
    try:
        numeros = [converter_numero_certificado(valor) for valor in valores_sem_separadores[:5]]
        veff_token = valores_sem_separadores[5]
        veff = (
            float("inf")
            if chave_coluna(veff_token) in {"INFINITO", "INF"} or veff_token == "∞"
            else converter_numero_certificado(veff_token)
        )
    except ValueError:
        return None
    unidade = extrair_unidade(contexto)
    return {
        "Valor indicado": numeros[0],
        "Valor de referência": numeros[1],
        "Correção": numeros[2],
        "Fator k": numeros[3],
        "Incerteza": numeros[4],
        "Veff": veff,
        "Unidade da correção": unidade,
        "Unidade da incerteza": unidade,
    }


def localizar_criterio_certificado(equipamento, tipo, criterios):
    equipamento_normalizado = chave_coluna(equipamento)
    nomes_criterios = criterios["Equipamento"].fillna("").astype(str)
    indices_exatos = [
        indice
        for indice, nome in nomes_criterios.items()
        if chave_coluna(nome) == equipamento_normalizado
    ]
    if len(indices_exatos) == 1:
        return criterios.loc[indices_exatos[0]], "nome exato"
    if len(indices_exatos) > 1:
        return None, "há mais de um critério com o mesmo nome"

    qualificadores_genericos = {
        "ANALOGICA",
        "ANALOGICO",
        "DIGITAL",
        "ELETRONICA",
        "ELETRONICO",
        "MECANICA",
        "MECANICO",
        "PADRAO",
    }
    indices_candidatos = [
        indice
        for indice, nome in nomes_criterios.items()
        if (
            equipamento_normalizado in chave_coluna(nome)
            or (
                chave_coluna(nome) in equipamento_normalizado
                and set(equipamento_normalizado.split())
                - set(chave_coluna(nome).split())
                <= qualificadores_genericos
            )
        )
    ]
    tipo_normalizado = chave_coluna(tipo)
    if "ANAL" in tipo_normalizado:
        indices_candidatos = [
            indice
            for indice in indices_candidatos
            if "DIGITAL" not in chave_coluna(nomes_criterios.loc[indice])
        ]
    elif "DIGITAL" in tipo_normalizado:
        indices_candidatos = [
            indice
            for indice in indices_candidatos
            if "DIGITAL" in chave_coluna(nomes_criterios.loc[indice])
        ]

    if len(indices_candidatos) == 1:
        return criterios.loc[indices_candidatos[0]], (
            f"objeto '{equipamento}' associado pelo tipo '{tipo}'"
        )
    if indices_candidatos:
        alternativas = ", ".join(
            str(nomes_criterios.loc[indice]) for indice in indices_candidatos
        )
        return None, f"associação ambígua entre: {alternativas}"

    alternativas = get_close_matches(
        equipamento_normalizado,
        [chave_coluna(nome) for nome in nomes_criterios if nome.strip()],
        n=3,
        cutoff=0.45,
    )
    if alternativas:
        nomes_sugeridos = [
            str(nome)
            for nome in nomes_criterios
            if chave_coluna(nome) in alternativas
        ]
        return None, "possíveis critérios: " + ", ".join(nomes_sugeridos)
    return None, "equipamento não encontrado na lista de critérios"


def analisar_certificado_pdf(arquivo, criterios, linha_criterio_manual=None):
    nome_arquivo = str(arquivo.name or "PDF sem nome")
    conteudo = arquivo.getvalue()
    if len(conteudo) > MAX_CERTIFICATE_FILE_MB * 1024 * 1024:
        return {
            "Arquivo": nome_arquivo,
            "Resultado": "Não analisado — arquivo excede o limite individual",
            "Equipamento": "",
            "Tipo": "",
            "Faixa nominal": "",
            "Unidade": "",
            "Critério aplicado": "",
            "Identificação": "",
            "Certificado": "",
            "Medições extraídas": 0,
            "Detalhe": f"Limite individual: {MAX_CERTIFICATE_FILE_MB} MB.",
        }, pd.DataFrame()

    try:
        pdf = PdfReader(io.BytesIO(conteudo), strict=True)
        if pdf.is_encrypted:
            return {
                "Arquivo": nome_arquivo,
                "Resultado": "Não analisado — PDF protegido por senha",
                "Equipamento": "",
                "Tipo": "",
                "Faixa nominal": "",
                "Unidade": "",
                "Critério aplicado": "",
                "Identificação": "",
                "Certificado": "",
                "Medições extraídas": 0,
                "Detalhe": "Remova a proteção ou envie uma cópia autorizada sem senha.",
            }, pd.DataFrame()
        if len(pdf.pages) > MAX_CERTIFICATE_PAGES:
            return {
                "Arquivo": nome_arquivo,
                "Resultado": "Não analisado — PDF excede o limite de páginas",
                "Equipamento": "",
                "Tipo": "",
                "Faixa nominal": "",
                "Unidade": "",
                "Critério aplicado": "",
                "Identificação": "",
                "Certificado": "",
                "Medições extraídas": 0,
                "Detalhe": f"Limite: {MAX_CERTIFICATE_PAGES} páginas por arquivo.",
            }, pd.DataFrame()

        paginas = []
        total_caracteres = 0
        for pagina in pdf.pages:
            texto_pagina = pagina.extract_text() or ""
            total_caracteres += len(texto_pagina)
            if total_caracteres > MAX_CERTIFICATE_TEXT_CHARS:
                return {
                    "Arquivo": nome_arquivo,
                    "Resultado": "Não analisado — extração excedeu o limite de texto",
                    "Equipamento": "",
                    "Tipo": "",
                    "Faixa nominal": "",
                    "Unidade": "",
                    "Critério aplicado": "",
                    "Identificação": "",
                    "Certificado": "",
                    "Medições extraídas": 0,
                    "Detalhe": (
                        f"Limite: {MAX_CERTIFICATE_TEXT_CHARS:,} caracteres extraídos."
                    ),
                }, pd.DataFrame()
            paginas.append(texto_pagina)
    except (PdfReadError, ValueError, OSError, NotImplementedError) as erro:
        return {
            "Arquivo": nome_arquivo,
            "Resultado": "Não analisado — falha na leitura do PDF",
            "Equipamento": "",
            "Tipo": "",
            "Faixa nominal": "",
            "Unidade": "",
            "Critério aplicado": "",
            "Identificação": "",
            "Certificado": "",
            "Medições extraídas": 0,
            "Detalhe": str(erro),
        }, pd.DataFrame()

    texto = "\n".join(paginas)
    if not texto.strip():
        return {
            "Arquivo": nome_arquivo,
            "Resultado": "Revisão necessária — não foi possível extrair texto",
            "Equipamento": "",
            "Tipo": "",
            "Faixa nominal": "",
            "Unidade": "",
            "Critério aplicado": "",
            "Identificação": "",
            "Certificado": "",
            "Medições extraídas": 0,
            "Detalhe": "O PDF pode ser digitalizado como imagem; OCR não está habilitado.",
        }, pd.DataFrame()

    padroes_campos = {
        "Equipamento": (
            r"(?im)^\s*Objeto\s+da\s+calibra[^\r\n:]{0,12}:\s*(.+?)"
            r"(?=\s+(?:Protocolo|Faixa\s+nominal|Valor\s+nominal)\s*:|$)"
        ),
        "Tipo": r"(?i)\bTipo\s*:\s*([^\s\r\n]+)",
        "Faixa nominal": (
            r"(?i)\bFaixa\s+nominal\s*:\s*([^\s\r\n]+"
            r"(?:\s+(?!Protocolo\b)[^\s\r\n]+)?)"
        ),
        "Identificação": r"(?im)^\s*Identifica[^\r\n:]{0,12}:\s*([A-Z]\s*\d[\d./_-]*|[A-Z0-9][A-Z0-9./_-]*)",
        "Certificado": r"(?im)^\s*N[^\d\r\n]{0,5}(\d{4,}-\d+)\s*$",
    }
    campos = {}
    for nome, padrao in padroes_campos.items():
        correspondencia = re.search(padrao, texto)
        campos[nome] = correspondencia.group(1).strip() if correspondencia else ""
    if not campos["Equipamento"]:
        correspondencia = re.search(
            r"(?im)^\s*:\s*(Cron[oô]metro(?:\s+Digital)?)\s*$",
            texto,
        )
        campos["Equipamento"] = (
            correspondencia.group(1).strip() if correspondencia else ""
        )
    campos["Identificação"] = re.sub(r"\s+", "", campos["Identificação"])
    campos["Unidade"] = extrair_unidade_resultados(texto)
    linhas_medicao = []
    inicio_resultados = re.search(
        r"(?im)^\s*(?:Resultados|Results)(?:\s*\([^)]*\))?\s*$",
        texto,
    )
    if inicio_resultados:
        fim_resultados = re.search(
            r"(?im)^\s*(?:"
            r"Resultados\s+finais|"
            r"Last\s+results|"
            r"Valores\s+m.{0,5}ximos\s+obtidos|"
            r"Corre[cç][aã]o\s+m.{0,5}xima\s+na\s+faixa\s+calibrada|"
            r"An.{0,4}lise\s+dos?\s+resultados?|"
            r"Erro\s+relativo\s+de\s+zero|"
            r"Classifica[cç][aã]o\s+conforme\s+norma|"
            r"Item\s+Corre[cç][aã]o\s*\+\s*U|"
            r"P.{0,3}gina\s+\d+\s+de\s+\d+"
            r")\b",
            texto[inicio_resultados.end() :],
        )
        fim_posicao = (
            inicio_resultados.end() + fim_resultados.start()
            if fim_resultados
            else len(texto)
        )
        secao_resultados = texto[inicio_resultados.end() : fim_posicao]
        contexto_medicao = (
            f"Resultados ({campos['Unidade']})\n"
            if campos["Unidade"]
            else ""
        ) + secao_resultados
        for linha in secao_resultados.splitlines():
            medicao = parsear_linha_medicao_certificado(
                linha,
                contexto_medicao,
            )
            if medicao is not None:
                linhas_medicao.append(medicao)
    if not campos["Unidade"] and linhas_medicao:
        campos["Unidade"] = linhas_medicao[0]["Unidade da correção"]

    detalhe = []
    resultado = "Revisão necessária — aprovação/reprovação ainda não determinada"
    criterio_aplicado = ""
    if not campos["Equipamento"]:
        detalhe.append("não foi possível identificar o equipamento no certificado")
    else:
        criterio, forma_associacao = localizar_criterio_certificado(
            campos["Equipamento"],
            campos["Tipo"],
            criterios,
        )
        if criterio is None and linha_criterio_manual is not None:
            selecoes_validas = criterios_para_selecao_manual(criterios)
            if linha_criterio_manual in selecoes_validas.values():
                criterio = criterios.loc[
                    criterios["Linha no Excel"] == linha_criterio_manual
                ].iloc[0]
                forma_associacao = (
                    f"selecionado manualmente pelo avaliador "
                    f"(associação automática: {forma_associacao})"
                )
        if criterio is not None:
            criterio_aplicado = str(criterio["Equipamento"])
            detalhe.append(
                f"critério '{criterio_aplicado}' associado por {forma_associacao}"
            )
            if criterio["Tipo de avaliação"] == "Visual/sonoro / revisão humana":
                resultado = "Revisão humana — critério visual/sonoro"
                detalhe.append(
                    "Critério visual/sonoro; sugiro avaliação do responsável."
                )
            elif criterio["Tipo de avaliação"] != "Limites / confirmar no PDF":
                detalhe.append(
                    "o critério contém fórmula sem valor calculado; requer revisão "
                    "antes da avaliação numérica"
                )
            elif not linhas_medicao:
                detalhe.append(
                    "não há medições numéricas suficientes para aplicar o critério"
                )
            elif not campos["Unidade"]:
                resultado = "Revisão necessária — unidade ausente"
                detalhe.append(
                    "unidade de medição não foi localizada nos resultados do PDF"
                )
            else:
                observacoes_criterio = str(criterio["Observações"] or "")
                observacoes_normalizadas = chave_coluna(observacoes_criterio)
                unidades_observacao = [
                    normalizar_unidade(unidade)
                    for unidade in PADRAO_UNIDADE.findall(observacoes_criterio)
                ]
                percentual = re.search(
                    r"(\d+(?:[.,]\d+)?)\s*%",
                    observacoes_criterio,
                )
                percentual_valor = (
                    float(percentual.group(1).replace(",", "."))
                    if percentual
                    else None
                )
                usar_percentual = (
                    percentual is not None
                    and (
                        "PONTO MEDIDO" in observacoes_normalizadas
                        or "DA MASSA" in observacoes_normalizadas
                        or "FINAL DE ESCALA" in observacoes_normalizadas
                    )
                )
                limite_maximo_criterio = limite_numerico_criterio(
                    criterio["Máximo (C)"]
                )
                limite_minimo_criterio = limite_numerico_criterio(
                    criterio["Mínimo (D)"]
                )
                limites_percentuais_coerentes = (
                    not usar_percentual
                    or (
                        percentual_valor is not None
                        and limite_maximo_criterio is not None
                        and math.isclose(
                            abs(limite_maximo_criterio),
                            percentual_valor,
                        )
                        and (
                            limite_ausente(criterio["Mínimo (D)"])
                            or (
                                limite_minimo_criterio is not None
                                and math.isclose(
                                    abs(limite_minimo_criterio),
                                    percentual_valor,
                                )
                            )
                        )
                    )
                )
                percentual_composto = (
                    percentual is not None
                    and " OU " in f" {observacoes_normalizadas} "
                    and any(unidade not in {"%"} for unidade in unidades_observacao)
                )
                if percentual_composto:
                    resultado = "Revisão necessária — critério percentual composto"
                    detalhe.append(
                        "a observação do critério combina percentual e outro limite; "
                        "a regra precisa ser confirmada pelo responsável"
                    )
                elif usar_percentual and not limites_percentuais_coerentes:
                    resultado = "Revisão necessária — percentuais divergentes"
                    detalhe.append(
                        "o percentual da observação não coincide com os limites "
                        "numéricos do critério"
                    )
                elif percentual is not None and not usar_percentual:
                    detalhe.append(
                        "base da tolerância percentual não está definida para cálculo "
                        "automático"
                    )
                elif not usar_percentual and not extrair_unidade(
                    observacoes_criterio
                ):
                    resultado = "Revisão necessária — unidade ausente no critério"
                    detalhe.append(
                        "a unidade do critério não foi identificada"
                    )
                elif (
                    not usar_percentual
                    and converter_unidade_valor(
                        1.0,
                        campos["Unidade"],
                        extrair_unidade(observacoes_criterio),
                    )
                    is None
                ):
                    resultado = "Revisão necessária — unidades incompatíveis"
                    detalhe.append(
                        f"unidade do PDF ({campos['Unidade']}) "
                        "não é compatível com a unidade do critério "
                        f"({extrair_unidade(observacoes_criterio)})"
                    )
                else:
                    limite_maximo = limite_numerico_criterio(criterio["Máximo (C)"])
                    limite_minimo = limite_numerico_criterio(criterio["Mínimo (D)"])
                    minimo_nao_definido = limite_ausente(criterio["Mínimo (D)"])
                    if limite_maximo is None or (
                        not minimo_nao_definido and limite_minimo is None
                    ):
                        detalhe.append(
                            "um limite obrigatório não está definido numericamente; "
                            "avaliação automática interrompida para revisão"
                        )
                    else:
                        usar_limite_minimo = not minimo_nao_definido
                        fora_dos_limites = []
                        incerteza_indisponivel = []
                        unidade_incompativel = []
                        referencia_indisponivel = []
                        unidade_destino = (
                            campos["Unidade"]
                            if usar_percentual
                            else extrair_unidade(observacoes_criterio)
                        )
                        faixa_escala = None
                        unidade_faixa = extrair_unidade(campos["Faixa nominal"])
                        if usar_percentual and "FINAL DE ESCALA" in observacoes_normalizadas:
                            faixa_match = re.search(
                                r"[-+]?\d+(?:[.,]\d+)?",
                                campos["Faixa nominal"],
                            )
                            if faixa_match and unidade_faixa:
                                faixa_escala = converter_numero_certificado(
                                    faixa_match.group(0)
                                )
                                faixa_escala = converter_unidade_valor(
                                    faixa_escala,
                                    unidade_faixa,
                                    unidade_destino,
                                )
                        for medicao in linhas_medicao:
                            unidade_correcao = (
                                medicao["Unidade da correção"] or campos["Unidade"]
                            )
                            unidade_incerteza = (
                                medicao["Unidade da incerteza"] or campos["Unidade"]
                            )
                            correcao = converter_unidade_valor(
                                medicao["Correção"],
                                unidade_correcao,
                                unidade_destino,
                            )
                            incerteza = converter_unidade_valor(
                                medicao["Incerteza"],
                                unidade_incerteza,
                                unidade_destino,
                            )
                            if correcao is None or incerteza is None:
                                medicao["Avaliação"] = "Revisão necessária"
                                unidade_incompativel.append(medicao["Correção"])
                                continue
                            if not math.isfinite(incerteza) or incerteza < 0:
                                medicao["Avaliação"] = "Revisão necessária"
                                incerteza_indisponivel.append(correcao)
                                continue

                            if usar_percentual:
                                if "FINAL DE ESCALA" in observacoes_normalizadas:
                                    referencia = faixa_escala
                                else:
                                    referencia_bruta = medicao.get(
                                        "Valor de referência"
                                    )
                                    if referencia_bruta is None:
                                        referencia_bruta = medicao.get("Valor indicado")
                                    referencia = (
                                        converter_unidade_valor(
                                            abs(referencia_bruta),
                                            unidade_correcao,
                                            unidade_destino,
                                        )
                                        if referencia_bruta is not None
                                        else None
                                    )
                                if referencia is None:
                                    medicao["Avaliação"] = "Revisão necessária"
                                    referencia_indisponivel.append(correcao)
                                    continue
                                fator_percentual = percentual_valor / 100
                                maximo_medicao = fator_percentual * referencia
                                minimo_medicao = (
                                    limite_minimo / 100 * referencia
                                    if usar_limite_minimo
                                    else None
                                )
                            else:
                                maximo_medicao = converter_unidade_valor(
                                    limite_maximo,
                                    unidade_destino,
                                    campos["Unidade"],
                                )
                                minimo_medicao = (
                                    converter_unidade_valor(
                                        limite_minimo,
                                        unidade_destino,
                                        campos["Unidade"],
                                    )
                                    if usar_limite_minimo
                                    else None
                                )
                                if maximo_medicao is None or (
                                    usar_limite_minimo and minimo_medicao is None
                                ):
                                    medicao["Avaliação"] = "Revisão necessária"
                                    unidade_incompativel.append(correcao)
                                    continue

                            limite_inferior_com_incerteza = correcao - incerteza
                            limite_superior_com_incerteza = correcao + incerteza
                            aprovada = (
                                (
                                    minimo_medicao
                                    <= limite_inferior_com_incerteza
                                    if usar_limite_minimo
                                    else True
                                )
                                and limite_superior_com_incerteza <= maximo_medicao
                            )
                            medicao["Correção"] = correcao
                            medicao["Incerteza"] = incerteza
                            medicao["Correção - U"] = limite_inferior_com_incerteza
                            medicao["Correção + U"] = limite_superior_com_incerteza
                            medicao["Limite mínimo"] = (
                                minimo_medicao if usar_limite_minimo else None
                            )
                            medicao["Limite máximo"] = maximo_medicao
                            medicao["Avaliação"] = (
                                "Aprovada" if aprovada else "Reprovada"
                            )
                            if not aprovada:
                                fora_dos_limites.append(
                                    (
                                        (
                                            usar_limite_minimo
                                            and correcao < minimo_medicao
                                        )
                                        or correcao > maximo_medicao,
                                        correcao,
                                    )
                                )
                        if fora_dos_limites:
                            correcoes_fora = sum(
                                correcao_fora for correcao_fora, _ in fora_dos_limites
                            )
                            resultado = "Reprovado — Correção ± U fora dos limites"
                            detalhe.append(
                                f"{len(fora_dos_limites)} medição(ões) não atendem "
                                "aos limites após considerar U"
                            )
                            if correcoes_fora:
                                detalhe.append(
                                    f"{correcoes_fora} correção(ões) centrais já estão "
                                    "fora dos limites"
                                )
                        elif unidade_incompativel:
                            resultado = "Revisão necessária — unidades incompatíveis"
                            detalhe.append(
                                f"não foi possível converter a unidade de "
                                f"{len(unidade_incompativel)} medição(ões)"
                            )
                        elif referencia_indisponivel:
                            resultado = "Revisão necessária — referência percentual ausente"
                            detalhe.append(
                                f"não foi possível determinar a base percentual de "
                                f"{len(referencia_indisponivel)} medição(ões)"
                            )
                        elif incerteza_indisponivel:
                            resultado = "Revisão necessária — incerteza U indisponível"
                            detalhe.append(
                                f"não foi possível validar U em "
                                f"{len(incerteza_indisponivel)} medição(ões)"
                            )
                        elif usar_limite_minimo:
                            resultado = "Aprovado — Correção ± U dentro dos limites"
                            detalhe.append(
                                "todas as faixas Correção ± U estão dentro dos limites"
                            )
                        else:
                            resultado = "Aprovado — Correção + U abaixo do máximo"
                            detalhe.append(
                                "todas as faixas Correção + U estão até o máximo; "
                                "mínimo não definido"
                            )
                        if usar_percentual:
                            base_percentual = (
                                "final de escala"
                                if "FINAL DE ESCALA" in observacoes_normalizadas
                                else "valor de referência"
                            )
                            detalhe.append(
                                f"tolerância percentual calculada sobre {base_percentual}"
                            )
                        else:
                            detalhe.append(
                                "limites e medições convertidos para a unidade "
                                f"{unidade_destino}"
                            )
        else:
            detalhe.append(forma_associacao)

    if not campos["Identificação"]:
        detalhe.append("identificação do instrumento não foi localizada")
    if not campos["Certificado"]:
        detalhe.append("número do certificado não foi localizado")
    if not linhas_medicao:
        detalhe.append(
            "nenhuma linha de medição compatível foi extraída; conferir manualmente"
        )
    return {
        "Arquivo": nome_arquivo,
        "Resultado": resultado,
        "Equipamento": campos["Equipamento"],
        "Tipo": campos["Tipo"],
        "Faixa nominal": campos["Faixa nominal"],
        "Unidade": campos["Unidade"],
        "Critério aplicado": criterio_aplicado,
        "Identificação": campos["Identificação"],
        "Certificado": campos["Certificado"],
        "Medições extraídas": len(linhas_medicao),
        "Detalhe": "; ".join(detalhe),
    }, pd.DataFrame(linhas_medicao)


def validar_colunas_forms(df):
    regras = {
        "A": "texto",
        "B": "registro",
        "G": "mes",
        "H": "numero",
        "L": "texto",
        "O": "texto",
        "P": "texto",
        "R": "texto",
    }
    meses_validos = set(MONTH_ORDER)
    suspeitas = []
    colunas_ausentes = []

    def valor_presente(valor):
        return not pd.isna(valor) and str(valor).strip() != ""

    def parece_numero(valor):
        if isinstance(valor, bool):
            return False
        if isinstance(valor, (int, float)):
            return pd.notna(valor)

        texto = str(valor).strip().replace("\u00a0", "").replace(" ", "")
        if not texto:
            return False
        if "," in texto and "." in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        try:
            return math.isfinite(float(texto))
        except ValueError:
            return False

    def problema_para_valor(tipo, valor):
        texto = str(valor).strip()
        if tipo == "mes":
            mes_informado = limpar_texto(valor).rstrip(".")
            meses_aceitos = meses_validos | {mes[:3] for mes in meses_validos}
            if mes_informado in meses_aceitos:
                return None
            sugestao = get_close_matches(
                mes_informado,
                sorted(meses_validos),
                n=1,
                cutoff=0.75,
            )
            if sugestao:
                return f"valor não corresponde a um mês válido; sugestão: {sugestao[0]}"
            return "valor não corresponde a um mês válido"
        if tipo == "numero":
            return None if parece_numero(valor) else "esperado um valor numérico"
        if tipo == "registro":
            if isinstance(valor, bool):
                return "registro deve conter somente números"
            if isinstance(valor, int):
                return None if valor >= 0 else "registro deve conter somente números"
            if isinstance(valor, float) and valor >= 0 and valor.is_integer():
                return None
            return None if texto.isdigit() else "registro deve conter somente números"
        if tipo == "texto" and re.search(r"\d", texto):
            return "contém números em uma coluna destinada a texto"
        return None

    for letra, tipo in regras.items():
        posicao = ord(letra) - ord("A")
        if posicao >= len(df.columns):
            colunas_ausentes.append(letra)
            continue

        nome_coluna = str(df.columns[posicao]).strip()
        for linha_origem, valor in df.iloc[:, posicao].items():
            if not valor_presente(valor):
                continue
            motivo = problema_para_valor(tipo, valor)
            if motivo:
                suspeitas.append(
                    {
                        "Linha": linha_origem,
                        "Coluna": f"{letra} — {nome_coluna}",
                        "Valor encontrado": str(valor),
                        "Verificação": motivo,
                    }
                )

    return pd.DataFrame(suspeitas), colunas_ausentes


def letra_coluna_excel(indice):
    letra = ""
    while indice >= 0:
        indice, resto = divmod(indice, 26)
        letra = chr(ord("A") + resto) + letra
        indice -= 1
    return letra


def diagnosticar_estrutura_forms(df):
    ancoras = [
        ("A", lambda nome: "NOME" in nome, "Nome"),
        ("B", lambda nome: "REGISTRO" in nome, "Registro"),
        ("G", lambda nome: "MES" in nome, "Mês"),
        (
            "L",
            lambda nome: "SETOR" in nome and "RESPONSAVEL" in nome,
            "Setor responsável",
        ),
        ("O", lambda nome: "STATUS" in nome, "Status"),
        (
            "P",
            lambda nome: "SITUACAO" in nome and ("PRAZO" in nome or nome == "SITUACAO"),
            "Situação do prazo",
        ),
        ("R", lambda nome: "EMAIL" in nome or "E MAIL" in nome, "E-mail"),
    ]
    cabecalhos = [chave_coluna(coluna) for coluna in df.columns]
    achados = []
    diagnosticos = []
    for coluna_esperada, reconhece, nome_campo in ancoras:
        indice_esperado = ord(coluna_esperada) - ord("A")
        correspondencias = [
            indice for indice, cabecalho in enumerate(cabecalhos) if reconhece(cabecalho)
        ]
        indice_encontrado = correspondencias[0] if len(correspondencias) == 1 else None
        achados.append(
            (indice_esperado, indice_encontrado, coluna_esperada, nome_campo)
        )
        if indice_encontrado is None:
            motivo = (
                "cabeçalho duplicado ou ambíguo"
                if len(correspondencias) > 1
                else "cabeçalho esperado não encontrado"
            )
            diagnosticos.append(
                {
                    "Localização": coluna_esperada,
                    "Esperado": nome_campo,
                    "Encontrado": (
                        ", ".join(
                            f"{letra_coluna_excel(indice)} ({df.columns[indice]})"
                            for indice in correspondencias
                        )
                        if correspondencias
                        else str(df.columns[indice_esperado])
                        if indice_esperado < len(df.columns)
                        else "(coluna ausente)"
                    ),
                    "Diagnóstico": motivo,
                }
            )
        elif indice_encontrado != indice_esperado:
            diagnosticos.append(
                {
                    "Localização": (
                        f"{coluna_esperada} → "
                        f"{letra_coluna_excel(indice_encontrado)}"
                    ),
                    "Esperado": nome_campo,
                    "Encontrado": str(df.columns[indice_encontrado]),
                    "Diagnóstico": "coluna deslocada em relação ao modelo",
                }
            )

    deslocadas = [item for item in achados if item[1] is not None and item[1] != item[0]]
    if deslocadas:
        primeira_deslocada = min(deslocadas, key=lambda item: item[0])
        anteriores_alinhadas = [
            item
            for item in achados
            if item[1] == item[0] and item[0] < primeira_deslocada[0]
        ]
        if anteriores_alinhadas:
            anterior = max(anteriores_alinhadas, key=lambda item: item[0])
            inicio_inspecao = anterior[0] + 1
        else:
            inicio_inspecao = 0
        diagnostico = (
            "possível coluna adicionada"
            if primeira_deslocada[1] > primeira_deslocada[0]
            else "possível coluna removida ou deslocada"
        )
        inicio_letra = letra_coluna_excel(inicio_inspecao)
        fim_letra = letra_coluna_excel(primeira_deslocada[0])
        encontrado = (
            f"{primeira_deslocada[3]} deslocada para "
            f"{letra_coluna_excel(primeira_deslocada[1])}"
        )
        diagnosticos.append(
            {
                "Localização": (
                    f"entre {inicio_letra} e {fim_letra} (revisar)"
                ),
                "Esperado": "continuidade do modelo",
                "Encontrado": encontrado,
                "Diagnóstico": diagnostico,
            }
        )

    return pd.DataFrame(diagnosticos)


def chave_registro(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0") and texto[:-2].isdigit():
        return texto[:-2]
    return texto


def construir_mapa_equipe_excel(arquivo):
    nome = str(getattr(arquivo, "name", "") or "").lower()
    if not nome.endswith((".xlsx", ".xlsm")):
        return {}

    workbook = openpyxl.load_workbook(
        io.BytesIO(arquivo.getvalue()),
        read_only=True,
        data_only=True,
    )
    try:
        aba_mo = next(
            (sheet for sheet in workbook.sheetnames if chave_coluna(sheet) == "MO"),
            None,
        )
        if aba_mo is None:
            return {}

        worksheet = workbook[aba_mo]
        linhas = list(worksheet.iter_rows(values_only=True))
        linha_cabecalho = next(
            (
                indice
                for indice, linha in enumerate(linhas[:15])
                if "REGISTRO" in {chave_coluna(valor) for valor in linha if valor is not None}
            ),
            None,
        )
        if linha_cabecalho is None:
            raise ValueError("A aba MO foi encontrada, mas não contém o cabeçalho REGISTRO.")

        dataframe = pd.DataFrame(linhas[linha_cabecalho + 1 :], columns=linhas[linha_cabecalho])
        colunas = {
            "registro": localizar_coluna(dataframe, ["REGISTRO"]),
            "gerente": localizar_coluna(dataframe, ["SUPERVISOR", "GERENTE"]),
            "lider": localizar_coluna(dataframe, ["LIDER", "LÍDER"]),
            "turno": localizar_coluna(dataframe, ["TURNO"]),
        }
        mapa = {}
        for _, linha in dataframe.iterrows():
            registro = chave_registro(linha[colunas["registro"]])
            if not registro:
                continue
            mapa[registro] = {
                nome: "" if pd.isna(linha[coluna]) else str(linha[coluna]).strip()
                for nome, coluna in (
                    ("Gerente", colunas["gerente"]),
                    ("Líder", colunas["lider"]),
                    ("Turno", colunas["turno"]),
                )
            }
        return mapa
    finally:
        workbook.close()


def construir_dataframe_forms(df_bruto, mapa_equipe=None):
    if df_bruto is None or df_bruto.empty:
        raise ValueError("Nenhum dado foi lido no arquivo carregado.")

    df = df_bruto.copy()
    df.columns = [str(col).strip() for col in df.columns]

    col_status = localizar_coluna(df, ["STATUS"], fallback_positions=[14, 15, 16])
    col_mes = localizar_coluna(df, ["MES", "MÊS"], fallback_positions=[5, 6, 7])
    col_setor = localizar_coluna(df, ["SETOR RESPONSAVEL PELA ACAO", "SETOR RESPONSAVEL PELA AÇÃO", "SETOR RESPONSAVEL PELA ACAO?", "SETOR RESPONSAVEL PELA AÇÃO?"], fallback_positions=[10, 11, 12])
    col_prazo = localizar_coluna(df, ["SITUACAO PRAZO", "SITUACAO PRAZO?", "SITUAÇÃO PRAZO", "SITUACAO"], fallback_positions=[14, 15])
    col_email = localizar_coluna(df, ["E MAIL", "E-MAIL", "EMAIL"], fallback_positions=[17, 18])
    col_registro = localizar_coluna(df, ["REGISTRO"], fallback_positions=[1])
    try:
        col_data_prazo = localizar_coluna(
            df,
            ["PRAZO DA ACAO", "PRAZO DA AÇÃO", "DATA LIMITE"],
            fallback_positions=[13],
        )
    except KeyError:
        col_data_prazo = "_data_prazo_dashboard"
        df[col_data_prazo] = ""

    for coluna in [col_status, col_mes, col_setor, col_prazo, col_email]:
        if coluna in df.columns:
            df[coluna] = df[coluna].fillna("").astype(str).str.strip()

    df[col_status] = df[col_status].apply(lambda x: limpar_texto(x))
    df[col_mes] = df[col_mes].apply(normalizar_mes)
    df[col_prazo] = df[col_prazo].apply(lambda x: limpar_texto(x))
    df[col_setor] = df[col_setor].apply(lambda x: str(x).strip())
    df[col_email] = df[col_email].apply(lambda x: str(x).strip())

    mapa_equipe = mapa_equipe or {}
    colunas_diretas = {}
    for destino, aliases in (
        ("Gerente", ["GERENTE", "SUPERVISOR"]),
        ("Líder", ["LIDER", "LÍDER", "QUEM É O LIDER DESSA PESSOA"]),
        ("Turno", ["TURNO"]),
    ):
        try:
            colunas_diretas[destino] = localizar_coluna(df, aliases)
        except KeyError:
            colunas_diretas[destino] = None

    for destino in ("Gerente", "Líder", "Turno"):
        if colunas_diretas[destino] is None:
            df[destino] = ""
        else:
            df[destino] = df[colunas_diretas[destino]].fillna("").astype(str).str.strip()

    if mapa_equipe:
        equipe_por_registro = df[col_registro].map(
            lambda registro: mapa_equipe.get(chave_registro(registro), {})
        )
        for destino in ("Gerente", "Líder", "Turno"):
            valores_equipe = equipe_por_registro.map(
                lambda pessoa: pessoa.get(destino, "")
            )
            df[destino] = valores_equipe.where(
                valores_equipe.astype(str).str.strip().ne(""),
                df[destino],
            )

    for destino in ("Gerente", "Líder", "Turno"):
        df[destino] = df[destino].replace("", "(Não informado)")

    if df.empty:
        raise ValueError("A tabela FORMS foi carregada, mas não contém linhas válidas.")

    if not (df[col_status].astype(str).str.upper().str.contains("CONCLUIDO").any() or df[col_status].astype(str).str.upper().str.contains("ABERTO").any()):
        raise ValueError("Os dados carregados não parecem conter valores válidos em STATUS (Concluído/Aberto).")

    return df, {
        "status": col_status,
        "mes": col_mes,
        "setor": col_setor,
        "prazo": col_prazo,
        "data_prazo": col_data_prazo,
        "email": col_email,
        "registro": col_registro,
        "gerente": "Gerente",
        "lider": "Líder",
        "turno": "Turno",
    }


def ler_excel_bruto(arquivo, engine):
    conteudo = arquivo.getvalue()
    return pd.ExcelFile(io.BytesIO(conteudo), engine=engine)


def construir_dataframe_custos(arquivo):
    if arquivo is None:
        raise ValueError("Carregue a planilha de custos para visualizar o painel.")

    nome = str(arquivo.name or "").lower()
    if not nome.endswith((".xlsx", ".xlsm")):
        raise ValueError("A planilha de custos deve estar no formato .xlsx ou .xlsm.")

    workbook = ler_excel_bruto(arquivo, "openpyxl")
    aliases = {
        "Projeto": ["PROJETO"],
        "Valor": ["VALOR"],
        "Período Serviço": ["PERIODO SERVICO", "PERIODO", "MES"],
        "EMPRESA": ["EMPRESA"],
        "CUSTO COM": ["CUSTO COM", "CUSTO"],
    }
    df = None

    for nome_aba in workbook.sheet_names:
        dados_sem_cabecalho = workbook.parse(nome_aba, header=None)
        linha_cabecalho = None
        for idx, linha in dados_sem_cabecalho.iterrows():
            cabecalhos = {chave_coluna(valor) for valor in linha if pd.notna(valor)}
            if "PROJETO" in cabecalhos and "VALOR" in cabecalhos:
                linha_cabecalho = idx
                break
        if linha_cabecalho is not None:
            df = workbook.parse(nome_aba, skiprows=linha_cabecalho)
            break

    workbook.close()
    if df is None or df.empty:
        raise ValueError("Não foi encontrado cabeçalho com as colunas Projeto e Valor.")

    colunas_encontradas = {}
    lookup = {chave_coluna(coluna): coluna for coluna in df.columns}
    for coluna_canonica, nomes in aliases.items():
        original = next(
            (lookup[chave_coluna(nome)] for nome in nomes if chave_coluna(nome) in lookup),
            None,
        )
        if original is None:
            raise ValueError(
                f"A planilha de custos precisa conter a coluna '{coluna_canonica}'."
            )
        colunas_encontradas[original] = coluna_canonica

    colunas_reconhecidas = set(colunas_encontradas)
    colunas_adicionais = [
        {
            "Localização": letra_coluna_excel(indice),
            "Cabeçalho encontrado": str(coluna),
            "Diagnóstico": "coluna não utilizada pelo painel",
        }
        for indice, coluna in enumerate(df.columns)
        if coluna not in colunas_reconhecidas
    ]
    df = df.rename(columns=colunas_encontradas)
    df = df[list(aliases)].copy()
    df = df.dropna(how="all")

    def converter_valor(valor):
        if pd.isna(valor):
            return float("nan")
        if isinstance(valor, (int, float)):
            return float(valor)

        texto = str(valor).strip().replace("R$", "").replace("\u00a0", "").replace(" ", "")
        if not texto:
            return float("nan")
        if "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        try:
            return float(texto)
        except ValueError as erro:
            raise ValueError(
                f"Valor monetário inválido na planilha de custos: '{valor}'."
            ) from erro

    df["Valor"] = df["Valor"].apply(converter_valor)
    df = df.dropna(subset=["Projeto", "Valor"])
    if df.empty:
        raise ValueError("A planilha de custos não contém lançamentos válidos.")

    for coluna in ["Projeto", "Período Serviço", "EMPRESA", "CUSTO COM"]:
        df[coluna] = df[coluna].fillna("").astype(str).str.strip()

    return df, pd.DataFrame(colunas_adicionais)


def construir_dataframe_organograma(arquivo):
    if arquivo is None:
        raise ValueError("Carregue a planilha do organograma semanal para visualizar o painel.")

    nome = str(arquivo.name or "").lower()
    if nome.endswith((".xlsx", ".xlsm")):
        engine = "openpyxl"
    elif nome.endswith(".xlsb"):
        engine = "pyxlsb"
    else:
        raise ValueError("O organograma deve estar no formato .xlsx, .xlsm ou .xlsb.")

    workbook = ler_excel_bruto(arquivo, engine)
    linhas_brutas = None
    linha_cabecalho = None
    valores_cabecalho = None
    for nome_aba in workbook.sheet_names:
        candidato = workbook.parse(nome_aba, header=None)
        for idx, linha in candidato.iterrows():
            cabecalhos = [chave_coluna(valor) for valor in linha if pd.notna(valor)]
            if "ATIVIDADE" in cabecalhos and "HORAS" in cabecalhos:
                linhas_brutas = candidato.iloc[idx + 1 :]
                linha_cabecalho = int(idx) + 1
                valores_cabecalho = linha.tolist()
                break
        if linhas_brutas is not None:
            break
    workbook.close()

    if linhas_brutas is None:
        raise ValueError("Não foi encontrado cabeçalho com as colunas ATIVIDADE e HORAS.")

    diagnosticos_estrutura = []
    colunas_atividade = {1, 4, 7, 10, 13}
    colunas_horas = {2, 5, 8, 11, 14}
    atividade_encontrada = []
    horas_encontradas = []
    for indice, valor in enumerate(valores_cabecalho):
        cabecalho = chave_coluna(valor)
        if cabecalho == "ATIVIDADE":
            atividade_encontrada.append(indice)
            if indice not in colunas_atividade:
                diagnosticos_estrutura.append(
                    {
                        "Localização": f"{letra_coluna_excel(indice)}{linha_cabecalho}",
                        "Esperado": "Atividade nos blocos semanais",
                        "Encontrado": str(valor),
                        "Diagnóstico": "cabeçalho deslocado; conferir coluna inserida ou movida",
                    }
                )
        elif cabecalho == "HORAS":
            horas_encontradas.append(indice)
            if indice not in colunas_horas:
                diagnosticos_estrutura.append(
                    {
                        "Localização": f"{letra_coluna_excel(indice)}{linha_cabecalho}",
                        "Esperado": "Horas nos blocos semanais",
                        "Encontrado": str(valor),
                        "Diagnóstico": "cabeçalho deslocado; conferir coluna inserida ou movida",
                    }
                )

    if not any(indice in colunas_atividade for indice in atividade_encontrada):
        diagnosticos_estrutura.append(
            {
                "Localização": "B",
                "Esperado": "Atividade",
                "Encontrado": ", ".join(
                    letra_coluna_excel(indice) for indice in atividade_encontrada
                ),
                "Diagnóstico": "nenhum cabeçalho de atividade está na posição esperada",
            }
        )
    if not any(indice in colunas_horas for indice in horas_encontradas):
        diagnosticos_estrutura.append(
            {
                "Localização": "C",
                "Esperado": "Horas",
                "Encontrado": ", ".join(
                    letra_coluna_excel(indice) for indice in horas_encontradas
                ),
                "Diagnóstico": "nenhum cabeçalho de horas está na posição esperada",
            }
        )
    if diagnosticos_estrutura:
        return None, 0, pd.DataFrame(diagnosticos_estrutura)

    dias_mapeamento = [
        ("Segunda", 1, 2, 3),
        ("Terça", 4, 5, 6),
        ("Quarta", 7, 8, 9),
        ("Quinta", 10, 11, 12),
        ("Sexta", 13, 14, 15),
    ]
    linhas_estruturadas = []
    responsaveis_nao_identificados = 0
    colunas_adicionais = {}
    for indice_coluna in range(16, len(valores_cabecalho)):
        valor_cabecalho = valores_cabecalho[indice_coluna]
        if pd.notna(valor_cabecalho) and str(valor_cabecalho).strip():
            localizacao = letra_coluna_excel(indice_coluna)
            colunas_adicionais[localizacao] = {
                "Localização": f"{localizacao}{linha_cabecalho}",
                "Esperado": "colunas utilizadas pelo modelo semanal",
                "Encontrado": str(valor_cabecalho),
                "Diagnóstico": "cabeçalho em coluna não utilizada pelo painel",
            }

    for indice_linha, linha in linhas_brutas.iterrows():
        primeira_coluna = str(linha.iloc[0]).strip().upper() if len(linha) else ""
        if "TOTAL" in primeira_coluna or "HORAS PREVISTAS" in primeira_coluna:
            break

        for indice_coluna in range(16, len(linha)):
            if pd.notna(linha.iloc[indice_coluna]) and str(linha.iloc[indice_coluna]).strip():
                localizacao = letra_coluna_excel(indice_coluna)
                colunas_adicionais.setdefault(
                    localizacao,
                    {
                        "Localização": f"{localizacao}{int(indice_linha) + 1}",
                        "Esperado": "colunas utilizadas pelo modelo semanal",
                        "Encontrado": str(linha.iloc[indice_coluna]),
                        "Diagnóstico": "conteúdo em coluna não utilizada pelo painel",
                    },
                )

        for dia, col_atividade, col_horas, col_responsavel in dias_mapeamento:
            if max(col_atividade, col_horas, col_responsavel) >= len(linha):
                continue

            atividade = linha.iloc[col_atividade]
            horas = linha.iloc[col_horas]
            responsavel = chave_coluna(linha.iloc[col_responsavel])
            if pd.isna(atividade) or not str(atividade).strip() or pd.isna(horas):
                continue

            try:
                horas_valor = float(str(horas).strip().replace(",", "."))
            except ValueError as erro:
                raise ValueError(
                    f"Valor de horas inválido para a atividade '{atividade}': '{horas}'."
                ) from erro

            tokens_responsavel = set(responsavel.split())
            ambos_responsaveis = responsavel == "NG" or (
                "N" in tokens_responsavel and "G" in tokens_responsavel
            )
            nickolas = ambos_responsaveis or "NICKOLAS" in tokens_responsavel or "N" in tokens_responsavel
            giovana = ambos_responsaveis or "GIOVANA" in tokens_responsavel or "G" in tokens_responsavel
            if nickolas and giovana:
                linhas_estruturadas.extend(
                    [
                        [str(atividade).strip(), horas_valor / 2, "Nickolas", dia],
                        [str(atividade).strip(), horas_valor / 2, "Giovana", dia],
                    ]
                )
            elif nickolas:
                linhas_estruturadas.append(
                    [str(atividade).strip(), horas_valor, "Nickolas", dia]
                )
            elif giovana:
                linhas_estruturadas.append(
                    [str(atividade).strip(), horas_valor, "Giovana", dia]
                )
            else:
                responsaveis_nao_identificados += 1

    if not linhas_estruturadas:
        raise ValueError("A planilha não contém atividades com horas e responsáveis reconhecidos.")

    df = pd.DataFrame(
        linhas_estruturadas,
        columns=["Atividade", "Horas", "Responsavel", "Dia_Semana"],
    )
    return df, responsaveis_nao_identificados, pd.DataFrame(colunas_adicionais.values())


def formatar_reais(valor):
    formatado = f"{valor:,.2f}"
    return "R$ " + formatado.replace(",", "X").replace(".", ",").replace("X", ".")


def salvar_upload_se_necessario(arquivo, tabela, dataframe):
    quantidade = salvar_dataframe(tabela, dataframe, arquivo.name)
    if quantidade:
        st.cache_data.clear()
    return quantidade > 0


def carregar_dataframe_cacheado(tabela):
    return carregar_dataframe(tabela)


def exibir_painel_custos(df_custos, colunas_adicionais=None):
    colunas_adicionais = (
        colunas_adicionais
        if colunas_adicionais is not None
        else pd.DataFrame()
    )
    with st.container(border=True):
        st.subheader("Histórico de custos da qualidade")
        st.success(f"✓ {len(df_custos)} lançamento(s) carregado(s) do banco de dados")
        if colunas_adicionais.empty:
            st.success("✓ Nenhuma coluna adicional detectada")
        else:
            st.warning(
                f"⚠️ {len(colunas_adicionais)} coluna(s) fora do modelo; "
                "não são utilizadas pelo painel."
            )
            with st.expander("Ver localização das colunas adicionais"):
                st.dataframe(
                    colunas_adicionais,
                    hide_index=True,
                    use_container_width=True,
                )
    st.subheader("Custos da não Qualidade Cliente")
    filtro_empresa, filtro_periodo, filtro_custo = st.columns(3)

    with filtro_empresa:
        empresa = st.selectbox(
            "Empresa",
            ["Todos"] + sorted(df_custos["EMPRESA"].replace("", "(Sem empresa)").unique()),
            key="custos_empresa",
        )
    with filtro_periodo:
        periodo = st.selectbox(
            "Período Serviço",
            ["Todos"]
            + sorted(df_custos["Período Serviço"].replace("", "(Sem período)").unique()),
            key="custos_periodo",
        )
    with filtro_custo:
        tipo_custo = st.selectbox(
            "Custo com",
            ["Todos"] + sorted(df_custos["CUSTO COM"].replace("", "(Sem tipo)").unique()),
            key="custos_tipo",
        )

    df_filtrado = df_custos.copy()
    for coluna, valor in [
        ("EMPRESA", empresa),
        ("Período Serviço", periodo),
        ("CUSTO COM", tipo_custo),
    ]:
        if valor != "Todos":
            valor_real = "" if valor.startswith("(Sem ") else valor
            df_filtrado = df_filtrado[df_filtrado[coluna] == valor_real]

    maior_custo = df_filtrado["Valor"].max() if not df_filtrado.empty else 0.0
    metrica_total, metrica_registros, metrica_maior = st.columns(3)
    metrica_total.metric("Custo total acumulado", formatar_reais(df_filtrado["Valor"].sum()))
    metrica_registros.metric("Ocorrências / lançamentos", f"{len(df_filtrado)} itens")
    metrica_maior.metric("Maior despesa registrada", formatar_reais(maior_custo))

    if df_filtrado.empty:
        st.info("Não há lançamentos para os filtros selecionados.")
        return

    por_projeto = (
        df_filtrado.groupby("Projeto", as_index=False)["Valor"]
        .sum()
        .sort_values("Valor", ascending=True)
    )
    figura = go.Figure(
        go.Bar(
            y=por_projeto["Projeto"],
            x=por_projeto["Valor"],
            orientation="h",
            marker_color="#e99a2b",
            text=[formatar_reais(valor) for valor in por_projeto["Valor"]],
            textposition="inside",
            textfont=dict(color="#172b43", size=12),
        )
    )
    figura.update_layout(
        title="Custos acumulados por projeto",
        xaxis_title="Custo acumulado",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#f7f9fc"),
        xaxis=dict(gridcolor="rgba(247,249,252,.18)", zerolinecolor="rgba(247,249,252,.18)"),
        margin=dict(l=100, r=20, t=50, b=20),
        height=max(300, 80 * len(por_projeto)),
    )
    st.plotly_chart(figura, width="stretch")


def exibir_painel_organograma(arquivo):
    df_semana, responsaveis_nao_identificados, diagnosticos_estrutura = (
        construir_dataframe_organograma(arquivo)
    )
    with st.container(border=True):
        st.subheader("Validação do organograma semanal")
        if df_semana is None:
            st.error("✗ Estrutura semanal incompatível; o painel não foi calculado.")
            st.dataframe(
                diagnosticos_estrutura,
                hide_index=True,
                use_container_width=True,
            )
            return
        st.success("✓ Cabeçalhos dos blocos semanais encontrados")
        if diagnosticos_estrutura.empty:
            st.success("✓ Nenhuma coluna adicional detectada")
        else:
            st.warning(
                f"⚠️ {len(diagnosticos_estrutura)} conteúdo(s) em colunas não utilizadas"
            )
            with st.expander("Ver localização dos conteúdos adicionais"):
                st.dataframe(
                    diagnosticos_estrutura,
                    hide_index=True,
                    use_container_width=True,
                )
    if responsaveis_nao_identificados:
        st.warning(
            f"{responsaveis_nao_identificados} atividade(s) com horas foram ignoradas "
            "porque o responsável não foi reconhecido como Nickolas, Giovana ou ambos."
        )

    st.subheader("Planejamento Semanal — Capacidade Industrial")
    seletor_pessoa, seletor_dia = st.columns([1, 2])
    with seletor_pessoa:
        pessoa = st.selectbox(
            "Recurso",
            ["Equipe Total", "Giovana", "Nickolas"],
            key="organograma_pessoa",
        )
    with seletor_dia:
        dia = st.selectbox(
            "Janela temporal",
            ["Semana Completa", "Segunda", "Terça", "Quarta", "Quinta", "Sexta"],
            key="organograma_dia",
        )

    df_pessoa = df_semana.copy()
    if pessoa != "Equipe Total":
        df_pessoa = df_pessoa[df_pessoa["Responsavel"] == pessoa]
    df_grafico = df_pessoa.copy()
    capacidade = 80.0 if pessoa == "Equipe Total" else 40.0
    if dia != "Semana Completa":
        df_grafico = df_grafico[df_grafico["Dia_Semana"] == dia]
        capacidade = 16.0 if pessoa == "Equipe Total" else 8.0

    total_horas = df_grafico["Horas"].sum()
    percentual_ocupacao = total_horas / capacidade * 100 if capacidade else 0.0
    saldo = capacidade - total_horas
    total_atividades = df_grafico["Atividade"].nunique()
    maior_carga = df_grafico["Horas"].max() if not df_grafico.empty else 0.0

    metrica_horas, metrica_ocupacao, metrica_saldo, metrica_atividades, metrica_maior = st.columns(5)
    metrica_horas.metric("Carga calculada", f"{total_horas:.1f} h")
    metrica_ocupacao.metric("Taxa de ocupação", f"{percentual_ocupacao:.1f}%")
    metrica_saldo.metric("Saldo de capacidade", f"{saldo:+.1f} h")
    metrica_atividades.metric("Volumetria de tarefas", f"{total_atividades} itens")
    metrica_maior.metric("Maior carga única", f"{maior_carga:.1f} h")

    if pessoa == "Equipe Total":
        y_labels = ["Nickolas", "Giovana", "Equipe Total"]
        horas_por_recurso = [
            df_grafico.loc[df_grafico["Responsavel"] == "Nickolas", "Horas"].sum(),
            df_grafico.loc[df_grafico["Responsavel"] == "Giovana", "Horas"].sum(),
            total_horas,
        ]
    else:
        y_labels = [pessoa]
        horas_por_recurso = [total_horas]

    figura_barras = go.Figure(
        go.Bar(
            y=y_labels,
            x=horas_por_recurso,
            orientation="h",
            marker_color="#e99a2b",
            text=[f"{valor:.1f} h" if valor > 0 else "" for valor in horas_por_recurso],
            textposition="inside",
            textfont=dict(color="#172b43", size=12),
        )
    )
    figura_barras.add_shape(
        type="line",
        x0=capacidade,
        y0=-0.5,
        x1=capacidade,
        y1=len(y_labels) - 0.5,
        line=dict(color="#f7f9fc", width=2, dash="dash"),
    )
    figura_barras.update_layout(
        title="Carga de trabalho vs. meta de capacidade",
        xaxis_title="Horas acumuladas",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#f7f9fc"),
        xaxis=dict(gridcolor="rgba(247,249,252,.18)", zerolinecolor="rgba(247,249,252,.18)"),
        margin=dict(l=100, r=20, t=50, b=20),
        height=300,
        showlegend=False,
    )

    col_grafico_barras, col_grafico_dias = st.columns([3, 2])
    with col_grafico_barras:
        st.plotly_chart(figura_barras, width="stretch")
    with col_grafico_dias:
        ordem_dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
        horas_por_dia = (
            df_pessoa.groupby("Dia_Semana")["Horas"]
            .sum()
            .reindex(ordem_dias, fill_value=0)
        )
        figura_dias = go.Figure(
            go.Pie(
                labels=horas_por_dia.index,
                values=horas_por_dia.values,
                hole=0.4,
                marker=dict(colors=["#e99a2b", "#f2b95f", "#82c7ff", "#4f8fc8", "#203a58"]),
                textinfo="label+percent",
                textfont=dict(color="#f7f9fc", size=11),
            )
        )
        figura_dias.update_layout(
            title="Distribuição da carga por dia",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#f7f9fc"),
            height=300,
            margin=dict(l=20, r=20, t=50, b=20),
            showlegend=False,
        )
        st.plotly_chart(figura_dias, width="stretch")


def _carregar_indicadores_central():
    indicadores = {
        "grupo_qualidade": ("— ações abertas", "— vencidas"),
        "grupo_confiabilidade": ("— CNs", "— atrasadas"),
        "grupo_calibracao": ("— equipamentos", "— vencidos"),
    }

    try:
        historico_acoes = carregar_dataframe_cacheado("acoes_qualidade")
        if historico_acoes.empty:
            indicadores["grupo_qualidade"] = ("0 ações abertas", "0 vencidas")
        else:
            acoes, colunas = construir_dataframe_forms(historico_acoes)
            status = acoes[colunas["status"]].astype(str).str.upper().str.strip()
            prazo = acoes[colunas["prazo"]].astype(str).str.upper().str.strip()
            abertas = status.eq("ABERTO")
            indicadores["grupo_qualidade"] = (
                f"{int(abertas.sum())} ações abertas",
                f"{int((abertas & prazo.eq('VENCIDO')).sum())} vencidas",
            )
    except (KeyError, ValueError, SQLAlchemyError, sqlite3.Error):
        logger.exception("Falha ao carregar indicadores de ações da qualidade.")
        indicadores["grupo_qualidade"] = (
            "Indicadores indisponíveis",
            "Verifique o histórico de ações",
        )

    try:
        from database import cn_store

        resumir_cns = getattr(cn_store, "resumir_cns", None)
        if not callable(resumir_cns):
            logger.error(
                "O database/cn_store.py publicado não contém resumir_cns; "
                "publique a versão atualizada do módulo."
            )
            indicadores["grupo_confiabilidade"] = (
                "Indicadores indisponíveis",
                "Atualize database/cn_store.py",
            )
        else:
            total_cns, cns_atrasadas = resumir_cns()
            indicadores["grupo_confiabilidade"] = (
                f"{total_cns} CNs",
                f"{cns_atrasadas} atrasadas",
            )
    except ModuleNotFoundError as erro:
        if erro.name not in {"database.cn_store", "core.cn_management"}:
            raise
        logger.warning("Módulo de Gestão de CN's não disponível para o resumo da Home.")
        indicadores["grupo_confiabilidade"] = (
            "Indicadores indisponíveis",
            "Módulo de CN's não publicado",
        )
    except (KeyError, SQLAlchemyError, pd.errors.DatabaseError):
        logger.exception("Falha ao carregar indicadores das CNs.")
        indicadores["grupo_confiabilidade"] = (
            "Indicadores indisponíveis",
            "Verifique a Gestão de CN's",
        )

    try:
        from database.calibration_store import resumir_equipamentos

        total_equipamentos, vencidos = resumir_equipamentos()
        indicadores["grupo_calibracao"] = (
            f"{total_equipamentos} equipamentos",
            f"{vencidos} vencidos",
        )
    except (KeyError, SQLAlchemyError, pd.errors.DatabaseError):
        logger.exception("Falha ao carregar indicadores de calibração.")
        indicadores["grupo_calibracao"] = (
            "Indicadores indisponíveis",
            "Verifique a Gestão de Calibração",
        )

    return indicadores


if st.session_state.get("painel_ativo") == "Ações da Qualidade":
    st.session_state.painel_ativo = "Dashboard Operacional"
elif st.session_state.get("painel_ativo") not in DESTINOS_VALIDOS:
    st.session_state.painel_ativo = "Home"

tipo_painel = st.session_state.painel_ativo
if tipo_painel == "Home":
    with st.container(border=True):
        st.markdown(
            """
            <div class="welcome-hero">
                <div class="welcome-kicker">Central de indicadores</div>
                <div class="welcome-title">Gestão avançada de indicadores da qualidade</div>
                <div class="welcome-accent"></div>
                <div class="welcome-copy">Acompanhe indicadores, ações e resultados em um só lugar.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.subheader("Central de ferramentas")
    indicadores = _carregar_indicadores_central()
    for inicio in range(0, len(CARDS_CENTRAL), 2):
        colunas_cards = st.columns(2)
        for coluna, (titulo, icone, descricao, destino) in zip(
            colunas_cards, CARDS_CENTRAL[inicio : inicio + 2]
        ):
            with coluna:
                metricas = indicadores.get(destino)
                metricas_html = ""
                if metricas:
                    metricas_html = (
                        "<div class='hub-metrics'>"
                        f"<span class='hub-metric'>{metricas[0]}</span>"
                        f"<span class='hub-metric'>{metricas[1]}</span>"
                        "</div>"
                    )
                card_html = (
                    f'<div class="hub-card"><div class="hub-icon">{icone}</div>'
                    f'<div class="hub-content"><div class="hub-title">{titulo}</div>'
                    f'<div class="hub-copy">{descricao}</div>{metricas_html}'
                    "</div></div>"
                )
                st.markdown(card_html, unsafe_allow_html=True)
                rotulo_botao = (
                    "Abrir área →"
                    if destino in GRUPOS_FERRAMENTAS
                    else "Acessar custos →"
                    if destino == "Custos da não Qualidade"
                    else "Acessar administração →"
                )
                if st.button(
                    rotulo_botao,
                    key=f"acessar_{inicio}_{destino}",
                    use_container_width=True,
                ):
                    st.session_state.painel_ativo = destino
                    st.rerun()
    st.stop()

if tipo_painel in GRUPOS_FERRAMENTAS:
    grupo = GRUPOS_FERRAMENTAS[tipo_painel]
    if st.button("⬅️ Voltar à Central de Ferramentas", key="voltar_central_grupo"):
        st.session_state.painel_ativo = "Home"
        st.rerun()
    st.markdown(
        f"<div class='header-blue'>{grupo['icone']} {grupo['titulo']}</div>",
        unsafe_allow_html=True,
    )
    st.caption(grupo["descricao"])
    st.subheader("Ferramentas da área")
    paineis_grupo = [
        PAINEIS_POR_NOME[nome]
        for nome in grupo["paineis"]
    ]
    for inicio in range(0, len(paineis_grupo), 2):
        colunas_cards = st.columns(2)
        for coluna, (painel, icone, titulo, descricao) in zip(
            colunas_cards, paineis_grupo[inicio : inicio + 2]
        ):
            with coluna:
                st.markdown(
                    f"""
                    <div class="hub-card">
                        <div class="hub-icon">{icone}</div>
                        <div class="hub-title">{titulo}</div>
                        <div class="hub-copy">{descricao}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button(
                    "Acessar ferramenta →",
                    key=f"acessar_ferramenta_{tipo_painel}_{painel}",
                    use_container_width=True,
                ):
                    st.session_state.painel_ativo = painel
                    st.rerun()
    st.stop()

if tipo_painel in GRUPOS_FERRAMENTAS["grupo_qualidade"]["paineis"]:
    botao_central, botao_area = st.columns(2)
    if botao_central.button(
        "⬅️ Voltar à Central de Cards",
        key="voltar_central_cards",
        use_container_width=True,
    ):
        st.session_state.painel_ativo = "Home"
        st.rerun()
    if botao_area.button(
        "⬅️ Voltar às Ferramentas da Área",
        key="voltar_ferramentas_qualidade",
        use_container_width=True,
    ):
        st.session_state.painel_ativo = "grupo_qualidade"
        st.rerun()
elif st.button("⬅️ Voltar à Central de Cards", key="voltar_central_cards"):
    st.session_state.painel_ativo = "Home"
    st.rerun()
titulo_painel = (
    "Ações da Qualidade"
    if tipo_painel == "Dashboard Operacional"
    else tipo_painel
)
st.markdown(
    f"<div class='header-blue'>{titulo_painel}</div>",
    unsafe_allow_html=True,
)

arquivo_usuario = None
arquivo_certificado = []
if tipo_painel == "Certificados de Calibração":
    arquivo_certificado = st.file_uploader(
        "Carregue até 5 certificados de calibração em PDF",
        type=["pdf"],
        accept_multiple_files=True,
        max_upload_size=MAX_CERTIFICATE_FILE_MB,
        help=(
            f"Processamento sequencial; até {MAX_CERTIFICATE_FILES} arquivos, "
            f"{MAX_CERTIFICATE_FILE_MB} MB por PDF e "
            f"{MAX_CERTIFICATE_BATCH_MB} MB por lote."
        ),
        key="arquivo_certificado_calibracao",
    )
elif tipo_painel in {
    "Dashboard Operacional",
    "Custos da não Qualidade",
    "Organograma Semanal",
}:
    configuracao_upload = {
        "Dashboard Operacional": (
            "Carregue a planilha de dados",
            ["csv", "txt", "xlsx", "xlsm"],
            "arquivo_forms",
        ),
        "Custos da não Qualidade": (
            "Carregue a planilha de custos da não qualidade",
            ["xlsx", "xlsm"],
            "arquivo_custos",
        ),
        "Organograma Semanal": (
            "Carregue a planilha do organograma semanal",
            ["xlsx", "xlsm", "xlsb"],
            "arquivo_organograma",
        ),
    }
    rotulo_upload, extensoes_upload, chave_upload = configuracao_upload[tipo_painel]
    arquivo_usuario = st.file_uploader(
        rotulo_upload,
        type=extensoes_upload,
        key=chave_upload,
    )
elif tipo_painel == "Target Gerência":
    arquivo_target = st.file_uploader(
        "Carregue a planilha com a aba Target Gerência",
        type=["xlsx", "xlsm"],
        key="arquivo_target_gerencia",
    )

if tipo_painel == "Certificados de Calibração":
    st.subheader("Analisador de certificados de calibração")
    df_criterios = None
    try:
        df_criterios, avisos_criterios = ler_criterios_calibracao()
        with st.container(border=True):
            st.success(
                f"✓ Critérios internos disponíveis: {len(df_criterios)} equipamento(s)"
            )
            if len(df_criterios) != 69:
                st.warning(
                    f"⚠️ A lista contém {len(df_criterios)} equipamentos; "
                    "o total de referência informado é 69."
                )
            if avisos_criterios:
                st.warning(
                    f"⚠️ {len(avisos_criterios)} critério(s) precisam de atenção."
                )
                with st.expander("Ver critérios que exigem atenção"):
                    for aviso in avisos_criterios:
                        st.write(f"- {aviso}")
            st.dataframe(
                df_criterios,
                hide_index=True,
                use_container_width=True,
            )
    except (
        ValueError,
        KeyError,
        TypeError,
        OSError,
        InvalidFileException,
        zipfile.BadZipFile,
    ) as erro:
        st.error(f"Erro ao carregar os critérios internos de calibração: {erro}")

    if not arquivo_certificado:
        st.caption("Envie um ou mais PDFs para extrair os dados disponíveis.")
    elif df_criterios is None:
        st.warning(
            "Os PDFs não foram analisados porque os critérios não puderam ser carregados."
        )
    else:
        st.caption(
            "Até 2 PDFs são processados em paralelo, com progresso por arquivo. Limites: "
            f"{MAX_CERTIFICATE_FILES} PDFs por lote, "
            f"{MAX_CERTIFICATE_FILE_MB} MB por arquivo, "
            f"{MAX_CERTIFICATE_BATCH_MB} MB no total e "
            f"{MAX_CERTIFICATE_PAGES} páginas por PDF."
        )
        assinatura_lote = hashlib.sha256()
        for arquivo in arquivo_certificado:
            conteudo = arquivo.getvalue()
            assinatura_lote.update(arquivo.name.encode("utf-8"))
            assinatura_lote.update(len(conteudo).to_bytes(8, "big"))
            assinatura_lote.update(conteudo)
        assinatura_lote.update(
            pd.util.hash_pandas_object(df_criterios, index=True).values.tobytes()
        )
        assinatura_lote = assinatura_lote.hexdigest()

        cache_inicial = st.session_state.get("_cache_analise_pdf_inicial")
        if (
            not isinstance(cache_inicial, dict)
            or cache_inicial.get("assinatura") != assinatura_lote
        ):
            progresso_inicial = st.progress(
                0.0,
                text="Preparando análise dos certificados...",
            )

            def atualizar_progresso_inicial(concluidos, total, nome_arquivo):
                progresso_inicial.progress(
                    concluidos / total,
                    text=f"PDF {concluidos} de {total} concluído: {nome_arquivo}",
                )

            try:
                resultados_iniciais, detalhes_iniciais = analisar_lote_certificados(
                    arquivo_certificado,
                    df_criterios,
                    analisar_certificado_pdf,
                    ao_progresso=atualizar_progresso_inicial,
                )
            finally:
                progresso_inicial.empty()
            cache_inicial = {
                "assinatura": assinatura_lote,
                "resultados": resultados_iniciais,
                "detalhes": detalhes_iniciais,
            }
            st.session_state["_cache_analise_pdf_inicial"] = cache_inicial
        else:
            resultados_iniciais = cache_inicial["resultados"]
            detalhes_iniciais = cache_inicial["detalhes"]

        opcoes_criterios_manuais = criterios_para_selecao_manual(df_criterios)
        criterios_manuais = {}
        if opcoes_criterios_manuais:
            for indice, arquivo in enumerate(arquivo_certificado):
                resultado_inicial = resultados_iniciais.iloc[indice]
                if (
                    not resultado_inicial["Equipamento"]
                    or resultado_inicial["Critério aplicado"]
                    or not resultado_inicial["Medições extraídas"]
                    or not str(resultado_inicial["Resultado"]).startswith(
                        "Revisão necessária"
                    )
                ):
                    continue
                st.warning(
                    f"{arquivo.name}: não foi possível associar automaticamente "
                    f"“{resultado_inicial['Equipamento']}”. Selecione o critério "
                    "numérico correto após conferir o certificado."
                )
                assinatura = hashlib.sha256(arquivo.getvalue()).hexdigest()[:12]
                rotulos = ["— Selecione um critério —", *opcoes_criterios_manuais]
                escolha = st.selectbox(
                    f"Critério para {arquivo.name}",
                    rotulos,
                    key=f"criterio_manual_{indice}_{assinatura}",
                )
                if escolha != rotulos[0]:
                    criterios_manuais[indice] = opcoes_criterios_manuais[escolha]

        chave_cache_final = (
            assinatura_lote,
            tuple(sorted(criterios_manuais.items())),
        )
        cache_final = st.session_state.get("_cache_analise_pdf_final")
        if (
            isinstance(cache_final, dict)
            and cache_final.get("chave") == chave_cache_final
        ):
            df_resultados_pdf = cache_final["resultados"]
            detalhes_pdf = cache_final["detalhes"]
        else:
            progresso_manual = None
            if criterios_manuais:
                progresso_manual = st.progress(
                    0.0,
                    text="Reavaliando certificados com os critérios selecionados...",
                )

            def atualizar_progresso_manual(concluidos, total, nome_arquivo):
                if progresso_manual is not None:
                    progresso_manual.progress(
                        concluidos / total,
                        text=(
                            f"PDF {concluidos} de {total} reavaliado: "
                            f"{nome_arquivo}"
                        ),
                    )

            try:
                df_resultados_pdf, detalhes_pdf = analisar_lote_certificados(
                    arquivo_certificado,
                    df_criterios,
                    analisar_certificado_pdf,
                    criterios_manuais,
                    ao_progresso=atualizar_progresso_manual,
                    resultados_precalculados=(
                        resultados_iniciais,
                        detalhes_iniciais,
                    ),
                )
            finally:
                if progresso_manual is not None:
                    progresso_manual.empty()
            st.session_state["_cache_analise_pdf_final"] = {
                "chave": chave_cache_final,
                "resultados": df_resultados_pdf,
                "detalhes": detalhes_pdf,
            }
        for indice, arquivo in enumerate(arquivo_certificado):
            resultado = df_resultados_pdf.iloc[indice].to_dict()
            if str(resultado.get("Resultado", "")).startswith("Lote não analisado"):
                continue
            medicoes = detalhes_pdf.get(
                f"{indice + 1}. {arquivo.name}",
                pd.DataFrame(),
            )
            try:
                salvar_analise_certificado(
                    arquivo.name,
                    arquivo.getvalue(),
                    resultado,
                    medicoes,
                )
            except (SQLAlchemyError, ValueError) as erro:
                logger.exception("Falha ao persistir análise de certificado PDF.")
                st.error(
                    f"A análise de {arquivo.name} não foi salva para consulta "
                    f"compartilhada: {erro}"
                )
        st.dataframe(
            df_resultados_pdf,
            hide_index=True,
            use_container_width=True,
        )
        for titulo, df_medicoes in detalhes_pdf.items():
            with st.expander(f"Detalhes extraídos — {titulo}"):
                if df_medicoes.empty:
                    st.warning(
                        "Não foi possível extrair uma tabela de medições confiável."
                    )
                else:
                    st.dataframe(
                        df_medicoes,
                        hide_index=True,
                        use_container_width=True,
                    )
        try:
            historico_certificados = listar_analises_certificados()
        except SQLAlchemyError as erro:
            logger.exception("Falha ao consultar histórico compartilhado de certificados.")
            st.error(f"Não foi possível carregar o histórico de certificados: {erro}")
        else:
            st.subheader("Histórico compartilhado de certificados")
            if historico_certificados.empty:
                st.info("Ainda não há análises de certificados salvas.")
            else:
                historico_exibicao = historico_certificados.copy()
                historico_exibicao["Resultado"] = historico_exibicao[
                    "resultado_json"
                ].map(lambda valor: pd.read_json(io.StringIO(valor)).iloc[0].get("Resultado", ""))
                st.dataframe(
                    historico_exibicao.drop(columns=["resultado_json"]),
                    hide_index=True,
                    use_container_width=True,
                )
                analises = {
                    f"#{int(linha.id)} · {linha.Arquivo} · {linha.Resultado}": int(
                        linha.id
                    )
                    for linha in historico_exibicao.itertuples(index=False)
                }
                selecionada = st.selectbox(
                    "Consultar análise salva",
                    list(analises),
                    key="historico_certificado_selecionado",
                )
                try:
                    analise_salva = carregar_analise_certificado(
                        analises[selecionada]
                    )
                except (SQLAlchemyError, ValueError) as erro:
                    logger.exception("Falha ao abrir análise salva de certificado.")
                    st.error(f"Não foi possível abrir a análise: {erro}")
                else:
                    st.json(analise_salva["resultado"])
                    if not analise_salva["medicoes"].empty:
                        st.dataframe(
                            analise_salva["medicoes"],
                            hide_index=True,
                            use_container_width=True,
                        )
                    st.download_button(
                        "Baixar certificado original",
                        analise_salva["conteudo_pdf"],
                        analise_salva["nome_arquivo"],
                        "application/pdf",
                        key=f"download_certificado_{analises[selecionada]}",
                    )
        st.warning(
            "A decisão considera Correção ± U frente aos limites C/D. "
            "Quando o mínimo D não estiver definido, considera-se somente Correção + U "
            "frente ao máximo C. Critérios visuais/sonoros não recebem avaliação numérica; "
            "associação ambígua sem escolha do avaliador, dado ilegível ou limite "
            "não interpretável permanece em revisão."
        )
elif tipo_painel == "Gestão de Calibração":
    render_gestao_calibracao()
elif tipo_painel == "Gestão de CN's":
    try:
        from core.cn_management import render_gestao_cns
    except ModuleNotFoundError as erro:
        if erro.name not in {"core.cn_management", "database.cn_store"}:
            raise
        st.error(
            "O módulo Gestão de CN's não está completo nesta publicação."
        )
        st.info(
            "Publique também os arquivos `core/cn_management.py` e "
            "`database/cn_store.py` junto com `app.py`."
        )
    else:
        render_gestao_cns()
elif tipo_painel == "⚙ Administração":
    render_administracao_global()
elif tipo_painel == "Custos da não Qualidade":
    try:
        colunas_adicionais = None
        if arquivo_usuario is not None:
            df_custos_upload, colunas_adicionais = construir_dataframe_custos(
                arquivo_usuario
            )
            if salvar_upload_se_necessario(
                arquivo_usuario,
                "custos_qualidade",
                df_custos_upload,
            ):
                st.success(
                    f"✓ {len(df_custos_upload)} lançamento(s) adicionados ao histórico."
                )
            else:
                    st.info(
                        "Este conteúdo já está no histórico compartilhado; "
                        "nenhuma linha duplicada foi adicionada."
                    )

        df_custos_historico = carregar_dataframe_cacheado("custos_qualidade")
        if df_custos_historico.empty:
            st.info("Ainda não há custos salvos. Carregue uma planilha para iniciar o histórico.")
        else:
            exibir_painel_custos(df_custos_historico, colunas_adicionais)
    except (
        ValueError,
        KeyError,
        TypeError,
        OSError,
        ImportError,
        sqlite3.Error,
        zipfile.BadZipFile,
    ) as erro:
        st.error(f"Erro ao processar a planilha de custos: {erro}")
elif tipo_painel == "Organograma Semanal":
    if arquivo_usuario is not None:
        try:
            exibir_painel_organograma(arquivo_usuario)
        except (ValueError, KeyError, TypeError, OSError, ImportError, zipfile.BadZipFile) as erro:
            st.error(f"Erro ao processar a planilha do organograma: {erro}")
    else:
        st.caption("Formatos aceitos: Excel (.xlsx, .xlsm, .xlsb).")
elif tipo_painel == "Target Gerência":
    try:
        if arquivo_target is None:
            st.info(
                "Carregue a planilha de origem para recalcular os targets ponderados, "
                "o realizado e o ritmo semanal."
            )
        else:
            target_gerencia = ler_target_gerencia_excel(arquivo_target)
            exibir_painel_target_gerencia(target_gerencia)
    except (ValueError, KeyError, TypeError, OSError, InvalidFileException, zipfile.BadZipFile) as erro:
        st.error(f"Erro ao processar a aba Target Gerência: {erro}")
elif tipo_painel in {
    "Dashboard Operacional",
    "Visão Executiva Q-MAX",
    "Relatórios",
}:
    try:
        mapa_equipe_upload = {}
        if arquivo_usuario is not None and tipo_painel == "Dashboard Operacional":
            try:
                df_bruto, linhas_invalidas = tentar_ler_conteudo_do_arquivo(
                    arquivo_usuario
                )
                mapa_equipe_upload = construir_mapa_equipe_excel(arquivo_usuario)
                df_estrutura = diagnosticar_estrutura_forms(df_bruto)
                df_forms_upload, _ = construir_dataframe_forms(
                    df_bruto,
                    mapa_equipe_upload,
                )
                if df_estrutura.empty:
                    df_suspeitas, colunas_sem_validacao = validar_colunas_forms(
                        df_bruto
                    )
                else:
                    df_suspeitas = pd.DataFrame()
                    colunas_sem_validacao = []

                with st.container(border=True):
                    st.subheader("Validação do arquivo de ações")
                    if df_estrutura.empty:
                        st.success("✓ Estrutura de colunas correta")
                    else:
                        st.warning("⚠️ Possível alteração na estrutura das colunas")
                        with st.expander(
                            "Ver colunas deslocadas ou cabeçalhos suspeitos",
                            expanded=True,
                        ):
                            st.dataframe(
                                df_estrutura,
                                hide_index=True,
                                use_container_width=True,
                            )
                    st.success("✓ Colunas obrigatórias do painel encontradas")
                    if not df_estrutura.empty:
                        st.info(
                            "A validação dos tipos foi pausada para evitar alertas "
                            "incorretos enquanto a estrutura das colunas é revisada."
                        )
                    if colunas_sem_validacao:
                        st.warning(
                            "⚠️ Não foi possível validar as colunas "
                            f"{', '.join(colunas_sem_validacao)}: elas não existem "
                            "na planilha."
                        )
                    if df_suspeitas.empty:
                        st.success("✓ Tipos válidos nas colunas verificadas")
                    else:
                        registros_suspeitos = int(df_suspeitas["Linha"].nunique())
                        st.warning(f"⚠️ {registros_suspeitos} registro(s) suspeito(s)")
                        with st.expander("Ver registros para revisar"):
                            st.dataframe(
                                df_suspeitas,
                                hide_index=True,
                                use_container_width=True,
                            )

                if linhas_invalidas > 0:
                    st.warning(
                        f"⚠️ {linhas_invalidas} linha(s) foram ignoradas por "
                        "inconsistência de estrutura para garantir a apuração correta "
                        "dos indicadores."
                    )
                if salvar_upload_se_necessario(
                    arquivo_usuario,
                    "acoes_qualidade",
                    df_forms_upload,
                ):
                    st.success(
                        f"✓ {len(df_forms_upload)} ação(ões) adicionadas ao histórico."
                    )
                else:
                    st.info(
                        "Este conteúdo já está no histórico compartilhado; "
                        "nenhuma ação duplicada foi adicionada."
                    )
            except (
                ValueError,
                KeyError,
                TypeError,
                OSError,
                ImportError,
                InvalidFileException,
                zipfile.BadZipFile,
            ) as erro:
                st.error(f"Erro ao processar o arquivo de ações: {erro}")

        df_forms_historico = carregar_dataframe_cacheado("acoes_qualidade")
        if df_forms_historico.empty:
            st.info(
                "Ainda não há ações salvas. Carregue uma planilha para iniciar o histórico."
            )
            st.stop()
        df_forms, colunas = construir_dataframe_forms(
            df_forms_historico,
            mapa_equipe_upload,
        )
        st.caption(
            f"Histórico consolidado: {len(df_forms)} ação(ões) carregadas do SQLite."
        )

        meses_disponiveis = sorted(
            [mes for mes in df_forms[colunas["mes"]].dropna().unique() if str(mes).strip()],
            key=lambda mes: MONTH_ORDER.get(normalizar_mes(mes), 999),
        )
        if not meses_disponiveis:
            raise ValueError("Não foi possível identificar o mês da ação no campo MÊS da planilha FORMS.")

        if tipo_painel == "Dashboard Operacional":
            exibir_dashboard_operacional(
                df_forms,
                colunas,
                meses_disponiveis,
            )
        elif tipo_painel == "Visão Executiva Q-MAX":
            exibir_visao_executiva_acoes(
                df_forms,
                colunas,
                meses_disponiveis,
            )
        else:
            exibir_relatorios_acoes(
                df_forms,
                colunas,
                meses_disponiveis,
            )

    except Exception as e:
        st.error(f"Erro ao processar dados do arquivo: {e}")
