from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from dataclasses import dataclass
import logging
from typing import Callable

import pandas as pd

from config.settings import (
    CERTIFICATE_WORKERS,
    MAX_CERTIFICATE_BATCH_MB,
    MAX_CERTIFICATE_FILES,
)

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ArquivoCertificado:
    name: str
    conteudo: bytes

    @property
    def size(self):
        return len(self.conteudo)

    def getvalue(self):
        return self.conteudo


def _resultado_lote_nao_analisado(arquivo, resultado, detalhe):
    return {
        "Arquivo": arquivo.name,
        "Resultado": resultado,
        "Equipamento": "",
        "Tipo": "",
        "Faixa nominal": "",
        "Unidade": "",
        "Critério aplicado": "",
        "Identificação": "",
        "Certificado": "",
        "Medições extraídas": 0,
        "Detalhe": detalhe,
    }


def analisar_lote_certificados(
    arquivos,
    criterios,
    analisar_pdf: Callable,
    criterios_manuais=None,
    ao_progresso: Callable[[int, int, str], None] | None = None,
    resultados_precalculados=None,
):
    if not arquivos:
        return pd.DataFrame(), {}

    if len(arquivos) > MAX_CERTIFICATE_FILES:
        mensagem = f"Envie no máximo {MAX_CERTIFICATE_FILES} PDFs por lote."
        return (
            pd.DataFrame(
                [
                    _resultado_lote_nao_analisado(
                        arquivo,
                        "Lote não analisado — quantidade máxima excedida",
                        mensagem,
                    )
                    for arquivo in arquivos
                ]
            ),
            {},
        )

    tamanho_lote = sum(arquivo.size for arquivo in arquivos)
    if tamanho_lote > MAX_CERTIFICATE_BATCH_MB * 1024 * 1024:
        mensagem = f"Limite total do lote: {MAX_CERTIFICATE_BATCH_MB} MB."
        return (
            pd.DataFrame(
                [
                    _resultado_lote_nao_analisado(
                        arquivo,
                        "Lote não analisado — tamanho total excedido",
                        mensagem,
                    )
                    for arquivo in arquivos
                ]
            ),
            {},
        )

    criterios_manuais = criterios_manuais or {}
    resultados = [None] * len(arquivos)
    detalhes = {}
    tarefas = {}

    if resultados_precalculados is not None:
        resultados_anteriores, detalhes_anteriores = resultados_precalculados
        if len(resultados_anteriores) == len(arquivos):
            for indice, arquivo in enumerate(arquivos):
                if indice in criterios_manuais:
                    continue
                resultados[indice] = resultados_anteriores.iloc[indice].to_dict()
                chave = f"{indice + 1}. {arquivo.name}"
                detalhes[chave] = detalhes_anteriores.get(chave, pd.DataFrame())

    for indice, arquivo in enumerate(arquivos):
        if resultados[indice] is not None:
            continue
        linha_criterio_manual = criterios_manuais.get(indice)
        snapshot = ArquivoCertificado(arquivo.name, arquivo.getvalue())
        tarefas[indice] = (snapshot, linha_criterio_manual)

    total_tarefas = len(tarefas)
    if not total_tarefas:
        return pd.DataFrame(resultados), detalhes

    with ThreadPoolExecutor(
        max_workers=min(CERTIFICATE_WORKERS, total_tarefas),
        thread_name_prefix="qmax-pdf",
    ) as executor:
        futuros: dict[Future, int] = {
            executor.submit(
                analisar_pdf,
                snapshot,
                criterios,
                linha_criterio_manual,
            ): indice
            for indice, (snapshot, linha_criterio_manual) in tarefas.items()
        }
        concluidas = 0
        for futuro in as_completed(futuros):
            indice = futuros[futuro]
            arquivo, _ = tarefas[indice]
            chave = f"{indice + 1}. {arquivo.name}"
            try:
                resultado, medicoes = futuro.result()
            except Exception as erro:
                logger.exception("Falha inesperada ao analisar %s", arquivo.name)
                resultado = _resultado_lote_nao_analisado(
                    arquivo,
                    "Falha ao processar — erro inesperado",
                    f"Erro durante a análise deste PDF: {erro}",
                )
                medicoes = pd.DataFrame()
            resultados[indice] = resultado
            detalhes[chave] = medicoes
            concluidas += 1
            if ao_progresso is not None:
                ao_progresso(concluidas, total_tarefas, arquivo.name)

    return pd.DataFrame(resultados), detalhes
