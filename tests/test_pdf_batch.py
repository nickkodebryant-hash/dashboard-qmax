import time
import unittest
from threading import Barrier

import pandas as pd

from core.pdf_batch import analisar_lote_certificados


class UploadedPdf:
    def __init__(self, name, content=b"pdf"):
        self.name = name
        self.content = content
        self.size = len(content)

    def getvalue(self):
        return self.content


def fake_analyzer(arquivo, criterios, criterio_manual=None):
    if arquivo.name == "slow.pdf":
        time.sleep(0.05)
    return (
        {
            "Arquivo": arquivo.name,
            "Resultado": "Revisão necessária",
            "Critério aplicado": criterio_manual or "",
        },
        pd.DataFrame([{"Arquivo": arquivo.name}]),
    )


class CertificateBatchTests(unittest.TestCase):
    def test_batch_runs_multiple_workers_at_the_same_time(self):
        barreira = Barrier(2)

        def analyzer(arquivo, criterios, criterio_manual=None):
            barreira.wait(timeout=2)
            return fake_analyzer(arquivo, criterios, criterio_manual)

        resultados, _ = analisar_lote_certificados(
            [UploadedPdf("first.pdf"), UploadedPdf("second.pdf")],
            pd.DataFrame(),
            analyzer,
        )

        self.assertTrue(
            resultados["Resultado"].eq("Revisão necessária").all()
        )

    def test_batch_preserves_input_order_and_reports_progress(self):
        arquivos = [UploadedPdf("slow.pdf"), UploadedPdf("fast.pdf")]
        progresso = []

        resultados, detalhes = analisar_lote_certificados(
            arquivos,
            pd.DataFrame(),
            fake_analyzer,
            ao_progresso=lambda concluidas, total, nome: progresso.append(
                (concluidas, total, nome)
            ),
        )

        self.assertEqual(resultados["Arquivo"].tolist(), ["slow.pdf", "fast.pdf"])
        self.assertEqual(set(detalhes), {"1. slow.pdf", "2. fast.pdf"})
        self.assertEqual(len(progresso), 2)
        self.assertEqual(progresso[-1][0:2], (2, 2))

    def test_manual_criterion_reuses_other_precalculated_results(self):
        arquivos = [UploadedPdf("first.pdf"), UploadedPdf("second.pdf")]
        chamadas = []

        def analyzer(arquivo, criterios, criterio_manual=None):
            chamadas.append((arquivo.name, criterio_manual))
            return fake_analyzer(arquivo, criterios, criterio_manual)

        resultados_iniciais, detalhes_iniciais = analisar_lote_certificados(
            arquivos,
            pd.DataFrame(),
            analyzer,
        )
        chamadas.clear()

        resultados_finais, detalhes_finais = analisar_lote_certificados(
            arquivos,
            pd.DataFrame(),
            analyzer,
            criterios_manuais={1: 17},
            resultados_precalculados=(resultados_iniciais, detalhes_iniciais),
        )

        self.assertEqual(chamadas, [("second.pdf", 17)])
        self.assertEqual(resultados_finais["Critério aplicado"].tolist(), ["", 17])
        self.assertEqual(set(detalhes_finais), {"1. first.pdf", "2. second.pdf"})

    def test_unexpected_pdf_failure_is_reported_without_aborting_batch(self):
        def analyzer(arquivo, criterios, criterio_manual=None):
            if arquivo.name == "broken.pdf":
                raise ValueError("PDF inválido")
            return fake_analyzer(arquivo, criterios, criterio_manual)

        with self.assertLogs("core.pdf_batch", level="ERROR"):
            resultados, _ = analisar_lote_certificados(
                [UploadedPdf("ok.pdf"), UploadedPdf("broken.pdf")],
                pd.DataFrame(),
                analyzer,
            )

        falha = resultados.loc[resultados["Arquivo"] == "broken.pdf"].iloc[0]
        self.assertEqual(
            falha["Resultado"],
            "Falha ao processar — erro inesperado",
        )
        self.assertIn("PDF inválido", falha["Detalhe"])
