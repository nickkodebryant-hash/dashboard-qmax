import unittest

from core.home_navigation import (
    DESTINOS_VALIDOS,
    GRUPOS_FERRAMENTAS,
    PAINEIS,
    PAINEIS_POR_NOME,
)
from core.text_utils import chave_coluna, limpar_texto, normalizar_mes


class AppUtilityTests(unittest.TestCase):
    def test_text_helpers_preserve_existing_normalization(self):
        self.assertEqual(limpar_texto("  Calibração  "), "CALIBRACAO")
        self.assertEqual(chave_coluna(" Código-da Linha "), "CODIGO DA LINHA")
        self.assertEqual(normalizar_mes("jan"), "JANEIRO")

    def test_home_navigation_groups_keep_their_existing_destinations(self):
        self.assertEqual(
            GRUPOS_FERRAMENTAS["grupo_confiabilidade"]["paineis"],
            ("Organograma Semanal", "Gestão de CN's"),
        )
        self.assertEqual(
            GRUPOS_FERRAMENTAS["grupo_calibracao"]["paineis"],
            ("Gestão de Calibração", "Certificados de Calibração"),
        )
        self.assertEqual(set(PAINEIS_POR_NOME), {painel[0] for painel in PAINEIS})
        self.assertTrue(
            {"grupo_confiabilidade", "grupo_calibracao"}.issubset(DESTINOS_VALIDOS)
        )


if __name__ == "__main__":
    unittest.main()
