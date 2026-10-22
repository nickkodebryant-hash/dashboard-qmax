import inspect
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from unittest.mock import patch

import database.connection as database
from database.calibration_store import (
    DATABASE_URL_SECRET,
    _engine_para_url,
)
from database.connection import carregar_dataframe, inicializar_banco, salvar_dataframe


class DatabasePersistenceTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temp_dir.name) / "qualidade.db"
        self.database_path_original = database.DATABASE_PATH
        database.DATABASE_PATH = self.database_path
        inicializar_banco(self.database_path)

    def tearDown(self):
        database.DATABASE_PATH = self.database_path_original
        self.temp_dir.cleanup()

    def test_initializes_both_history_tables(self):
        connection = sqlite3.connect(self.database_path)
        try:
            tables = {
                row[0]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }
        finally:
            connection.close()
        self.assertTrue({"acoes_qualidade", "custos_qualidade"}.issubset(tables))

    def test_save_function_accepts_exactly_three_arguments(self):
        self.assertEqual(
            list(inspect.signature(salvar_dataframe).parameters),
            ["tabela", "dataframe", "nome_arquivo"],
        )

    def test_appends_rows_and_preserves_numeric_values(self):
        first_upload = pd.DataFrame({"Projeto": ["A"], "Valor": [12.5]})
        second_upload = pd.DataFrame({"Projeto": ["B"], "Valor": [7.0]})

        self.assertEqual(
            salvar_dataframe(
                "custos_qualidade",
                first_upload,
                "custos.xlsx",
            ),
            1,
        )
        salvar_dataframe(
            "custos_qualidade",
            second_upload,
            "custos-2.xlsx",
        )

        history = carregar_dataframe("custos_qualidade")
        self.assertEqual(history["Projeto"].tolist(), ["A", "B"])
        self.assertEqual(history["Valor"].tolist(), [12.5, 7.0])
        self.assertEqual(
            history["origem_arquivo"].tolist(),
            ["custos.xlsx", "custos-2.xlsx"],
        )

    def test_adds_new_columns_without_losing_existing_history(self):
        salvar_dataframe(
            "acoes_qualidade",
            pd.DataFrame({"Status": ["Aberto"]}),
            "acoes.xlsx",
        )
        salvar_dataframe(
            "acoes_qualidade",
            pd.DataFrame({"Status": ["Concluído"], "Mês": ["Outubro"]}),
            "acoes-outubro.xlsx",
        )

        history = carregar_dataframe("acoes_qualidade")
        self.assertEqual(history["Status"].tolist(), ["Aberto", "Concluído"])
        self.assertTrue(pd.isna(history.loc[0, "Mês"]))
        self.assertEqual(history.loc[1, "Mês"], "Outubro")
        self.assertEqual(
            history["origem_arquivo"].tolist(),
            ["acoes.xlsx", "acoes-outubro.xlsx"],
        )

    def test_rejects_internal_column_names(self):
        with self.assertRaisesRegex(ValueError, "reservados"):
            salvar_dataframe(
                "acoes_qualidade",
                pd.DataFrame({"__qmax_id": [1]}),
                "acoes.xlsx",
            )

    def test_shared_history_is_visible_across_sessions_and_deduplicated(self):
        database_url = (
            "sqlite:///" + (Path(self.temp_dir.name) / "shared.db").as_posix()
        )
        self.database_path_original = database.DATABASE_PATH
        database.DATABASE_PATH = database._DEFAULT_DATABASE_PATH
        with patch.dict(
            os.environ,
            {DATABASE_URL_SECRET: database_url},
        ), patch("database.connection._migrar_historicos_sqlite_legados"):
            _engine_para_url.cache_clear()
            try:
                dados = pd.DataFrame(
                    {"Número da Ação": ["A-1"], "Status": ["Aberto"]}
                )
                self.assertEqual(
                    salvar_dataframe("acoes_qualidade", dados, "acoes.xlsx"),
                    1,
                )
                self.assertEqual(
                    salvar_dataframe("acoes_qualidade", dados, "acoes.xlsx"),
                    0,
                )
                outra_sessao = carregar_dataframe("acoes_qualidade")
                self.assertEqual(outra_sessao["Número da Ação"].tolist(), ["A-1"])
                self.assertEqual(outra_sessao["origem_arquivo"].tolist(), ["acoes.xlsx"])
            finally:
                _engine_para_url(database_url).dispose()
                _engine_para_url.cache_clear()

    def test_migrates_legacy_sqlite_history_into_shared_store_once(self):
        legado = Path(self.temp_dir.name) / "quality-legacy.db"
        conexao_legado = sqlite3.connect(legado)
        try:
            conexao_legado.execute(
                'CREATE TABLE acoes_qualidade ('
                '"__qmax_id" INTEGER PRIMARY KEY, '
                '"__qmax_data_upload" TEXT NOT NULL, '
                '"__qmax_arquivo_origem" TEXT NOT NULL, '
                '"Ação" TEXT, "Status" TEXT)'
            )
            conexao_legado.execute(
                "INSERT INTO acoes_qualidade VALUES (?, ?, ?, ?, ?)",
                (1, "2026-10-01T10:00:00+00:00", "legado.xlsx", "Q-1", "Aberto"),
            )
            conexao_legado.commit()
        finally:
            conexao_legado.close()
        database_url = (
            "sqlite:///" + (Path(self.temp_dir.name) / "shared-migrated.db").as_posix()
        )
        self.database_path_original = database.DATABASE_PATH
        self.database_default_original = database._DEFAULT_DATABASE_PATH
        database.DATABASE_PATH = legado
        database._DEFAULT_DATABASE_PATH = legado
        with patch.dict(
            os.environ,
            {DATABASE_URL_SECRET: database_url},
        ):
            _engine_para_url.cache_clear()
            try:
                inicializar_banco()
                inicializar_banco()
                historico = carregar_dataframe("acoes_qualidade")
                self.assertEqual(historico["Ação"].tolist(), ["Q-1"])
                self.assertEqual(historico["Status"].tolist(), ["Aberto"])
                self.assertEqual(len(historico), 1)
            finally:
                database.DATABASE_PATH = self.database_path_original
                database._DEFAULT_DATABASE_PATH = self.database_default_original
                _engine_para_url(database_url).dispose()
                _engine_para_url.cache_clear()


if __name__ == "__main__":
    unittest.main()
