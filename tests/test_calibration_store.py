from __future__ import annotations

from datetime import date, timedelta
from io import BytesIO
import os
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

from database.calibration_store import (
    DATABASE_URL_SECRET,
    _engine_para_url,
    aprovar_solicitacao,
    autenticar_usuario,
    carregar_auditoria_alteracoes,
    carregar_auditoria_eventos,
    carregar_analise_certificado,
    carregar_equipamentos,
    carregar_historico_equipamento,
    criar_usuario,
    importar_equipamentos,
    inicializar_banco,
    listar_usuarios,
    listar_analises_certificados,
    _migrar_banco_local,
    provisionar_administrador,
    redefinir_senha_administrador,
    registrar_evento,
    salvar_analise_certificado,
    resumir_equipamentos,
)
from core.calibration_management import importar_planilha_cal


def dataframe_equipamentos(fabricante: str = "Acme") -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Nº Equipamento": "E-10060",
                "Descrição": "Paquímetro",
                "Unidade": "mm",
                "Capacidade": "150",
                "Valor Divisão / Resolução": "0,01",
                "Faixa de Uso": "0 a 150",
                "Local de Uso": "Laboratório",
                "Próxima Calibração": pd.Timestamp(date.today() + timedelta(days=30)),
                "Status Calibração": "A VENCER",
                "Data Calibração": pd.Timestamp(date.today()),
                "Setor": "Qualidade",
                "Fabricante": fabricante,
                "Status Equipamento": "ATIVO",
                "FREQUENCIA_CALIBRACAO": "365",
                "Código Normalizado": "E10060",
            }
        ]
    )


class CalibrationStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        database_path = Path(self.temp_dir.name) / "calibracao.db"
        database_url = "sqlite:///" + database_path.as_posix()
        self.database_url = database_url
        self.env_patch = patch.dict(
            os.environ,
            {DATABASE_URL_SECRET: database_url},
        )
        self.env_patch.start()
        inicializar_banco()
        provisionar_administrador(
            "Admin QMAX",
            "admin@example.com",
            "SenhaSegura!123",
        )

    def tearDown(self):
        self.env_patch.stop()
        _engine_para_url(self.database_url).dispose()
        _engine_para_url.cache_clear()
        self.temp_dir.cleanup()

    def test_registration_requires_admin_approval_and_hashes_password(self):
        solicitacao_id = criar_usuario(
            "Pessoa Qualidade",
            "PESSOA@example.com",
            "SenhaSegura!456",
        )
        self.assertIsNone(autenticar_usuario("pessoa@example.com", "SenhaSegura!456"))
        pendentes = listar_usuarios("PENDENTE")
        self.assertEqual(len(pendentes), 1)
        self.assertNotIn("senha_hash", pendentes.columns)

        admin = autenticar_usuario("admin@example.com", "SenhaSegura!123")
        self.assertIsNotNone(admin)
        aprovar_solicitacao(solicitacao_id, "admin@example.com", "QUALIDADE")

        pessoa = autenticar_usuario("pessoa@example.com", "SenhaSegura!456")
        self.assertEqual(pessoa["perfil"], "QUALIDADE")
        self.assertIsNone(autenticar_usuario("pessoa@example.com", "senha errada"))

        eventos = carregar_auditoria_eventos()
        self.assertIn("Usuário aprovado", eventos["Tipo de Evento"].tolist())
        self.assertIn("Tentativa de login inválida", eventos["Tipo de Evento"].tolist())

    def test_approval_and_account_creation_validate_permissions(self):
        request_id = criar_usuario("Visitante", "visitante@example.com", "SenhaSegura!456")
        with self.assertRaises(PermissionError):
            aprovar_solicitacao(request_id, "visitante@example.com", "LEITOR")
        with self.assertRaises(ValueError):
            criar_usuario("Pessoa", "pessoa@example.com", "curta")
        with self.assertRaisesRegex(ValueError, "72 bytes"):
            criar_usuario("Pessoa", "pessoa@example.com", "SenhaSegura!123" * 8)

    def test_initialization_migrates_existing_audit_table(self):
        legacy_path = Path(self.temp_dir.name) / "legacy.db"
        legacy_url = "sqlite:///" + legacy_path.as_posix()
        connection = sqlite3.connect(legacy_path)
        connection.execute(
            """
            CREATE TABLE auditoria_alteracoes (
                id INTEGER PRIMARY KEY,
                data_hora TEXT NOT NULL,
                usuario TEXT NOT NULL,
                equipamento TEXT NOT NULL,
                campo TEXT NOT NULL,
                valor_antigo TEXT NOT NULL,
                valor_novo TEXT NOT NULL,
                tipo_alteracao TEXT NOT NULL
            )
            """
        )
        connection.commit()
        connection.close()
        inicializar_banco(legacy_url)
        from sqlalchemy import create_engine, inspect

        legacy_engine = create_engine(legacy_url)
        columns = {
            column["name"]
            for column in inspect(legacy_engine).get_columns(
                "auditoria_alteracoes"
            )
        }
        self.assertIn("codigo_normalizado", columns)
        legacy_engine.dispose()
        _engine_para_url(legacy_url).dispose()

    def test_local_database_migration_preserves_existing_data_and_is_one_time(self):
        legacy_path = Path(self.temp_dir.name) / "network" / "calibracao.db"
        local_path = Path(self.temp_dir.name) / "local" / "calibracao.db"
        legacy_path.parent.mkdir()
        connection = sqlite3.connect(legacy_path)
        try:
            connection.execute("CREATE TABLE migration_test (value TEXT)")
            connection.execute("INSERT INTO migration_test VALUES ('preserved')")
            connection.commit()
        finally:
            connection.close()

        _migrar_banco_local(legacy_path, local_path)
        _migrar_banco_local(legacy_path, local_path)

        connection = sqlite3.connect(local_path)
        try:
            self.assertEqual(
                connection.execute("SELECT value FROM migration_test").fetchone(),
                ("preserved",),
            )
        finally:
            connection.close()

    def test_admin_password_can_be_reset_and_new_password_authenticates(self):
        redefinir_senha_administrador(
            "admin@example.com",
            "NovaSenhaSegura!789",
        )

        self.assertIsNone(autenticar_usuario("admin@example.com", "SenhaSegura!123"))
        admin = autenticar_usuario("admin@example.com", "NovaSenhaSegura!789")
        self.assertEqual(admin["perfil"], "ADMIN")
        eventos = carregar_auditoria_eventos()
        self.assertIn(
            "Senha de administrador redefinida",
            eventos["Tipo de Evento"].tolist(),
        )

    def test_import_upserts_equipment_and_logs_field_and_event_history(self):
        quality = criar_usuario(
            "Pessoa Qualidade",
            "qualidade@example.com",
            "SenhaSegura!456",
        )
        aprovar_solicitacao(quality, "admin@example.com", "QUALIDADE")

        first = importar_equipamentos(
            dataframe_equipamentos(),
            "qualidade@example.com",
        )
        second = importar_equipamentos(
            dataframe_equipamentos("Fabricante Novo"),
            "qualidade@example.com",
        )
        self.assertEqual(first, {"novos": 1, "atualizados": 0, "inalterados": 0})
        self.assertEqual(second, {"novos": 0, "atualizados": 1, "inalterados": 0})

        equipamentos = carregar_equipamentos()
        self.assertEqual(len(equipamentos), 1)
        self.assertEqual(equipamentos.iloc[0]["Fabricante"], "Fabricante Novo")
        self.assertEqual(equipamentos.iloc[0]["Código Normalizado"], "E10060")

        alteracoes = carregar_auditoria_alteracoes()
        self.assertIn("Cadastro", alteracoes["Campo Alterado"].tolist())
        self.assertIn("Fabricante", alteracoes["Campo Alterado"].tolist())
        self.assertEqual(len(carregar_historico_equipamento("E 10060")), 2)
        registrar_evento(
            "qualidade@example.com",
            "Equipamento excluído",
            "Registro excluído por solicitação administrativa.",
            "E 10060",
        )
        historico = carregar_historico_equipamento("E 10060")
        self.assertEqual(len(historico), 3)
        self.assertIn("Equipamento excluído", historico["Campo"].tolist())
        eventos = carregar_auditoria_eventos()
        self.assertIn("Equipamento cadastrado", eventos["Tipo de Evento"].tolist())
        self.assertEqual(
            eventos["Tipo de Evento"].tolist().count("Importação realizada"),
            2,
        )

    def test_equipment_summary_reads_existing_data_without_initializing_database(self):
        quality = criar_usuario(
            "Pessoa Qualidade",
            "qualidade@example.com",
            "SenhaSegura!456",
        )
        aprovar_solicitacao(quality, "admin@example.com", "QUALIDADE")
        equipment = dataframe_equipamentos()
        equipment.loc[0, "Próxima Calibração"] = pd.Timestamp(
            date.today() - timedelta(days=1)
        )
        importar_equipamentos(equipment, "qualidade@example.com")

        with patch(
            "database.calibration_store.inicializar_banco",
            side_effect=AssertionError("The summary must not initialize or migrate the database."),
        ):
            self.assertEqual(resumir_equipamentos(), (1, 1))

    def test_import_rejects_duplicate_normalized_codes_and_readers(self):
        quality = criar_usuario(
            "Pessoa Qualidade",
            "qualidade@example.com",
            "SenhaSegura!456",
        )
        aprovar_solicitacao(quality, "admin@example.com", "QUALIDADE")
        duplicados = pd.concat(
            [
                dataframe_equipamentos(),
                dataframe_equipamentos().assign(**{"Nº Equipamento": "E 10060"}),
            ],
            ignore_index=True,
        )
        with self.assertRaisesRegex(ValueError, "mais de uma linha"):
            importar_equipamentos(duplicados, "qualidade@example.com")

        leitor = criar_usuario(
            "Leitor",
            "leitor@example.com",
            "SenhaSegura!456",
        )
        aprovar_solicitacao(leitor, "admin@example.com", "LEITOR")
        with self.assertRaises(PermissionError):
            importar_equipamentos(dataframe_equipamentos(), "leitor@example.com")

    def test_certificate_analysis_is_persisted_and_shared(self):
        pdf = b"%PDF-1.4 shared certificate"
        resultado = {
            "Arquivo": "certificado.pdf",
            "Resultado": "Conforme",
            "Equipamento": "E-10060",
            "Critério aplicado": "± 1%",
        }
        medicoes = pd.DataFrame(
            [{"Ponto": "1", "Erro": 0.1, "Resultado": "Conforme"}]
        )

        primeira_gravacao = salvar_analise_certificado(
            "certificado.pdf",
            pdf,
            resultado,
            medicoes,
            "admin@example.com",
        )
        segunda_gravacao = salvar_analise_certificado(
            "certificado.pdf",
            pdf,
            resultado,
            medicoes,
            "admin@example.com",
        )

        self.assertEqual(primeira_gravacao, segunda_gravacao)
        historico = listar_analises_certificados()
        self.assertEqual(len(historico), 1)
        analise = carregar_analise_certificado(primeira_gravacao)
        self.assertEqual(analise["conteudo_pdf"], pdf)
        self.assertEqual(analise["resultado"]["Resultado"], "Conforme")
        self.assertEqual(analise["medicoes"]["Ponto"].tolist(), ["1"])

    def test_importar_planilha_cal_loads_workbook_and_persists_it(self):
        quality = criar_usuario(
            "Pessoa Qualidade",
            "qualidade@example.com",
            "SenhaSegura!456",
        )
        aprovar_solicitacao(quality, "admin@example.com", "QUALIDADE")
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "CAL"
        sheet.append(
            [
                "Nº EQUIPAMENTO",
                "DESCRIÇÃO",
                "PRÓXIMA CALIBRAÇÃO",
                "STATUS CALIBRAÇÃO",
                "FABRICANTE",
                "SETOR",
            ]
        )
        sheet.append(
            [
                "E-20001",
                "Micrômetro",
                date.today() + timedelta(days=20),
                "FORMULA IGNORADA",
                "-",
                "Qualidade",
            ]
        )
        content = BytesIO()
        workbook.save(content)
        result = importar_planilha_cal(
            content.getvalue(),
            "qualidade@example.com",
        )
        self.assertEqual(result, {"novos": 1, "atualizados": 0, "inalterados": 0})
        persisted = carregar_equipamentos()
        self.assertEqual(persisted.loc[0, "Nº Equipamento"], "E-20001")
        self.assertEqual(persisted.loc[0, "Status Calibração"], "A VENCER")
        self.assertEqual(persisted.loc[0, "Fabricante"], "Não Definido")


if __name__ == "__main__":
    unittest.main()
