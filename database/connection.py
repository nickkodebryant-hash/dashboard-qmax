from contextlib import contextmanager
from datetime import date, datetime
import hashlib
import json
from pathlib import Path
import sqlite3
from io import StringIO
from typing import Any

import pandas as pd
from sqlalchemy import Column, DateTime, Integer, MetaData, String, Table, Text, UniqueConstraint, insert, select

from database.calibration_store import _engine


DATABASE_PATH = Path(__file__).resolve().parent / "qualidade.db"
_DEFAULT_DATABASE_PATH = DATABASE_PATH
TABLES = {
    "acoes_qualidade": "acoes_qualidade",
    "custos_qualidade": "custos_qualidade",
}
_INTERNAL_COLUMNS = {
    "__qmax_id",
    "__qmax_data_upload",
    "__qmax_arquivo_origem",
}
_HISTORICO_METADATA = MetaData()
historicos = Table(
    "qmax_historicos",
    _HISTORICO_METADATA,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("tabela", String(50), nullable=False),
    Column("data_upload", DateTime(timezone=True), nullable=False),
    Column("arquivo_origem", Text, nullable=False),
    Column("conteudo_json", Text, nullable=False),
    Column("fingerprint", String(64), nullable=False),
    UniqueConstraint("tabela", "fingerprint", name="uq_qmax_historico_fingerprint"),
)
_MIGRACOES_METADATA = MetaData()
migracoes_historico = Table(
    "qmax_historicos_migracoes",
    _MIGRACOES_METADATA,
    Column("origem", String(100), primary_key=True),
    Column("migrado_em", DateTime(timezone=True), nullable=False),
)


def _identificador_sql(nome: str) -> str:
    return '"' + nome.replace('"', '""') + '"'


def _validar_tabela(tabela: str) -> str:
    try:
        return TABLES[tabela]
    except KeyError as erro:
        raise ValueError(f"Tabela de histórico inválida: {tabela!r}.") from erro


@contextmanager
def _conectar(caminho_banco: str | Path | None = None):
    caminho = Path(caminho_banco) if caminho_banco else DATABASE_PATH
    caminho.parent.mkdir(parents=True, exist_ok=True)
    conexao = sqlite3.connect(caminho, timeout=30)
    try:
        with conexao:
            yield conexao
    finally:
        conexao.close()


def inicializar_banco(caminho_banco: str | Path | None = None) -> None:
    if caminho_banco is None and DATABASE_PATH == _DEFAULT_DATABASE_PATH:
        engine = _engine()
        _HISTORICO_METADATA.create_all(engine)
        _MIGRACOES_METADATA.create_all(engine)
        _migrar_historicos_sqlite_legados(engine)
        return
    _inicializar_banco_sqlite(caminho_banco)


def _inicializar_banco_sqlite(caminho_banco: str | Path | None = None) -> None:
    with _conectar(caminho_banco) as conexao:
        for tabela in TABLES.values():
            conexao.execute(
                f"""
                CREATE TABLE IF NOT EXISTS "{tabela}" (
                    "__qmax_id" INTEGER PRIMARY KEY AUTOINCREMENT,
                    "__qmax_data_upload" TEXT NOT NULL,
                    "__qmax_arquivo_origem" TEXT NOT NULL
                )
                """
            )


def _normalizar_valor_json(valor: Any) -> Any:
    if valor is None or valor is pd.NA or valor is pd.NaT:
        return None
    if hasattr(valor, "item"):
        valor = valor.item()
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return valor.isoformat()
    if isinstance(valor, (str, int, float, bool)):
        return valor
    return str(valor)


def _json_linhas(linhas: list[dict[str, Any]]) -> str:
    return json.dumps(
        linhas,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=_normalizar_valor_json,
    )


def _fingerprint(tabela: str, linhas_json: str, sal: str = "") -> str:
    return hashlib.sha256(
        f"{tabela}\0{sal}\0{linhas_json}".encode("utf-8")
    ).hexdigest()


def _salvar_no_banco_compartilhado(
    tabela: str,
    dataframe: pd.DataFrame,
    nome_arquivo: str,
    sal_fingerprint: str = "",
) -> int:
    dataframe_para_salvar = dataframe.copy()
    dataframe_para_salvar.columns = [
        str(coluna) for coluna in dataframe_para_salvar.columns
    ]
    nomes_colunas = list(dataframe_para_salvar.columns)
    if len(nomes_colunas) != len(set(nomes_colunas)):
        raise ValueError("O arquivo possui nomes de colunas duplicados.")
    colunas_reservadas = _INTERNAL_COLUMNS.intersection(nomes_colunas)
    if colunas_reservadas:
        raise ValueError(
            "O arquivo utiliza nomes de coluna reservados pelo banco: "
            + ", ".join(sorted(colunas_reservadas))
        )
    dataframe_para_salvar["origem_arquivo"] = str(nome_arquivo)
    linhas = [
        {
            str(chave): _normalizar_valor_json(valor)
            for chave, valor in linha.items()
        }
        for linha in dataframe_para_salvar.to_dict(orient="records")
    ]
    conteudo_json = _json_linhas(linhas)
    fingerprint = _fingerprint(tabela, conteudo_json, sal_fingerprint)
    engine = _engine()
    with engine.begin() as conexao:
        existente = conexao.execute(
            select(historicos.c.id).where(
                historicos.c.tabela == tabela,
                historicos.c.fingerprint == fingerprint,
            )
        ).first()
        if existente is not None:
            return 0
        conexao.execute(
            insert(historicos).values(
                tabela=tabela,
                data_upload=datetime.now().astimezone(),
                arquivo_origem=str(nome_arquivo),
                conteudo_json=conteudo_json,
                fingerprint=fingerprint,
            )
        )
    return len(dataframe_para_salvar)


def _migrar_historicos_sqlite_legados(engine: Any) -> None:
    origem = "database.quality_sqlite.v1"
    with engine.connect() as conexao:
        ja_migrado = conexao.execute(
            select(migracoes_historico.c.origem).where(
                migracoes_historico.c.origem == origem
            )
        ).first()
    if ja_migrado is not None:
        return

    caminho = DATABASE_PATH
    if caminho.exists():
        legado = sqlite3.connect(caminho)
        try:
            tabelas_existentes = {
                linha[0]
                for linha in legado.execute(
                    "SELECT name FROM sqlite_master WHERE type = 'table'"
                )
            }
            with engine.begin() as conexao:
                for tabela in TABLES.values():
                    if tabela not in tabelas_existentes:
                        continue
                    dataframe = pd.read_sql_query(
                        f'SELECT * FROM "{tabela}" ORDER BY "__qmax_id"',
                        legado,
                    )
                    colunas_publicas = [
                        coluna
                        for coluna in dataframe.columns
                        if coluna not in _INTERNAL_COLUMNS
                    ]
                    for _, linha in dataframe.iterrows():
                        dados = {
                            coluna: _normalizar_valor_json(linha[coluna])
                            for coluna in colunas_publicas
                        }
                        conteudo_json = _json_linhas([dados])
                        fingerprint = _fingerprint(
                            tabela,
                            conteudo_json,
                            str(linha.get("__qmax_id", "")),
                        )
                        conexao.execute(
                            insert(historicos).values(
                                tabela=tabela,
                                data_upload=pd.to_datetime(
                                    linha.get("__qmax_data_upload"),
                                    errors="coerce",
                                ).to_pydatetime()
                                if pd.notna(linha.get("__qmax_data_upload"))
                                else datetime.now().astimezone(),
                                arquivo_origem=str(
                                    linha.get("__qmax_arquivo_origem", "")
                                ),
                                conteudo_json=conteudo_json,
                                fingerprint=fingerprint,
                            )
                        )
                conexao.execute(
                    insert(migracoes_historico).values(
                        origem=origem,
                        migrado_em=datetime.now().astimezone(),
                    )
                )
        finally:
            legado.close()
    else:
        with engine.begin() as conexao:
            conexao.execute(
                insert(migracoes_historico).values(
                    origem=origem,
                    migrado_em=datetime.now().astimezone(),
                )
            )


def _tipo_sqlite(coluna: pd.Series) -> str:
    if pd.api.types.is_bool_dtype(coluna.dtype) or pd.api.types.is_integer_dtype(
        coluna.dtype
    ):
        return "INTEGER"
    if pd.api.types.is_float_dtype(coluna.dtype):
        return "REAL"
    return "TEXT"


def _valor_sqlite(valor: Any) -> Any:
    if valor is None or valor is pd.NA or valor is pd.NaT:
        return None
    if hasattr(valor, "item"):
        valor = valor.item()
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return valor.isoformat()
    if isinstance(valor, (str, int, float, bytes)):
        return valor
    if isinstance(valor, bool):
        return int(valor)
    return str(valor)


def salvar_dataframe(
    tabela: str,
    dataframe: pd.DataFrame,
    nome_arquivo: str,
) -> int:
    tabela = _validar_tabela(tabela)
    if dataframe.empty:
        raise ValueError("Não há registros válidos para salvar no histórico.")
    if DATABASE_PATH == _DEFAULT_DATABASE_PATH:
        inicializar_banco()
        return _salvar_no_banco_compartilhado(tabela, dataframe, nome_arquivo)

    dataframe_para_salvar = dataframe.copy()
    dataframe_para_salvar.columns = [
        str(coluna) for coluna in dataframe_para_salvar.columns
    ]
    nomes_colunas = list(dataframe_para_salvar.columns)
    if len(nomes_colunas) != len(set(nomes_colunas)):
        raise ValueError("O arquivo possui nomes de colunas duplicados.")
    colunas_reservadas = _INTERNAL_COLUMNS.intersection(nomes_colunas)
    if colunas_reservadas:
        raise ValueError(
            "O arquivo utiliza nomes de coluna reservados pelo banco: "
            + ", ".join(sorted(colunas_reservadas))
        )

    dataframe_para_salvar["origem_arquivo"] = str(nome_arquivo)
    dataframe_para_salvar["__qmax_data_upload"] = datetime.now().astimezone().isoformat(
        timespec="seconds"
    )
    dataframe_para_salvar["__qmax_arquivo_origem"] = str(nome_arquivo)

    inicializar_banco()
    with _conectar() as conexao:
        colunas_existentes = {
            linha[1]
            for linha in conexao.execute(
                f"PRAGMA table_info({_identificador_sql(tabela)})"
            )
        }
        for nome, serie in dataframe_para_salvar.items():
            if nome not in colunas_existentes:
                conexao.execute(
                    f"ALTER TABLE {_identificador_sql(tabela)} "
                    f"ADD COLUMN {_identificador_sql(nome)} {_tipo_sqlite(serie)}"
                )

        for coluna in dataframe_para_salvar.columns:
            dataframe_para_salvar[coluna] = dataframe_para_salvar[coluna].map(
                _valor_sqlite
            )
        dataframe_para_salvar.to_sql(
            tabela,
            conexao,
            if_exists="append",
            index=False,
        )
    return len(dataframe_para_salvar)


def carregar_dataframe(
    tabela: str,
    caminho_banco: str | Path | None = None,
) -> pd.DataFrame:
    tabela = _validar_tabela(tabela)
    if caminho_banco is None and DATABASE_PATH == _DEFAULT_DATABASE_PATH:
        inicializar_banco()
        with _engine().connect() as conexao:
            historicos_tabela = conexao.execute(
                select(historicos.c.conteudo_json)
                .where(historicos.c.tabela == tabela)
                .order_by(historicos.c.id)
            ).scalars()
            registros = [
                registro
                for conteudo in historicos_tabela
                for registro in json.loads(conteudo)
            ]
        if not registros:
            return pd.DataFrame()
        return pd.read_json(StringIO(_json_linhas(registros)), orient="records")
    inicializar_banco(caminho_banco)
    with _conectar(caminho_banco) as conexao:
        colunas = [
            linha[1]
            for linha in conexao.execute(
                f"PRAGMA table_info({_identificador_sql(tabela)})"
            )
            if linha[1] not in _INTERNAL_COLUMNS
        ]
        if not colunas:
            return pd.DataFrame()
        colunas_sql = ", ".join(_identificador_sql(coluna) for coluna in colunas)
        return pd.read_sql_query(
            f"SELECT {colunas_sql} FROM {_identificador_sql(tabela)} "
            f"ORDER BY {_identificador_sql('__qmax_id')}",
            conexao,
        )


inicializar_banco()
