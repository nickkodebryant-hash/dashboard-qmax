from __future__ import annotations

from datetime import date, datetime, timezone
from functools import lru_cache
import hashlib
import json
import logging
import os
from pathlib import Path
import re
import sqlite3
import tempfile
from typing import Any

import bcrypt
import pandas as pd
from sqlalchemy import (
    Column,
    DateTime,
    Integer,
    LargeBinary,
    MetaData,
    String,
    Table,
    Text,
    create_engine,
    inspect,
    insert,
    select,
    update,
)
from sqlalchemy.engine import Engine
from sqlalchemy.exc import IntegrityError
from streamlit.errors import StreamlitSecretNotFoundError
import streamlit as st


LEGACY_DATABASE_PATH = Path(__file__).resolve().parent / "calibracao.db"
DATABASE_PATH = (
    Path(os.environ.get("LOCALAPPDATA", Path.home() / ".local" / "share"))
    / "QMAX"
    / "calibracao.db"
)
DATABASE_URL_SECRET = "QMAX_CALIBRACAO_DATABASE_URL"
PERFIS_USUARIO = ("ADMIN", "QUALIDADE", "LEITOR")
STATUS_USUARIO = ("PENDENTE", "ATIVO", "REJEITADO")
PAPEIS_AUDITORIA = {"ADMIN", "QUALIDADE"}
logger = logging.getLogger(__name__)

COLUNAS_EQUIPAMENTO = {
    "numero_equipamento": "Nº Equipamento",
    "descricao": "Descrição",
    "unidade": "Unidade",
    "capacidade": "Capacidade",
    "valor_divisao_resolucao": "Valor Divisão / Resolução",
    "faixa_uso": "Faixa de Uso",
    "local_uso": "Local de Uso",
    "proxima_calibracao": "Próxima Calibração",
    "status_calibracao": "Status Calibração",
    "data_calibracao": "Data Calibração",
    "setor": "Setor",
    "fabricante": "Fabricante",
    "status_equipamento": "Status Equipamento",
    "frequencia_calibracao": "FREQUENCIA_CALIBRACAO",
}

metadata = MetaData()

usuarios = Table(
    "cal_usuarios",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("nome", String(200), nullable=False),
    Column("email", String(320), nullable=False, unique=True),
    Column("senha_hash", String(200), nullable=False),
    Column("perfil", String(30), nullable=False),
    Column("status", String(20), nullable=False),
    Column("criado_em", DateTime(timezone=True), nullable=False),
    Column("aprovado_por", String(320)),
    Column("aprovado_em", DateTime(timezone=True)),
)

equipamentos = Table(
    "cal_equipamentos",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("codigo_normalizado", String(200), nullable=False, unique=True),
    *[
        Column(nome, Text)
        for nome in COLUNAS_EQUIPAMENTO
    ],
    Column("atualizado_em", DateTime(timezone=True), nullable=False),
)

certificados = Table(
    "cal_certificados",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("chave_analise", String(64), nullable=False, unique=True),
    Column("hash_pdf", String(64), nullable=False),
    Column("nome_arquivo", Text, nullable=False),
    Column("conteudo_pdf", LargeBinary, nullable=False),
    Column("resultado_json", Text, nullable=False),
    Column("medicoes_json", Text, nullable=False),
    Column("usuario", String(320), nullable=False),
    Column("criado_em", DateTime(timezone=True), nullable=False),
)

auditoria_alteracoes = Table(
    "auditoria_alteracoes",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("data_hora", DateTime(timezone=True), nullable=False),
    Column("usuario", String(320), nullable=False),
    Column("equipamento", String(200), nullable=False),
    Column("codigo_normalizado", String(200), nullable=False, default=""),
    Column("campo", String(200), nullable=False),
    Column("valor_antigo", Text, nullable=False),
    Column("valor_novo", Text, nullable=False),
    Column("tipo_alteracao", String(30), nullable=False),
)

auditoria_eventos = Table(
    "auditoria_eventos",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("data_hora", DateTime(timezone=True), nullable=False),
    Column("usuario", String(320), nullable=False),
    Column("tipo_evento", String(100), nullable=False),
    Column("descricao", Text, nullable=False),
    Column("equipamento", String(200), nullable=False, default=""),
    Column("codigo_normalizado", String(200), nullable=False, default=""),
)


def _url_configurada() -> str:
    url = os.environ.get(DATABASE_URL_SECRET)
    if not url:
        try:
            url = st.secrets.get(DATABASE_URL_SECRET)
        except StreamlitSecretNotFoundError:
            url = None
    if not url:
        _migrar_banco_local(LEGACY_DATABASE_PATH, DATABASE_PATH)
        return f"sqlite:///{DATABASE_PATH.resolve().as_posix()}"
    if url.startswith("postgres://"):
        url = "postgresql+psycopg://" + url[len("postgres://") :]
    elif url.startswith("postgresql://"):
        url = "postgresql+psycopg://" + url[len("postgresql://") :]
    return url


def _migrar_banco_local(origem: Path, destino: Path) -> None:
    if destino.exists() or not origem.exists() or origem.resolve() == destino.resolve():
        return
    destino.parent.mkdir(parents=True, exist_ok=True)
    arquivo_temporario = tempfile.NamedTemporaryFile(
        prefix="calibracao-",
        suffix=".db",
        dir=destino.parent,
        delete=False,
    )
    caminho_temporario = Path(arquivo_temporario.name)
    arquivo_temporario.close()
    banco_origem: sqlite3.Connection | None = None
    banco_destino: sqlite3.Connection | None = None
    try:
        banco_origem = sqlite3.connect(str(origem), timeout=30)
        banco_destino = sqlite3.connect(str(caminho_temporario), timeout=30)
        banco_origem.backup(banco_destino)
        banco_destino.close()
        banco_destino = None
        banco_origem.close()
        banco_origem = None
        os.replace(caminho_temporario, destino)
    except (OSError, sqlite3.Error) as erro:
        logger.exception(
            "Falha ao migrar o banco de calibração para o armazenamento local."
        )
        raise OSError(
            "Não foi possível copiar o banco de calibração da unidade de rede "
            "para o armazenamento local."
        ) from erro
    finally:
        if banco_destino is not None:
            banco_destino.close()
        if banco_origem is not None:
            banco_origem.close()
        caminho_temporario.unlink(missing_ok=True)
    logger.info("Banco de calibração migrado para armazenamento local: %s", destino)


@lru_cache(maxsize=4)
def _engine_para_url(url: str) -> Engine:
    if url.startswith("sqlite:///") and ":memory:" not in url:
        caminho = url[len("sqlite:///") :]
        Path(caminho).parent.mkdir(parents=True, exist_ok=True)
        return create_engine(
            url,
            connect_args={"check_same_thread": False, "timeout": 30},
        )
    if url.endswith(":memory:"):
        return create_engine(
            url,
            connect_args={"check_same_thread": False},
        )
    return create_engine(url, pool_pre_ping=True)


def _engine(database_url: str | None = None) -> Engine:
    return _engine_para_url(database_url or _url_configurada())


def inicializar_banco(database_url: str | None = None) -> None:
    engine = _engine(database_url)
    metadata.create_all(engine)
    colunas_auditoria = {
        coluna["name"]
        for coluna in inspect(engine).get_columns("auditoria_alteracoes")
    }
    if "codigo_normalizado" not in colunas_auditoria:
        with engine.begin() as conexao:
            conexao.exec_driver_sql(
                "ALTER TABLE auditoria_alteracoes "
                "ADD COLUMN codigo_normalizado VARCHAR(200) NOT NULL DEFAULT ''"
            )
    colunas_eventos = {
        coluna["name"]
        for coluna in inspect(engine).get_columns("auditoria_eventos")
    }
    if "codigo_normalizado" not in colunas_eventos:
        with engine.begin() as conexao:
            conexao.exec_driver_sql(
                "ALTER TABLE auditoria_eventos "
                "ADD COLUMN codigo_normalizado VARCHAR(200) NOT NULL DEFAULT ''"
            )


def _agora() -> datetime:
    return datetime.now(timezone.utc).replace(microsecond=0)


def _normalizar_email(email: str) -> str:
    return email.strip().casefold()


def _validar_email(email: str) -> str:
    email_normalizado = _normalizar_email(email)
    if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email_normalizado):
        raise ValueError("Informe um endereço de e-mail válido.")
    return email_normalizado


def normalizar_codigo(codigo: Any) -> str:
    if codigo is None or pd.isna(codigo):
        return ""
    return re.sub(r"[\s-]+", "", str(codigo)).upper()


def _hash_senha(senha: str) -> str:
    senha_bytes = senha.encode("utf-8")
    if len(senha) < 12:
        raise ValueError("A senha deve ter pelo menos 12 caracteres.")
    if len(senha_bytes) > 72:
        raise ValueError("A senha não pode exceder 72 bytes.")
    return bcrypt.hashpw(
        senha_bytes,
        bcrypt.gensalt(rounds=12),
    ).decode("ascii")


def _registrar_evento(
    conexao: Any,
    usuario: str,
    tipo_evento: str,
    descricao: str,
    equipamento: str = "",
) -> None:
    conexao.execute(
        insert(auditoria_eventos).values(
            data_hora=_agora(),
            usuario=usuario,
            tipo_evento=tipo_evento,
            descricao=descricao,
            equipamento=equipamento,
            codigo_normalizado=normalizar_codigo(equipamento),
        )
    )


def registrar_evento(
    usuario: str,
    tipo_evento: str,
    descricao: str,
    equipamento: str = "",
    database_url: str | None = None,
) -> None:
    if not tipo_evento.strip() or not descricao.strip():
        raise ValueError("Tipo e descrição do evento são obrigatórios.")
    with _engine(database_url).begin() as conexao:
        _registrar_evento(
            conexao,
            usuario.strip() or "sistema",
            tipo_evento.strip(),
            descricao.strip(),
            equipamento.strip(),
        )


def registrar_auditoria(
    usuario: str,
    equipamento: str,
    campo: str,
    valor_antigo: Any,
    valor_novo: Any,
    tipo_alteracao: str = "ALTERACAO",
    database_url: str | None = None,
) -> None:
    if not campo.strip() or not equipamento.strip():
        raise ValueError("Equipamento e campo são obrigatórios na auditoria.")
    codigo_normalizado = normalizar_codigo(equipamento)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE"})
        conexao.execute(
            insert(auditoria_alteracoes).values(
                data_hora=_agora(),
                usuario=actor["email"],
                equipamento=equipamento.strip(),
                codigo_normalizado=codigo_normalizado,
                campo=campo.strip(),
                valor_antigo="" if valor_antigo is None else str(valor_antigo),
                valor_novo="" if valor_novo is None else str(valor_novo),
                tipo_alteracao=tipo_alteracao.strip().upper(),
            )
        )
        _registrar_evento(
            conexao,
            actor["email"],
            tipo_alteracao.strip().upper(),
            f"{campo.strip()} alterado de "
            f"'{'' if valor_antigo is None else valor_antigo}' para "
            f"'{'' if valor_novo is None else valor_novo}'.",
            equipamento.strip(),
        )


def _exigir_perfil(
    conexao: Any,
    email: str,
    perfis: set[str],
) -> Any:
    usuario = conexao.execute(
        select(usuarios).where(usuarios.c.email == _normalizar_email(email))
    ).mappings().first()
    if usuario is None or usuario["status"] != "ATIVO":
        raise PermissionError("Usuário não autenticado ou inativo.")
    if usuario["perfil"] not in perfis:
        raise PermissionError("Seu perfil não tem permissão para esta operação.")
    return usuario


def criar_usuario(nome: str, email: str, senha: str) -> int:
    nome = nome.strip()
    email = _validar_email(email)
    if not nome:
        raise ValueError("Informe o nome do usuário.")
    inicializar_banco()
    hash_senha = _hash_senha(senha)
    with _engine().begin() as conexao:
        try:
            resultado = conexao.execute(
                insert(usuarios).values(
                    nome=nome,
                    email=email,
                    senha_hash=hash_senha,
                    perfil="LEITOR",
                    status="PENDENTE",
                    criado_em=_agora(),
                )
            )
        except IntegrityError as erro:
            raise ValueError("Já existe uma conta ou solicitação para esse e-mail.") from erro
        _registrar_evento(
            conexao,
            email,
            "Solicitação de acesso",
            f"Solicitação de acesso criada para {email}.",
        )
        return int(resultado.inserted_primary_key[0])


def provisionar_administrador(
    nome: str,
    email: str,
    senha: str,
    database_url: str | None = None,
) -> None:
    nome = nome.strip()
    email = _validar_email(email)
    if not nome:
        raise ValueError("Informe o nome do administrador.")
    inicializar_banco(database_url)
    hash_senha = _hash_senha(senha)
    with _engine(database_url).begin() as conexao:
        if conexao.execute(select(usuarios.c.id).limit(1)).first() is not None:
            raise ValueError(
                "Já existe um usuário. O administrador inicial só pode ser criado "
                "quando o cadastro estiver vazio."
            )
        conexao.execute(
            insert(usuarios).values(
                nome=nome,
                email=email,
                senha_hash=hash_senha,
                perfil="ADMIN",
                status="ATIVO",
                criado_em=_agora(),
                aprovado_por=email,
                aprovado_em=_agora(),
            )
        )
        _registrar_evento(
            conexao,
            email,
            "Administrador inicial provisionado",
            f"Conta de administrador inicial criada para {email}.",
        )


def redefinir_senha_administrador(email: str, senha: str) -> None:
    email = _validar_email(email)
    hash_senha = _hash_senha(senha)
    inicializar_banco()
    with _engine().begin() as conexao:
        administrador = conexao.execute(
            select(usuarios.c.id).where(
                usuarios.c.email == email,
                usuarios.c.perfil == "ADMIN",
                usuarios.c.status == "ATIVO",
            )
        ).first()
        if administrador is None:
            raise ValueError("Não foi encontrado um administrador ativo com esse e-mail.")
        conexao.execute(
            update(usuarios)
            .where(usuarios.c.id == administrador.id)
            .values(senha_hash=hash_senha)
        )
        _registrar_evento(
            conexao,
            email,
            "Senha de administrador redefinida",
            "A senha do administrador foi redefinida pelo utilitário local.",
        )


def autenticar_usuario(email: str, senha: str) -> dict[str, Any] | None:
    email = _validar_email(email)
    inicializar_banco()
    with _engine().begin() as conexao:
        usuario = conexao.execute(
            select(usuarios).where(usuarios.c.email == email)
        ).mappings().first()
        senha_valida = (
            usuario is not None
            and len(senha.encode("utf-8")) <= 72
            and bcrypt.checkpw(
                senha.encode("utf-8"),
                usuario["senha_hash"].encode("ascii"),
            )
        )
        if usuario is None or not senha_valida or usuario["status"] != "ATIVO":
            _registrar_evento(
                conexao,
                email,
                "Tentativa de login inválida",
                "Credenciais inválidas ou solicitação ainda não aprovada.",
            )
            return None
        _registrar_evento(conexao, email, "Login realizado", "Autenticação concluída.")
        return {
            "id": usuario["id"],
            "nome": usuario["nome"],
            "email": usuario["email"],
            "perfil": usuario["perfil"],
        }


def usuario_ativo(email: str) -> dict[str, Any] | None:
    inicializar_banco()
    with _engine().connect() as conexao:
        usuario = conexao.execute(
            select(usuarios).where(
                usuarios.c.email == _normalizar_email(email),
                usuarios.c.status == "ATIVO",
            )
        ).mappings().first()
    if usuario is None:
        return None
    return {
        "id": usuario["id"],
        "nome": usuario["nome"],
        "email": usuario["email"],
        "perfil": usuario["perfil"],
    }


def listar_usuarios(status: str | None = None) -> pd.DataFrame:
    inicializar_banco()
    consulta = select(
        usuarios.c.id,
        usuarios.c.nome,
        usuarios.c.email,
        usuarios.c.perfil,
        usuarios.c.status,
        usuarios.c.criado_em,
        usuarios.c.aprovado_por,
        usuarios.c.aprovado_em,
    )
    if status is not None:
        if status not in STATUS_USUARIO:
            raise ValueError("Status de usuário inválido.")
        consulta = consulta.where(usuarios.c.status == status)
    with _engine().connect() as conexao:
        return pd.read_sql(consulta.order_by(usuarios.c.criado_em), conexao)


def aprovar_solicitacao(
    solicitacao_id: int,
    aprovador: str,
    perfil: str,
) -> None:
    perfil = perfil.strip().upper()
    if perfil not in PERFIS_USUARIO:
        raise ValueError("Perfil de usuário inválido.")
    with _engine().begin() as conexao:
        admin = _exigir_perfil(conexao, aprovador, {"ADMIN"})
        solicitacao = conexao.execute(
            select(usuarios).where(
                usuarios.c.id == solicitacao_id,
                usuarios.c.status == "PENDENTE",
            )
        ).mappings().first()
        if solicitacao is None:
            raise ValueError("A solicitação não existe ou já foi processada.")
        agora = _agora()
        conexao.execute(
            update(usuarios)
            .where(usuarios.c.id == solicitacao_id)
            .values(
                perfil=perfil,
                status="ATIVO",
                aprovado_por=admin["email"],
                aprovado_em=agora,
            )
        )
        _registrar_evento(
            conexao,
            admin["email"],
            "Usuário aprovado",
            f"Usuário {solicitacao['email']} aprovado com perfil {perfil}.",
        )


def rejeitar_solicitacao(solicitacao_id: int, aprovador: str) -> None:
    with _engine().begin() as conexao:
        admin = _exigir_perfil(conexao, aprovador, {"ADMIN"})
        solicitacao = conexao.execute(
            select(usuarios).where(
                usuarios.c.id == solicitacao_id,
                usuarios.c.status == "PENDENTE",
            )
        ).mappings().first()
        if solicitacao is None:
            raise ValueError("A solicitação não existe ou já foi processada.")
        conexao.execute(
            update(usuarios)
            .where(usuarios.c.id == solicitacao_id)
            .values(status="REJEITADO")
        )
        _registrar_evento(
            conexao,
            admin["email"],
            "Solicitação rejeitada",
            f"Solicitação de {solicitacao['email']} rejeitada.",
        )


def alterar_perfil_usuario(usuario_id: int, admin_email: str, perfil: str) -> None:
    perfil = perfil.strip().upper()
    if perfil not in PERFIS_USUARIO:
        raise ValueError("Perfil de usuário inválido.")
    with _engine().begin() as conexao:
        admin = _exigir_perfil(conexao, admin_email, {"ADMIN"})
        usuario = conexao.execute(
            select(usuarios).where(
                usuarios.c.id == usuario_id,
                usuarios.c.status == "ATIVO",
            )
        ).mappings().first()
        if usuario is None:
            raise ValueError("Usuário ativo não encontrado.")
        if usuario["perfil"] == perfil:
            return
        conexao.execute(
            update(usuarios)
            .where(usuarios.c.id == usuario_id)
            .values(perfil=perfil)
        )
        _registrar_evento(
            conexao,
            admin["email"],
            "Perfil alterado",
            f"Perfil de {usuario['email']} alterado de "
            f"{usuario['perfil']} para {perfil}.",
        )


def importar_equipamentos(dataframe: pd.DataFrame, usuario: str) -> dict[str, int]:
    if dataframe.empty:
        raise ValueError("A planilha não contém equipamentos para importar.")
    linhas: list[dict[str, Any]] = []
    codigos: set[str] = set()
    for _, linha in dataframe.iterrows():
        registro = {
            coluna: linha.get(rotulo)
            for coluna, rotulo in COLUNAS_EQUIPAMENTO.items()
        }
        codigo = normalizar_codigo(registro["numero_equipamento"])
        if not codigo:
            raise ValueError("Há equipamento sem número; corrija a planilha antes de importar.")
        if codigo in codigos:
            raise ValueError(
                f"A planilha contém mais de uma linha para o equipamento "
                f"{registro['numero_equipamento']}."
            )
        codigos.add(codigo)
        registro["codigo_normalizado"] = codigo
        for coluna in COLUNAS_EQUIPAMENTO:
            valor = registro[coluna]
            if valor is None or pd.isna(valor):
                registro[coluna] = ""
            elif isinstance(valor, (pd.Timestamp, datetime, date)):
                registro[coluna] = valor.isoformat()
            else:
                registro[coluna] = str(valor).strip()
        linhas.append(registro)

    novos = 0
    atualizados = 0
    inalterados = 0
    with _engine().begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE"})
        for registro in linhas:
            codigo = registro.pop("codigo_normalizado")
            existente = conexao.execute(
                select(equipamentos).where(
                    equipamentos.c.codigo_normalizado == codigo
                )
            ).mappings().first()
            if existente is None:
                valores = {
                    "codigo_normalizado": codigo,
                    **registro,
                    "atualizado_em": _agora(),
                }
                conexao.execute(insert(equipamentos).values(**valores))
                conexao.execute(
                    insert(auditoria_alteracoes).values(
                        data_hora=_agora(),
                        usuario=actor["email"],
                        equipamento=registro["numero_equipamento"],
                        codigo_normalizado=codigo,
                        campo="Cadastro",
                        valor_antigo="",
                        valor_novo="Equipamento cadastrado.",
                        tipo_alteracao="CADASTRO",
                    )
                )
                _registrar_evento(
                    conexao,
                    actor["email"],
                    "Equipamento cadastrado",
                    "Equipamento incluído pela importação da aba CAL.",
                    registro["numero_equipamento"],
                )
                novos += 1
                continue

            alteracoes = [
                (COLUNAS_EQUIPAMENTO[coluna], existente[coluna], valor)
                for coluna, valor in registro.items()
                if (existente[coluna] or "") != (valor or "")
            ]
            if not alteracoes:
                inalterados += 1
                continue
            conexao.execute(
                update(equipamentos)
                .where(equipamentos.c.codigo_normalizado == codigo)
                .values(**registro, atualizado_em=_agora())
            )
            for campo, antigo, novo in alteracoes:
                conexao.execute(
                    insert(auditoria_alteracoes).values(
                        data_hora=_agora(),
                        usuario=actor["email"],
                        equipamento=registro["numero_equipamento"],
                        codigo_normalizado=codigo,
                        campo=campo,
                        valor_antigo=antigo or "",
                        valor_novo=novo or "",
                        tipo_alteracao="IMPORTACAO",
                    )
                )
            _registrar_evento(
                conexao,
                actor["email"],
                "Equipamento atualizado",
                f"{len(alteracoes)} campo(s) atualizado(s) pela importação da aba CAL.",
                registro["numero_equipamento"],
            )
            atualizados += 1
        _registrar_evento(
            conexao,
            actor["email"],
            "Importação realizada",
            f"Importação concluída: {novos} novo(s), {atualizados} atualizado(s), "
            f"{inalterados} sem alteração.",
        )
    return {
        "novos": novos,
        "atualizados": atualizados,
        "inalterados": inalterados,
    }


def salvar_analise_certificado(
    nome_arquivo: str,
    conteudo_pdf: bytes,
    resultado: dict[str, Any],
    medicoes: pd.DataFrame,
    usuario: str = "Não autenticado",
) -> int:
    if not conteudo_pdf:
        raise ValueError("O certificado PDF está vazio.")
    hash_pdf = hashlib.sha256(conteudo_pdf).hexdigest()
    criterio = str(resultado.get("Critério aplicado", "") or "")
    chave_analise = hashlib.sha256(
        f"{hash_pdf}\0{criterio}".encode("utf-8")
    ).hexdigest()
    resultado_json = pd.DataFrame([resultado]).to_json(
        orient="records",
        force_ascii=False,
        date_format="iso",
    )
    medicoes_json = medicoes.to_json(
        orient="records",
        force_ascii=False,
        date_format="iso",
    )
    inicializar_banco()
    with _engine().begin() as conexao:
        existente = conexao.execute(
            select(certificados).where(
                certificados.c.chave_analise == chave_analise
            )
        ).mappings().first()
        valores = {
            "hash_pdf": hash_pdf,
            "nome_arquivo": str(nome_arquivo),
            "conteudo_pdf": conteudo_pdf,
            "resultado_json": resultado_json,
            "medicoes_json": medicoes_json,
            "usuario": _normalizar_email(usuario)
            if usuario != "Não autenticado"
            else usuario,
            "criado_em": _agora(),
        }
        if existente is None:
            chave = conexao.execute(
                insert(certificados).values(
                    chave_analise=chave_analise,
                    **valores,
                )
            )
            return int(chave.inserted_primary_key[0])
        if all(
            existente[campo] == valor
            for campo, valor in valores.items()
            if campo != "criado_em"
        ):
            return int(existente["id"])
        conexao.execute(
            update(certificados)
            .where(certificados.c.id == existente["id"])
            .values(**valores)
        )
        return int(existente["id"])


def listar_analises_certificados(limite: int = 100) -> pd.DataFrame:
    if limite < 1:
        raise ValueError("O limite de análises deve ser maior que zero.")
    inicializar_banco()
    with _engine().connect() as conexao:
        return pd.read_sql(
            select(
                certificados.c.id,
                certificados.c.nome_arquivo.label("Arquivo"),
                certificados.c.usuario.label("Analisado por"),
                certificados.c.criado_em.label("Data/Hora"),
                certificados.c.resultado_json,
            )
            .order_by(certificados.c.criado_em.desc())
            .limit(limite),
            conexao,
        )


def carregar_analise_certificado(analise_id: int) -> dict[str, Any]:
    inicializar_banco()
    with _engine().connect() as conexao:
        registro = conexao.execute(
            select(
                certificados.c.nome_arquivo,
                certificados.c.conteudo_pdf,
                certificados.c.resultado_json,
                certificados.c.medicoes_json,
            ).where(certificados.c.id == analise_id)
        ).mappings().first()
    if registro is None:
        raise ValueError("A análise de certificado selecionada não existe.")
    return {
        "nome_arquivo": registro["nome_arquivo"],
        "conteudo_pdf": bytes(registro["conteudo_pdf"]),
        "resultado": json.loads(registro["resultado_json"])[0],
        "medicoes": pd.DataFrame(json.loads(registro["medicoes_json"])),
    }


def carregar_equipamentos() -> pd.DataFrame:
    inicializar_banco()
    with _engine().connect() as conexao:
        registros = pd.read_sql(
            select(
                equipamentos.c.codigo_normalizado,
                *[equipamentos.c[coluna] for coluna in COLUNAS_EQUIPAMENTO],
            ).order_by(equipamentos.c.numero_equipamento),
            conexao,
        )
    registros = registros.rename(
        columns={
            **{coluna: rotulo for coluna, rotulo in COLUNAS_EQUIPAMENTO.items()},
            "codigo_normalizado": "Código Normalizado",
        }
    )
    if "Próxima Calibração" in registros:
        registros["Próxima Calibração"] = pd.to_datetime(
            registros["Próxima Calibração"],
            errors="coerce",
        )
        registros["Dias Restantes"] = (
            registros["Próxima Calibração"].dt.normalize()
            - pd.Timestamp(date.today())
        ).dt.days.astype("Int64")
        from core.calibration_management import classificar_vencimento

        registros["Status Calibração"] = registros["Próxima Calibração"].apply(
            classificar_vencimento
        )
    return registros


def resumir_equipamentos() -> tuple[int, int]:
    with _engine().connect() as conexao:
        registros = pd.read_sql(
            select(equipamentos.c.proxima_calibracao),
            conexao,
        )
    if registros.empty:
        return 0, 0

    from core.calibration_management import classificar_vencimento

    status = registros["proxima_calibracao"].apply(classificar_vencimento)
    return len(registros), int(status.eq("VENCIDO").sum())


def carregar_auditoria_alteracoes() -> pd.DataFrame:
    inicializar_banco()
    with _engine().connect() as conexao:
        return pd.read_sql(
            select(
                auditoria_alteracoes.c.data_hora.label("Data/Hora"),
                auditoria_alteracoes.c.usuario.label("Usuário"),
                auditoria_alteracoes.c.equipamento.label("Equipamento"),
                auditoria_alteracoes.c.campo.label("Campo Alterado"),
                auditoria_alteracoes.c.valor_antigo.label("Valor Antigo"),
                auditoria_alteracoes.c.valor_novo.label("Valor Novo"),
                auditoria_alteracoes.c.tipo_alteracao.label("Tipo de Alteração"),
            ).order_by(auditoria_alteracoes.c.data_hora.desc()),
            conexao,
        )


def carregar_historico_equipamento(codigo: str) -> pd.DataFrame:
    inicializar_banco()
    codigo = normalizar_codigo(codigo)
    with _engine().connect() as conexao:
        alteracoes = pd.read_sql(
            select(
                auditoria_alteracoes.c.data_hora.label("Data"),
                auditoria_alteracoes.c.usuario.label("Usuário"),
                auditoria_alteracoes.c.campo.label("Campo"),
                auditoria_alteracoes.c.valor_antigo.label("Valor Anterior"),
                auditoria_alteracoes.c.valor_novo.label("Valor Novo"),
            )
            .where(auditoria_alteracoes.c.codigo_normalizado == codigo)
            .order_by(auditoria_alteracoes.c.data_hora.desc()),
            conexao,
        )
        eventos = pd.read_sql(
            select(
                auditoria_eventos.c.data_hora.label("Data"),
                auditoria_eventos.c.usuario.label("Usuário"),
                auditoria_eventos.c.tipo_evento.label("Campo"),
                auditoria_eventos.c.descricao.label("Valor Novo"),
                auditoria_eventos.c.equipamento.label("Equipamento"),
            )
            .where(
                auditoria_eventos.c.codigo_normalizado == codigo,
                ~auditoria_eventos.c.tipo_evento.in_(
                    ("Equipamento cadastrado", "Equipamento atualizado")
                ),
            )
            .order_by(auditoria_eventos.c.data_hora.desc()),
            conexao,
        )
    if not eventos.empty:
        eventos["Valor Anterior"] = ""
        eventos = eventos[
            ["Data", "Usuário", "Campo", "Valor Anterior", "Valor Novo"]
        ]
    else:
        eventos = pd.DataFrame(
            columns=[
                "Data",
                "Usuário",
                "Campo",
                "Valor Anterior",
                "Valor Novo",
            ]
        )
    historico = pd.concat([alteracoes, eventos], ignore_index=True)
    return historico.sort_values("Data", ascending=False, kind="stable")


def carregar_auditoria_eventos() -> pd.DataFrame:
    inicializar_banco()
    with _engine().connect() as conexao:
        return pd.read_sql(
            select(
                auditoria_eventos.c.data_hora.label("Data/Hora"),
                auditoria_eventos.c.usuario.label("Usuário"),
                auditoria_eventos.c.tipo_evento.label("Tipo de Evento"),
                auditoria_eventos.c.descricao.label("Descrição"),
                auditoria_eventos.c.equipamento.label("Equipamento"),
            ).order_by(auditoria_eventos.c.data_hora.desc()),
            conexao,
        )


def exportar_dataframe_excel(dataframe: pd.DataFrame, nome_aba: str) -> bytes:
    from io import BytesIO

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        dataframe.to_excel(
            escritor,
            sheet_name=nome_aba[:31],
            index=False,
        )
        escritor.sheets[nome_aba[:31]].freeze_panes = "A2"
        escritor.sheets[nome_aba[:31]].auto_filter.ref = (
            escritor.sheets[nome_aba[:31]].dimensions
        )
    return buffer.getvalue()


def exportar_dataframe_csv(dataframe: pd.DataFrame) -> bytes:
    return dataframe.to_csv(index=False).encode("utf-8-sig")


__all__ = [
    "PAPEIS_AUDITORIA",
    "PERFIS_USUARIO",
    "STATUS_USUARIO",
    "aprovar_solicitacao",
    "autenticar_usuario",
    "alterar_perfil_usuario",
    "carregar_analise_certificado",
    "carregar_auditoria_alteracoes",
    "carregar_auditoria_eventos",
    "carregar_equipamentos",
    "carregar_historico_equipamento",
    "criar_usuario",
    "exportar_dataframe_csv",
    "exportar_dataframe_excel",
    "importar_equipamentos",
    "inicializar_banco",
    "listar_analises_certificados",
    "listar_usuarios",
    "normalizar_codigo",
    "provisionar_administrador",
    "redefinir_senha_administrador",
    "registrar_auditoria",
    "registrar_evento",
    "rejeitar_solicitacao",
    "salvar_analise_certificado",
    "usuario_ativo",
]
