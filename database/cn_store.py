from __future__ import annotations

from datetime import date, datetime, timezone
import json
import re
from typing import Any

import pandas as pd
from sqlalchemy import Column, DateTime, Integer, MetaData, String, Table, Text, delete, insert, select, update
from sqlalchemy import inspect
from sqlalchemy.engine import Engine
from sqlalchemy.exc import IntegrityError

from database.calibration_store import (
    _engine,
    _exigir_perfil,
    _validar_email,
    inicializar_banco as inicializar_banco_calibracao,
    usuarios,
)


ETAPAS_CN = (
    "MYLAR/AJUDA VISUAL (ATBO)",
    "MYLAR/AJUDA VISUAL (QUALY)",
    "MASTER BOARD (ATBO)",
    "MASTER BOARD (QUALY)",
    "PROG. (ATBO)",
    "PROG. (QUALY)",
    "DUPLICAÇÃO",
    "MYLAR DE INSP.",
    "DATA BP",
    "PPAP",
)
STATUS_ETAPA = ("PENDENTE", "EM ANDAMENTO", "CONCLUÍDO", "ATRASADO")
STATUS_CN = ("PROGRAMADO", "IMPLEMENTADO", "ATRASADO")
CRITICIDADES = ("Alta", "Média", "Baixa")

metadata = MetaData()

cn_referencias = Table(
    "cn_referencias",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("familia", Text, nullable=False, default=""),
    Column("descricao", Text, nullable=False, default=""),
    Column("codigo_linha", Text, nullable=False, default=""),
)

cn_registros = Table(
    "cn_registros",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("numero_cn", String(200), nullable=False, unique=True),
    Column("familia", Text, nullable=False),
    Column("descricao", Text, nullable=False),
    Column("codigo_linha", Text, nullable=False, default=""),
    Column("codigo_cn", Text, nullable=False),
    Column("resumo_modificacao", Text, nullable=False, default=""),
    Column("observacao", Text, nullable=False, default=""),
    Column("status", String(30), nullable=False),
    Column("responsavel_email", String(320), nullable=False),
    Column("criticidade", String(20), nullable=False, default="Média"),
    Column("data_criacao", DateTime(timezone=True), nullable=False),
    Column("data_atualizacao", DateTime(timezone=True), nullable=False),
    Column("usuario_criacao", String(320), nullable=False),
)

cn_etapas = Table(
    "cn_etapas",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("cn_id", Integer, nullable=False),
    Column("etapa", String(100), nullable=False),
    Column("data_planejada", String(30), nullable=False, default="NA"),
    Column("data_conclusao", String(30)),
    Column("responsavel_email", String(320), nullable=False),
    Column("data_recebimento", String(30)),
    Column("recebido_por", String(320)),
)

cn_responsaveis = Table(
    "cn_responsaveis",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("usuario", String(320), nullable=False, unique=True),
    Column("email", String(320), nullable=False, unique=True),
    Column("setor", String(200), nullable=False),
    Column("ativo", Integer, nullable=False, default=1),
)

cn_confirmacoes = Table(
    "cn_confirmacoes",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("cn_id", Integer, nullable=False, unique=True),
    Column("status", String(30), nullable=False),
    Column("solicitada_em", DateTime(timezone=True), nullable=False),
    Column("data_hora", DateTime(timezone=True)),
    Column("usuario", String(320)),
)

cn_auditoria = Table(
    "cn_auditoria",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("data_hora", DateTime(timezone=True), nullable=False),
    Column("usuario", String(320), nullable=False),
    Column("numero_cn", String(200), nullable=False),
    Column("campo", String(200), nullable=False),
    Column("valor_antigo", Text, nullable=False),
    Column("valor_novo", Text, nullable=False),
    Column("tipo_alteracao", String(50), nullable=False),
    Column("detalhe", Text, nullable=False, default=""),
)

cn_notificacoes = Table(
    "cn_notificacoes",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("cn_id", Integer, nullable=False),
    Column("destinatario", String(320), nullable=False),
    Column("tipo", String(100), nullable=False),
    Column("mensagem", Text, nullable=False),
    Column("data_hora", DateTime(timezone=True), nullable=False),
    Column("lida", Integer, nullable=False, default=0),
)

cn_backups = Table(
    "cn_backups",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("criado_em", DateTime(timezone=True), nullable=False),
    Column("criado_por", String(320), nullable=False),
    Column("quantidade_cns", Integer, nullable=False),
    Column("conteudo_json", Text, nullable=False),
)

cn_migracoes = Table(
    "cn_migracoes",
    metadata,
    Column("versao", String(100), primary_key=True),
    Column("aplicada_em", DateTime(timezone=True), nullable=False),
)

FAMILIA_REFERENCIA_GLOBAL = "Todas as famílias"

_VERSAO_REFERENCIAS_INICIAIS = "cn_referencias_mestres_2026_01"
_VERSAO_REFERENCIAS_COMPLEMENTARES = "cn_referencias_mestres_2026_02"
_REFERENCIAS_CNS_INICIAIS = (
    ("GM BSUV", "BODY BSUV", "GM GEM-20M"),
    ("GM BSUV", "CONSOLE BSUV", "GM GEM-80M"),
    ("GM BSUV", "FOG LAMP BSUV", "GM GEM-60M"),
    ("GM BSUV", "FRONT BUMPER BSUV", "GM GEM-90M"),
    ("GM BSUV", "FRONT DOOR CO DRIVER BSUV", "GM BSUV-PKP-70M"),
    ("GM BSUV", "FRONT DOOR DRIVER BSUV", "GM BSUV-PKP-74M"),
    ("GM BSUV", "IP EXTENSION BSUV", "GM GEM-42M"),
    ("GM BSUV", "LIFT GATE JPR BSUV", "GM GEM-43M"),
    ("GM BSUV", "LIFT GATE WRG BSUV", "GM GEM-41M"),
    ("GM BSUV", "REAR BUMPER BSUV", "GM GEM-91M"),
    ("GM BSUV", "REAR DOOR LEFT BSUV", "GM GEM-75M"),
    ("GM BSUV", "REAR DOOR RIGHT BSUV", "GM GEM-71M"),
    ("GM BSUV", "ROOF BSUV", "GM GEM-50M"),
    ("GM BSUV", "SUN ROOF BSUV", "GM GEM-50M"),
    ("GM PKP", "BODY PKP", "GM PKP-21M"),
    ("GM PKP", "BODY REAR JPR PKP", "GM PKP-110M"),
    ("GM PKP", "BODY REAR WRG PKP", "GM PKP-121M"),
    ("GM PKP", "CARGO LAMP PKP", "GM PKP-40M"),
    ("GM PKP", "CONSOLE PKP", "GM PKP-81M"),
    ("GM PKP", "END GATE PKP", "GM PKP-100M"),
    ("GM PKP", "FOG LAMP PKP", "GM PKP-60M"),
    ("GM PKP", "FRONT BUMPER PKP", "GM PKP-60/65M"),
    ("GM PKP", "FRONT DOOR CO DRIVER PKP", "GM BSUV-PKP-70M"),
    ("GM PKP", "FRONT DOOR DRIVER PKP", "GM BSUV-PKP-74M"),
    ("GM PKP", "IP EXTENSION PKP", "GM PKP-42M"),
    ("GM PKP", "REAR BUMPER PKP", "GM PKP-120M"),
    ("GM PKP", "REAR DOOR LEFT PKP", "GM PKP-73M"),
    ("GM PKP", "REAR DOOR RIGHT PKP", "GM PKP-72M"),
    ("GM PKP", "REAR VIEW CAMERA PKP", "GM PKP-40M"),
    ("GM PKP", "ROOF PKP", "GM PKP-51M"),
    ("GM BSUV/PKP", "CONSOLE BSUV/PKP", "GM BSUV/PKP-82M BC"),
    ("GM BSUV/PKP", "ENGINE LINHA 1 - MM1", "GM BSUV/PKP-34M"),
    ("GM BSUV/PKP", "ENGINE LINHA 2 - MM1", "GM BSUV/PKP-30/32M"),
    ("GM BSUV/PKP", "ENGINE LINHA 3 - MM3", "GM BSUV/PKP-31/33M"),
    ("GM BSUV/PKP", "IP BSUV/PKP", "GM GEM-12M BC"),
    ("VW", "AMAROK MIRGOR", "AMK-80M BC/AMK-81M BC"),
    ("VW", "YAPP", "VW YAPP-15M"),
    ("VW", "PRINCIPAL GP", "VW GP -13M"),
    ("VW", "PRINCIPAL KSK 246", "VW 246 - 20M BC"),
)
_REFERENCIAS_CNS_COMPLEMENTARES = (
    ("VW", "CABO MASSA", "NF-91M/E"),
    ("VW", "CABO BATERIA", "FX-92M/E"),
    ("GM BSUV/PKP", "ENGINE", "GM BSUV/PKP-36M BC"),
    ("GM PKP", "ENGINE MM3", "GM PKP-35M"),
    ("GM BSUV/PKP", "FRONT BUMPER PKP-BSUV", "GM BSUV-PKP-60E"),
    (
        FAMILIA_REFERENCIA_GLOBAL,
        "GERAL",
        "afeta todas as familias",
    ),
)


def _agora() -> datetime:
    return datetime.now(timezone.utc).replace(microsecond=0)


def _tem_valor(valor: Any) -> bool:
    if valor is None or pd.isna(valor):
        return False
    return str(valor).strip().upper() not in {"", "NA", "NAN", "NAT", "NONE"}


def _texto_referencia(valor: Any) -> str:
    return "" if valor is None or pd.isna(valor) else str(valor).strip()


def inicializar_banco(database_url: str | None = None) -> None:
    inicializar_banco_calibracao(database_url)
    engine = _engine(database_url)
    metadata.create_all(engine)
    colunas_etapas = {
        coluna["name"] for coluna in inspect(engine).get_columns("cn_etapas")
    }
    with engine.begin() as conexao:
        if "data_recebimento" not in colunas_etapas:
            conexao.exec_driver_sql(
                "ALTER TABLE cn_etapas ADD COLUMN data_recebimento VARCHAR(30)"
            )
        if "recebido_por" not in colunas_etapas:
            conexao.exec_driver_sql(
                "ALTER TABLE cn_etapas ADD COLUMN recebido_por VARCHAR(320)"
            )
    colunas_confirmacao = {
        coluna["name"]
        for coluna in inspect(engine).get_columns("cn_confirmacoes")
    }
    if "solicitada_em" not in colunas_confirmacao:
        with engine.begin() as conexao:
            conexao.exec_driver_sql(
                "ALTER TABLE cn_confirmacoes ADD COLUMN solicitada_em DATETIME"
            )
    with engine.begin() as conexao:
        for versao, referencias in (
            (_VERSAO_REFERENCIAS_INICIAIS, _REFERENCIAS_CNS_INICIAIS),
            (
                _VERSAO_REFERENCIAS_COMPLEMENTARES,
                _REFERENCIAS_CNS_COMPLEMENTARES,
            ),
        ):
            migracao_aplicada = conexao.execute(
                select(cn_migracoes.c.versao).where(
                    cn_migracoes.c.versao == versao
                )
            ).first()
            if migracao_aplicada is not None:
                continue

            existentes = {
                (linha["familia"], linha["descricao"]): linha["id"]
                for linha in conexao.execute(
                    select(
                        cn_referencias.c.id,
                        cn_referencias.c.familia,
                        cn_referencias.c.descricao,
                    )
                ).mappings()
            }
            for familia, descricao, codigo_linha in referencias:
                referencia_id = existentes.get((familia, descricao))
                if referencia_id is None:
                    resultado = conexao.execute(
                        insert(cn_referencias).values(
                            familia=familia,
                            descricao=descricao,
                            codigo_linha=codigo_linha,
                        )
                    )
                    existentes[(familia, descricao)] = int(
                        resultado.inserted_primary_key[0]
                    )
                else:
                    conexao.execute(
                        update(cn_referencias)
                        .where(cn_referencias.c.id == referencia_id)
                        .values(codigo_linha=codigo_linha)
                    )
            conexao.execute(
                insert(cn_migracoes).values(
                    versao=versao,
                    aplicada_em=_agora(),
                )
            )


def _normalizar_linhas_referencias(
    referencias: dict[str, Any],
) -> set[tuple[str, str, str]]:
    return {
        (
            _texto_referencia(linha.get("familia")),
            _texto_referencia(linha.get("descricao")),
            _texto_referencia(linha.get("codigo_linha")),
        )
        for linha in referencias.get("linhas", [])
        if any(
            _texto_referencia(linha.get(campo))
            for campo in ("familia", "descricao", "codigo_linha")
        )
    }


def _salvar_linhas_referencias_em_conexao(
    conexao: Any,
    linhas: set[tuple[str, str, str]],
) -> None:
    existentes = {
        (
            linha["familia"],
            linha["descricao"],
            linha["codigo_linha"],
        )
        for linha in conexao.execute(
            select(
                cn_referencias.c.familia,
                cn_referencias.c.descricao,
                cn_referencias.c.codigo_linha,
            )
        ).mappings()
    }
    if linhas == existentes:
        return
    conexao.execute(delete(cn_referencias))
    if linhas:
        conexao.execute(
            insert(cn_referencias),
            [
                {
                    "familia": familia,
                    "descricao": descricao,
                    "codigo_linha": codigo_linha,
                }
                for familia, descricao, codigo_linha in sorted(linhas)
            ],
        )


def salvar_referencias_cns(
    referencias: dict[str, Any],
    usuario: str,
    database_url: str | None = None,
) -> None:
    linhas = _normalizar_linhas_referencias(referencias)
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE"})
        _salvar_linhas_referencias_em_conexao(conexao, linhas)


def carregar_referencias_cns(
    database_url: str | None = None,
) -> dict[str, Any]:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        linhas = [
            dict(linha)
            for linha in conexao.execute(
                select(
                    cn_referencias.c.familia,
                    cn_referencias.c.descricao,
                    cn_referencias.c.codigo_linha,
                ).order_by(
                    cn_referencias.c.familia,
                    cn_referencias.c.descricao,
                )
            ).mappings()
        ]
    return {
        "familias": sorted({linha["familia"] for linha in linhas if linha["familia"]}),
        "descricoes": sorted(
            {linha["descricao"] for linha in linhas if linha["descricao"]}
        ),
        "codigo_linha": {
            linha["descricao"].casefold(): linha["codigo_linha"]
            for linha in linhas
            if linha["descricao"]
        },
        "linhas": linhas,
    }


def _auditar(
    conexao: Any,
    usuario: str,
    numero_cn: str,
    campo: str,
    antigo: Any,
    novo: Any,
    tipo: str = "EDIÇÃO",
    detalhe: str = "",
) -> None:
    conexao.execute(
        insert(cn_auditoria).values(
            data_hora=_agora(),
            usuario=usuario,
            numero_cn=numero_cn,
            campo=campo,
            valor_antigo="" if antigo is None else str(antigo),
            valor_novo="" if novo is None else str(novo),
            tipo_alteracao=tipo,
            detalhe=detalhe,
        )
    )


def _notificar(
    conexao: Any,
    cn_id: int,
    destinatario: str,
    tipo: str,
    mensagem: str,
) -> None:
    conexao.execute(
        insert(cn_notificacoes).values(
            cn_id=cn_id,
            destinatario=destinatario,
            tipo=tipo,
            mensagem=mensagem,
            data_hora=_agora(),
            lida=0,
        )
    )


def _criar_backup_cns_em_conexao(
    conexao: Any,
    usuario: str,
) -> tuple[int, str, int]:
    tabelas = {
        "cn_registros": cn_registros,
        "cn_etapas": cn_etapas,
        "cn_confirmacoes": cn_confirmacoes,
        "cn_notificacoes": cn_notificacoes,
    }
    conteudo = {
        nome: [
            dict(linha)
            for linha in conexao.execute(select(tabela)).mappings()
        ]
        for nome, tabela in tabelas.items()
    }
    conteudo_json = json.dumps(conteudo, ensure_ascii=False, default=str)
    quantidade_cns = len(conteudo["cn_registros"])
    resultado = conexao.execute(
        insert(cn_backups).values(
            criado_em=_agora(),
            criado_por=usuario,
            quantidade_cns=quantidade_cns,
            conteudo_json=conteudo_json,
        )
    )
    return int(resultado.inserted_primary_key[0]), conteudo_json, quantidade_cns


def criar_backup_cns(
    usuario: str,
    database_url: str | None = None,
) -> dict[str, Any]:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN"})
        backup_id, conteudo_json, quantidade_cns = _criar_backup_cns_em_conexao(
            conexao,
            actor["email"],
        )
    return {
        "id": backup_id,
        "conteudo": conteudo_json.encode("utf-8"),
        "quantidade_cns": quantidade_cns,
    }


def listar_backups_cns(
    usuario: str,
    database_url: str | None = None,
) -> pd.DataFrame:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        _exigir_perfil(conexao, usuario, {"ADMIN"})
        return pd.read_sql(
            select(
                cn_backups.c.id,
                cn_backups.c.criado_em.label("Criado em"),
                cn_backups.c.criado_por.label("Criado por"),
                cn_backups.c.quantidade_cns.label("Quantidade de CNs"),
            ).order_by(cn_backups.c.id.desc()),
            conexao,
        )


def baixar_backup_cns(
    backup_id: int,
    usuario: str,
    database_url: str | None = None,
) -> bytes:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        _exigir_perfil(conexao, usuario, {"ADMIN"})
        conteudo = conexao.execute(
            select(cn_backups.c.conteudo_json).where(
                cn_backups.c.id == backup_id
            )
        ).scalar_one_or_none()
    if conteudo is None:
        raise ValueError("O backup de CN selecionado não existe.")
    return conteudo.encode("utf-8")


def _validar_codigo(codigo: Any, rotulo: str) -> str:
    texto = "" if codigo is None else str(codigo).strip()
    if not texto:
        raise ValueError(f"{rotulo} é obrigatório.")
    return texto


def _data_iso(valor: Any, campo: str, permitir_na: bool = False) -> str:
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return "NA" if permitir_na else ""
    if permitir_na and str(valor).strip().upper() == "NA":
        return "NA"
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    texto = str(valor).strip()
    try:
        return date.fromisoformat(texto).isoformat()
    except ValueError as erro:
        raise ValueError(f"{campo}: use uma data válida ou NA.") from erro


def criar_cn(
    dados: dict[str, Any],
    etapas: dict[str, dict[str, Any]],
    usuario: str,
    database_url: str | None = None,
) -> int:
    inicializar_banco(database_url)
    numero = _validar_codigo(dados.get("numero_cn"), "Número CN")
    codigo = _validar_codigo(dados.get("codigo_cn"), "Código CN")
    responsavel = _validar_codigo(dados.get("responsavel_email"), "Responsável")
    familia = _validar_codigo(dados.get("familia"), "Família")
    descricao = _validar_codigo(dados.get("descricao"), "Descrição")
    criticidade = str(dados.get("criticidade", "Média"))
    if criticidade not in CRITICIDADES:
        raise ValueError("Criticidade inválida.")
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE"})
        responsavel_ativo = conexao.execute(
            select(cn_responsaveis.c.email).where(
                cn_responsaveis.c.email == responsavel.casefold(),
                cn_responsaveis.c.ativo == 1,
            )
        ).first()
        if responsavel_ativo is None:
            raise ValueError("Selecione um responsável ativo cadastrado.")
        try:
            resultado = conexao.execute(
                insert(cn_registros).values(
                    numero_cn=numero,
                    familia=familia,
                    descricao=descricao,
                    codigo_linha=str(dados.get("codigo_linha") or ""),
                    codigo_cn=codigo,
                    resumo_modificacao=str(dados.get("resumo_modificacao") or ""),
                    observacao=str(dados.get("observacao") or ""),
                    status="PROGRAMADO",
                    responsavel_email=responsavel,
                    criticidade=criticidade,
                    data_criacao=_agora(),
                    data_atualizacao=_agora(),
                    usuario_criacao=actor["email"],
                )
            )
        except IntegrityError as erro:
            raise ValueError("Já existe uma CN com esse número.") from erro
        cn_id = int(resultado.inserted_primary_key[0])
        for nome_etapa in ETAPAS_CN:
            valores = etapas.get(nome_etapa, {})
            responsavel_etapa = str(
                valores.get("responsavel_email") or responsavel
            ).casefold()
            if conexao.execute(
                select(cn_responsaveis.c.email).where(
                    cn_responsaveis.c.email == responsavel_etapa,
                    cn_responsaveis.c.ativo == 1,
                )
            ).first() is None:
                raise ValueError(f"Responsável inválido para a etapa {nome_etapa}.")
            conexao.execute(
                insert(cn_etapas).values(
                    cn_id=cn_id,
                    etapa=nome_etapa,
                    data_planejada=_data_iso(
                        valores.get("data_planejada", "NA"),
                        f"Prazo de {nome_etapa}",
                        permitir_na=True,
                    ),
                    data_conclusao=_data_iso(
                        valores.get("data_conclusao"),
                        f"Conclusão de {nome_etapa}",
                    )
                    or None,
                    responsavel_email=responsavel_etapa,
                )
            )
        conexao.execute(
            insert(cn_confirmacoes).values(
                cn_id=cn_id,
                status="PENDENTE",
                solicitada_em=_agora(),
            )
        )
        _auditar(conexao, actor["email"], numero, "Cadastro", "", codigo, "CRIAÇÃO")
        _notificar(
            conexao,
            cn_id,
            responsavel,
            "CN Criada",
            (
                f"A CN {numero} foi atribuída a você. Confirme o recebimento "
                "das etapas sob sua responsabilidade."
            ),
        )
    return cn_id


def _importar_cns_em_conexao(
    conexao: Any,
    registros: list[dict[str, Any]],
    actor: Any,
) -> dict[str, int]:
    resultado = {"novas": 0, "duplicadas": 0}
    for registro in registros:
        numero = _validar_codigo(registro.get("numero_cn"), "Número CN")
        codigo = _validar_codigo(registro.get("codigo_cn"), "Código CN")
        existente = conexao.execute(
            select(cn_registros.c.id).where(
                cn_registros.c.numero_cn == numero
            )
        ).first()
        if existente:
            resultado["duplicadas"] += 1
            continue
        responsavel = _validar_codigo(
            registro.get("responsavel_email"), "Responsável"
        )
        responsavel_ativo = conexao.execute(
            select(cn_responsaveis.c.email).where(
                cn_responsaveis.c.email == responsavel.casefold(),
                cn_responsaveis.c.ativo == 1,
            )
        ).first()
        if responsavel_ativo is None:
            raise ValueError("Selecione um responsável ativo cadastrado.")
        valores = {
            "numero_cn": numero,
            "familia": _validar_codigo(registro.get("familia"), "Família"),
            "descricao": _validar_codigo(registro.get("descricao"), "Descrição"),
            "codigo_linha": str(registro.get("codigo_linha") or ""),
            "codigo_cn": codigo,
            "resumo_modificacao": str(registro.get("resumo_modificacao") or ""),
            "observacao": str(registro.get("observacao") or ""),
            "status": "PROGRAMADO",
            "responsavel_email": responsavel,
            "criticidade": "Média",
            "data_criacao": _agora(),
            "data_atualizacao": _agora(),
            "usuario_criacao": actor["email"],
        }
        try:
            novo = conexao.execute(insert(cn_registros).values(**valores))
        except IntegrityError:
            resultado["duplicadas"] += 1
            continue
        cn_id = int(novo.inserted_primary_key[0])
        for nome_etapa in ETAPAS_CN:
            conexao.execute(
                insert(cn_etapas).values(
                    cn_id=cn_id,
                    etapa=nome_etapa,
                    data_planejada=_data_iso(
                        registro.get("etapas", {}).get(nome_etapa, "NA"),
                        f"Prazo de {nome_etapa}",
                        permitir_na=True,
                    ),
                    data_conclusao=None,
                    responsavel_email=responsavel,
                )
            )
        conexao.execute(
            insert(cn_confirmacoes).values(
                cn_id=cn_id,
                status="PENDENTE",
                solicitada_em=_agora(),
            )
        )
        _auditar(
            conexao,
            actor["email"],
            numero,
            "Importação da planilha",
            "",
            codigo,
            "CRIAÇÃO",
        )
        _notificar(
            conexao,
            cn_id,
            responsavel,
            "CN Criada",
            f"A CN {numero} foi importada e atribuída a você. Confirme o recebimento.",
        )
        resultado["novas"] += 1
    return resultado


def importar_cns(
    registros: list[dict[str, Any]],
    usuario: str,
    database_url: str | None = None,
) -> dict[str, int]:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE"})
        return _importar_cns_em_conexao(conexao, registros, actor)


def _limpar_todas_cns_em_conexao(
    conexao: Any,
    usuario: str,
    backup_id: int,
) -> int:
    registros = conexao.execute(
        select(cn_registros).order_by(cn_registros.c.id)
    ).mappings().all()
    for registro in registros:
        _auditar(
            conexao,
            usuario,
            registro["numero_cn"],
            "Registro CN",
            json.dumps(dict(registro), ensure_ascii=False, default=str),
            "",
            "LIMPEZA TOTAL",
            f"CN removida após criação do backup #{backup_id}.",
        )
    conexao.execute(delete(cn_notificacoes))
    conexao.execute(delete(cn_confirmacoes))
    conexao.execute(delete(cn_etapas))
    conexao.execute(delete(cn_registros))
    return len(registros)


def limpar_dados_cns(
    usuario: str,
    database_url: str | None = None,
) -> dict[str, int]:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN"})
        backup_id, _, quantidade_cns = _criar_backup_cns_em_conexao(
            conexao,
            actor["email"],
        )
        removidas = _limpar_todas_cns_em_conexao(
            conexao,
            actor["email"],
            backup_id,
        )
    return {
        "backup_id": backup_id,
        "cns_removidas": removidas,
        "cns_no_backup": quantidade_cns,
    }


def reimportar_base_cns(
    registros: list[dict[str, Any]],
    usuario: str,
    database_url: str | None = None,
    referencias: dict[str, Any] | None = None,
) -> dict[str, int]:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN"})
        backup_id, _, quantidade_cns = _criar_backup_cns_em_conexao(
            conexao,
            actor["email"],
        )
        removidas = _limpar_todas_cns_em_conexao(
            conexao,
            actor["email"],
            backup_id,
        )
        if referencias is not None:
            _salvar_linhas_referencias_em_conexao(
                conexao,
                _normalizar_linhas_referencias(referencias),
            )
        importacao = _importar_cns_em_conexao(conexao, registros, actor)
    return {
        "backup_id": backup_id,
        "cns_no_backup": quantidade_cns,
        "cns_removidas": removidas,
        **importacao,
    }


def atualizar_cn(
    cn_id: int,
    dados: dict[str, Any],
    usuario: str,
    database_url: str | None = None,
) -> None:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE"})
        atual = conexao.execute(
            select(cn_registros).where(cn_registros.c.id == cn_id)
        ).mappings().first()
        if atual is None:
            raise ValueError("A CN selecionada não existe.")
        colunas = (
            "familia",
            "descricao",
            "codigo_linha",
            "codigo_cn",
            "resumo_modificacao",
            "observacao",
            "responsavel_email",
            "criticidade",
        )
        valores = {coluna: dados[coluna] for coluna in colunas}
        if valores["criticidade"] not in CRITICIDADES:
            raise ValueError("Criticidade inválida.")
        if not valores["responsavel_email"]:
            raise ValueError("Cada CN deve ter um responsável.")
        if conexao.execute(
            select(cn_responsaveis.c.email).where(
                cn_responsaveis.c.email == valores["responsavel_email"].casefold(),
                cn_responsaveis.c.ativo == 1,
            )
        ).first() is None:
            raise ValueError("Selecione um responsável ativo cadastrado.")
        for campo, novo in valores.items():
            antigo = atual[campo]
            if str(antigo or "") != str(novo or ""):
                tipo = "ALTERAÇÃO DE RESPONSÁVEL" if campo == "responsavel_email" else "EDIÇÃO"
                _auditar(conexao, actor["email"], atual["numero_cn"], campo, antigo, novo, tipo)
        conexao.execute(
            update(cn_registros)
            .where(cn_registros.c.id == cn_id)
            .values(**valores, data_atualizacao=_agora())
        )
        if atual["responsavel_email"] != valores["responsavel_email"]:
            confirmacao = conexao.execute(
                select(cn_confirmacoes).where(cn_confirmacoes.c.cn_id == cn_id)
            ).mappings().first()
            if confirmacao is not None:
                conexao.execute(
                    update(cn_confirmacoes)
                    .where(cn_confirmacoes.c.cn_id == cn_id)
                    .values(
                        status="PENDENTE",
                        solicitada_em=_agora(),
                        data_hora=None,
                        usuario=None,
                    )
                )
            _notificar(
                conexao,
                cn_id,
                valores["responsavel_email"],
                "CN Criada",
                f"A CN {atual['numero_cn']} foi atribuída a você. Confirme o recebimento.",
            )


def excluir_cn(
    cn_id: int,
    usuario: str,
    database_url: str | None = None,
) -> None:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN"})
        registro = conexao.execute(
            select(cn_registros).where(cn_registros.c.id == cn_id)
        ).mappings().first()
        if registro is None:
            raise ValueError("A CN selecionada não existe.")

        backup_id, _, _ = _criar_backup_cns_em_conexao(
            conexao,
            actor["email"],
        )
        _excluir_registro_cn(
            conexao,
            actor["email"],
            cn_id,
            registro,
            "EXCLUSÃO",
            f"CN excluída pela Administração CN's após backup #{backup_id}.",
        )


def limpar_cns_importadas(
    usuario: str,
    database_url: str | None = None,
) -> int:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN"})
        registros = conexao.execute(
            select(cn_registros).where(
                cn_registros.c.numero_cn.in_(_consulta_numeros_cns_importadas())
            )
        ).mappings().all()
        backup_id = None
        if registros:
            backup_id, _, _ = _criar_backup_cns_em_conexao(
                conexao,
                actor["email"],
            )
        for registro in registros:
            _excluir_registro_cn(
                conexao,
                actor["email"],
                int(registro["id"]),
                registro,
                "LIMPEZA DE IMPORTAÇÃO",
                (
                    "CN importada removida pela opção Limpar importações da Administração CN's. "
                    f"Backup #{backup_id} gerado previamente."
                ),
            )
    return len(registros)


def listar_numeros_cns_importadas(
    usuario: str,
    database_url: str | None = None,
) -> list[str]:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        _exigir_perfil(conexao, usuario, {"ADMIN"})
        return list(
            conexao.execute(
                _consulta_numeros_cns_importadas().order_by(cn_auditoria.c.numero_cn)
            ).scalars()
        )


def _consulta_numeros_cns_importadas() -> Any:
    auditoria_posterior = cn_auditoria.alias("auditoria_posterior")
    return (
        select(cn_auditoria.c.numero_cn)
        .where(
            cn_auditoria.c.campo == "Importação da planilha",
            cn_auditoria.c.tipo_alteracao == "CRIAÇÃO",
            ~select(auditoria_posterior.c.id)
            .where(
                auditoria_posterior.c.numero_cn == cn_auditoria.c.numero_cn,
                auditoria_posterior.c.tipo_alteracao == "CRIAÇÃO",
                auditoria_posterior.c.id > cn_auditoria.c.id,
            )
            .exists(),
        )
        .distinct()
    )


def _excluir_registro_cn(
    conexao: Any,
    usuario: str,
    cn_id: int,
    registro: Any,
    tipo_auditoria: str,
    detalhe: str,
) -> None:
    confirmacao = conexao.execute(
        select(cn_confirmacoes).where(cn_confirmacoes.c.cn_id == cn_id)
    ).mappings().first()
    snapshot = {
        "registro": dict(registro),
        "etapas": [
            dict(etapa)
            for etapa in conexao.execute(
                select(cn_etapas).where(cn_etapas.c.cn_id == cn_id)
            ).mappings()
        ],
        "confirmacao": dict(confirmacao) if confirmacao is not None else None,
        "notificacoes": [
            dict(notificacao)
            for notificacao in conexao.execute(
                select(cn_notificacoes).where(cn_notificacoes.c.cn_id == cn_id)
            ).mappings()
        ],
    }
    _auditar(
        conexao,
        usuario,
        registro["numero_cn"],
        "Registro CN",
        json.dumps(snapshot, ensure_ascii=False, default=str),
        "",
        tipo_auditoria,
        detalhe,
    )
    conexao.execute(delete(cn_notificacoes).where(cn_notificacoes.c.cn_id == cn_id))
    conexao.execute(delete(cn_confirmacoes).where(cn_confirmacoes.c.cn_id == cn_id))
    conexao.execute(delete(cn_etapas).where(cn_etapas.c.cn_id == cn_id))
    conexao.execute(delete(cn_registros).where(cn_registros.c.id == cn_id))


def atualizar_etapa(
    cn_id: int,
    nome_etapa: str,
    data_planejada: Any,
    data_conclusao: Any,
    responsavel_email: str,
    usuario: str,
    database_url: str | None = None,
) -> None:
    if nome_etapa not in ETAPAS_CN:
        raise ValueError("Etapa inválida.")
    planejada = _data_iso(data_planejada, "Data planejada", permitir_na=True)
    conclusao = _data_iso(data_conclusao, "Data real de conclusão") or None
    if not responsavel_email:
        raise ValueError("A etapa deve ter um responsável.")
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE", "LEITOR"})
        if conexao.execute(
            select(cn_responsaveis.c.email).where(
                cn_responsaveis.c.email == responsavel_email.casefold(),
                cn_responsaveis.c.ativo == 1,
            )
        ).first() is None:
            raise ValueError("Selecione um responsável ativo cadastrado.")
        registro = conexao.execute(
            select(cn_registros.c.numero_cn).where(cn_registros.c.id == cn_id)
        ).first()
        if registro is None:
            raise ValueError("A CN selecionada não existe.")
        atual = conexao.execute(
            select(cn_etapas).where(
                cn_etapas.c.cn_id == cn_id, cn_etapas.c.etapa == nome_etapa
            )
        ).mappings().first()
        if atual is None:
            raise ValueError("A etapa não está cadastrada para esta CN.")
        if (
            actor["perfil"] == "LEITOR"
            and atual["responsavel_email"].casefold() != actor["email"].casefold()
        ):
            raise PermissionError(
                "Somente o responsável designado pode atualizar esta etapa."
            )
        if (
            actor["perfil"] == "LEITOR"
            and responsavel_email.casefold() != actor["email"].casefold()
        ):
            raise PermissionError("O responsável da etapa só pode ser alterado por ADMIN/QUALIDADE.")
        for campo, novo in (
            ("data_planejada", planejada),
            ("data_conclusao", conclusao),
            ("responsavel_email", responsavel_email),
        ):
            antigo = atual[campo]
            if str(antigo or "") != str(novo or ""):
                _auditar(
                    conexao,
                    actor["email"],
                    registro.numero_cn,
                    f"{nome_etapa} — {campo}",
                    antigo,
                    novo,
                    "ALTERAÇÃO DE DATAS"
                    if campo in ("data_planejada", "data_conclusao")
                    else "ALTERAÇÃO DE RESPONSÁVEL",
                )
        conexao.execute(
            update(cn_etapas)
            .where(cn_etapas.c.id == atual["id"])
            .values(
                data_planejada=planejada,
                data_conclusao=conclusao,
                responsavel_email=responsavel_email,
            )
        )
        conexao.execute(
            update(cn_registros)
            .where(cn_registros.c.id == cn_id)
            .values(data_atualizacao=_agora())
        )


def confirmar_recebimento(
    cn_id: int,
    usuario: str,
    database_url: str | None = None,
) -> None:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE", "LEITOR"})
        cn = conexao.execute(
            select(cn_registros).where(cn_registros.c.id == cn_id)
        ).mappings().first()
        if cn is None:
            raise ValueError("A CN selecionada não existe.")
        if cn["responsavel_email"].casefold() != actor["email"].casefold():
            raise PermissionError("Somente o responsável designado pode confirmar.")
        confirmacao = conexao.execute(
            select(cn_confirmacoes).where(cn_confirmacoes.c.cn_id == cn_id)
        ).mappings().first()
        if confirmacao is None or confirmacao["status"] == "CONFIRMADO":
            return
        conexao.execute(
            update(cn_confirmacoes)
            .where(cn_confirmacoes.c.cn_id == cn_id)
            .values(status="CONFIRMADO", data_hora=_agora(), usuario=actor["email"])
        )
        _auditar(
            conexao,
            actor["email"],
            cn["numero_cn"],
            "Recebimento",
            "Pendente de Confirmação",
            "Confirmado",
            "CONFIRMAÇÃO DE RECEBIMENTO",
        )


def confirmar_recebimento_etapa(
    cn_id: int,
    nome_etapa: str,
    usuario: str,
    database_url: str | None = None,
) -> None:
    if nome_etapa not in ETAPAS_CN:
        raise ValueError("Etapa inválida.")
    inicializar_banco(database_url)
    agora = _agora()
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE", "LEITOR"})
        cn = conexao.execute(
            select(cn_registros).where(cn_registros.c.id == cn_id)
        ).mappings().first()
        if cn is None:
            raise ValueError("A CN selecionada não existe.")
        etapa = conexao.execute(
            select(cn_etapas).where(
                cn_etapas.c.cn_id == cn_id,
                cn_etapas.c.etapa == nome_etapa,
            )
        ).mappings().first()
        if etapa is None:
            raise ValueError("A etapa não está cadastrada para esta CN.")
        if (
            actor["perfil"] == "LEITOR"
            and etapa["responsavel_email"].casefold() != actor["email"].casefold()
        ):
            raise PermissionError(
                "Somente o responsável designado pode confirmar esta etapa."
            )
        if _tem_valor(etapa["data_recebimento"]):
            return

        data_recebimento = agora.date().isoformat()
        conexao.execute(
            update(cn_etapas)
            .where(cn_etapas.c.id == etapa["id"])
            .values(
                data_recebimento=data_recebimento,
                recebido_por=actor["email"],
                data_conclusao=(
                    etapa["data_conclusao"]
                    if _tem_valor(etapa["data_conclusao"])
                    else data_recebimento
                ),
            )
        )
        _auditar(
            conexao,
            actor["email"],
            cn["numero_cn"],
            f"{nome_etapa} — Recebimento",
            "Pendente",
            f"Confirmado em {data_recebimento}",
            "CONFIRMAÇÃO DE RECEBIMENTO",
        )
        if not _tem_valor(etapa["data_conclusao"]):
            _auditar(
                conexao,
                actor["email"],
                cn["numero_cn"],
                f"{nome_etapa} — Data de conclusão",
                "",
                data_recebimento,
                "ALTERAÇÃO DE DATAS",
                "Conclusão registrada pela confirmação de recebimento da etapa.",
            )
        conexao.execute(
            update(cn_registros)
            .where(cn_registros.c.id == cn_id)
            .values(data_atualizacao=agora)
        )

        etapas = conexao.execute(
            select(cn_etapas.c.data_recebimento).where(cn_etapas.c.cn_id == cn_id)
        ).all()
        if len(etapas) == len(ETAPAS_CN) and all(
            _tem_valor(item.data_recebimento) for item in etapas
        ):
            confirmacao = conexao.execute(
                select(cn_confirmacoes).where(cn_confirmacoes.c.cn_id == cn_id)
            ).mappings().first()
            if confirmacao is None:
                conexao.execute(
                    insert(cn_confirmacoes).values(
                        cn_id=cn_id,
                        status="CONFIRMADO",
                        solicitada_em=agora,
                        data_hora=agora,
                        usuario=actor["email"],
                    )
                )
            elif confirmacao["status"] != "CONFIRMADO":
                conexao.execute(
                    update(cn_confirmacoes)
                    .where(cn_confirmacoes.c.cn_id == cn_id)
                    .values(
                        status="CONFIRMADO",
                        data_hora=agora,
                        usuario=actor["email"],
                    )
                )
                _auditar(
                    conexao,
                    actor["email"],
                    cn["numero_cn"],
                    "Recebimento de todas as etapas",
                    "Pendente",
                    "Confirmado",
                    "CONFIRMAÇÃO DE RECEBIMENTO",
                )


def _status_etapa(
    data_planejada: Any,
    data_conclusao: Any,
    hoje: date,
    data_recebimento: Any = None,
) -> str:
    if _tem_valor(data_conclusao) or _tem_valor(data_recebimento):
        return "CONCLUÍDO"
    if not data_planejada or str(data_planejada).upper() == "NA":
        return "PENDENTE"
    try:
        prazo = date.fromisoformat(str(data_planejada))
    except ValueError:
        return "PENDENTE"
    return "ATRASADO" if prazo < hoje else "EM ANDAMENTO"


def _status_cn(etapas: list[dict[str, Any]], hoje: date) -> str:
    status = [
        _status_etapa(
            etapa.get("data_planejada"),
            etapa.get("data_conclusao"),
            hoje,
            etapa.get("data_recebimento"),
        )
        for etapa in etapas
    ]
    if "ATRASADO" in status:
        return "ATRASADO"
    if etapas and all(etapa_status == "CONCLUÍDO" for etapa_status in status):
        return "IMPLEMENTADO"
    return "PROGRAMADO"


def resumir_cns(database_url: str | None = None) -> tuple[int, int]:
    engine = _engine(database_url)
    with engine.connect() as conexao:
        registros = pd.read_sql(select(cn_registros.c.id), conexao)
        etapas = pd.read_sql(
            select(
                cn_etapas.c.cn_id,
                cn_etapas.c.data_planejada,
                cn_etapas.c.data_conclusao,
                cn_etapas.c.data_recebimento,
            ),
            conexao,
        )
    if registros.empty:
        return 0, 0

    etapas_por_cn: dict[int, list[dict[str, Any]]] = {}
    for etapa in etapas.to_dict("records"):
        etapas_por_cn.setdefault(int(etapa["cn_id"]), []).append(etapa)

    hoje = date.today()
    atrasadas = sum(
        _status_cn(etapas_por_cn.get(int(registro_id), []), hoje) == "ATRASADO"
        for registro_id in registros["id"]
    )
    return len(registros), int(atrasadas)


def _carregar_cns_com_status_persistido(
    database_url: str | None = None,
) -> tuple[pd.DataFrame, dict[int, str]]:
    engine = _engine(database_url)
    with engine.connect() as conexao:
        registros = pd.read_sql(
            select(
                cn_registros,
                cn_confirmacoes.c.status.label("status_confirmacao"),
                cn_confirmacoes.c.data_hora.label("confirmado_em"),
                cn_confirmacoes.c.usuario.label("confirmado_por"),
                cn_responsaveis.c.setor.label("setor_responsavel"),
            )
            .select_from(
                cn_registros.outerjoin(
                    cn_confirmacoes, cn_registros.c.id == cn_confirmacoes.c.cn_id
                ).outerjoin(
                    cn_responsaveis,
                    cn_registros.c.responsavel_email == cn_responsaveis.c.email,
                )
            )
            .order_by(cn_registros.c.data_criacao.desc()),
            conexao,
        )
        etapas = pd.read_sql(select(cn_etapas), conexao)
    if registros.empty:
        return registros, {}
    status_persistidos = {
        int(registro["id"]): registro["status"]
        for registro in registros.to_dict("records")
    }
    mapa_etapas: dict[int, list[dict[str, Any]]] = {}
    if not etapas.empty:
        for item in etapas.to_dict("records"):
            mapa_etapas.setdefault(int(item["cn_id"]), []).append(item)
    hoje = date.today()
    for idx, registro in registros.iterrows():
        etapas_cn = mapa_etapas.get(int(registro["id"]), [])
        status_cn = _status_cn(etapas_cn, hoje)
        registros.at[idx, "status"] = status_cn
        for etapa in etapas_cn:
            etapa["status_calculado"] = _status_etapa(
                etapa.get("data_planejada"),
                etapa.get("data_conclusao"),
                hoje,
                etapa.get("data_recebimento"),
            )
        mapa_etapas[int(registro["id"])] = etapas_cn
    registros["etapas"] = registros["id"].map(mapa_etapas)
    return registros, status_persistidos


def carregar_cns(database_url: str | None = None) -> pd.DataFrame:
    registros, _ = _carregar_cns_com_status_persistido(database_url)
    return registros


def sincronizar_status_cns(database_url: str | None = None) -> int:
    inicializar_banco(database_url)
    registros, status_persistidos = _carregar_cns_com_status_persistido(database_url)
    alteracoes = [
        (
            int(registro["id"]),
            registro["numero_cn"],
            status_persistidos[int(registro["id"])],
            registro["status"],
        )
        for registro in registros.to_dict("records")
        if status_persistidos[int(registro["id"])] != registro["status"]
    ]
    if not alteracoes:
        return 0

    with _engine(database_url).begin() as conexao:
        for cn_id, numero, antigo, novo in alteracoes:
            _auditar(
                conexao,
                "sistema",
                numero,
                "Status",
                antigo,
                novo,
                "ALTERAÇÃO DE STATUS",
                "Status calculado automaticamente pelas etapas.",
            )
            conexao.execute(
                update(cn_registros)
                .where(cn_registros.c.id == cn_id)
                .values(status=novo)
            )
    return len(alteracoes)


def carregar_responsaveis(database_url: str | None = None) -> pd.DataFrame:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        return pd.read_sql(
            select(cn_responsaveis)
            .where(cn_responsaveis.c.ativo == 1)
            .order_by(cn_responsaveis.c.usuario),
            conexao,
        )


def listar_responsaveis_administracao(
    usuario: str,
    database_url: str | None = None,
) -> pd.DataFrame:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        _exigir_perfil(conexao, usuario, {"ADMIN"})
        return pd.read_sql(
            select(cn_responsaveis).order_by(cn_responsaveis.c.usuario),
            conexao,
        )


def administrar_responsavel(
    responsavel_id: int | None,
    nome: str,
    email: str,
    setor: str,
    ativo: bool,
    usuario: str,
    database_url: str | None = None,
) -> None:
    nome = str(nome or "").strip()
    email = _validar_email(str(email or ""))
    setor = str(setor or "").strip()
    if not nome or not setor:
        raise ValueError("Usuário, e-mail e setor são obrigatórios.")

    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        actor = _exigir_perfil(conexao, usuario, {"ADMIN"})
        if ativo and conexao.execute(
            select(usuarios.c.id).where(
                usuarios.c.email == email,
                usuarios.c.status == "ATIVO",
            )
        ).first() is None:
            raise ValueError("Somente um usuário ativo pode ser responsável ativo.")

        consulta = select(cn_responsaveis).where(
            cn_responsaveis.c.id == responsavel_id
        ) if responsavel_id is not None else select(cn_responsaveis).where(
            cn_responsaveis.c.email == email
        )
        atual = conexao.execute(consulta).mappings().first()
        duplicado = conexao.execute(
            select(cn_responsaveis.c.id).where(
                cn_responsaveis.c.email == email,
                cn_responsaveis.c.id != (atual["id"] if atual else -1),
            )
        ).first()
        if duplicado is not None:
            raise ValueError("Esse e-mail já está cadastrado como responsável.")

        valores = {
            "usuario": nome,
            "email": email,
            "setor": setor,
            "ativo": int(ativo),
        }
        if atual is None:
            conexao.execute(insert(cn_responsaveis).values(**valores))
            valores_antigos = {}
        else:
            valores_antigos = dict(atual)
            conexao.execute(
                update(cn_responsaveis)
                .where(cn_responsaveis.c.id == atual["id"])
                .values(**valores)
            )

        for campo, novo in valores.items():
            antigo = valores_antigos.get(campo, "")
            if str(antigo or "") != str(novo or ""):
                _auditar(
                    conexao,
                    actor["email"],
                    "RESPONSÁVEIS",
                    f"{campo} · {email}",
                    antigo,
                    novo,
                    "GESTÃO DE RESPONSÁVEL",
                )


def salvar_responsavel(
    nome: str,
    email: str,
    setor: str,
    criado_por: str,
    database_url: str | None = None,
) -> None:
    if not nome.strip() or not email.strip() or not setor.strip():
        raise ValueError("Usuário, e-mail e setor são obrigatórios.")
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        _exigir_perfil(conexao, criado_por, {"ADMIN", "QUALIDADE"})
        atual = conexao.execute(
            select(cn_responsaveis).where(cn_responsaveis.c.email == email.casefold())
        ).mappings().first()
        if atual is None:
            conexao.execute(
                insert(cn_responsaveis).values(
                    usuario=nome.strip(),
                    email=email.strip().casefold(),
                    setor=setor.strip(),
                    ativo=1,
                )
            )
        else:
            conexao.execute(
                update(cn_responsaveis)
                .where(cn_responsaveis.c.id == atual["id"])
                .values(usuario=nome.strip(), setor=setor.strip(), ativo=1)
            )


def carregar_auditoria(database_url: str | None = None) -> pd.DataFrame:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        return pd.read_sql(
            select(
                cn_auditoria.c.data_hora.label("Data/Hora"),
                cn_auditoria.c.usuario.label("Usuário"),
                cn_auditoria.c.numero_cn.label("CN"),
                cn_auditoria.c.campo.label("Campo"),
                cn_auditoria.c.valor_antigo.label("Valor Antigo"),
                cn_auditoria.c.valor_novo.label("Valor Novo"),
                cn_auditoria.c.tipo_alteracao.label("Tipo Alteração"),
                cn_auditoria.c.detalhe.label("Detalhe"),
            ).order_by(cn_auditoria.c.data_hora.desc()),
            conexao,
        )


def carregar_auditoria_administracao(
    usuario: str,
    database_url: str | None = None,
) -> pd.DataFrame:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        _exigir_perfil(conexao, usuario, {"ADMIN"})
        return pd.read_sql(
            select(
                cn_auditoria.c.data_hora.label("Data/Hora"),
                cn_auditoria.c.usuario.label("Usuário"),
                cn_auditoria.c.numero_cn.label("CN"),
                cn_auditoria.c.campo.label("Campo"),
                cn_auditoria.c.valor_antigo.label("Valor Antigo"),
                cn_auditoria.c.valor_novo.label("Valor Novo"),
                cn_auditoria.c.tipo_alteracao.label("Tipo Alteração"),
                cn_auditoria.c.detalhe.label("Detalhe"),
            ).order_by(cn_auditoria.c.data_hora.desc()),
            conexao,
        )


def carregar_notificacoes(
    usuario: str,
    database_url: str | None = None,
) -> pd.DataFrame:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        return pd.read_sql(
            select(
                cn_notificacoes.c.id,
                cn_notificacoes.c.tipo.label("Tipo"),
                cn_notificacoes.c.mensagem.label("Notificação"),
                cn_notificacoes.c.data_hora.label("Data/Hora"),
                cn_notificacoes.c.lida.label("Lida"),
                cn_registros.c.numero_cn.label("CN"),
            )
            .select_from(
                cn_notificacoes.join(
                    cn_registros, cn_notificacoes.c.cn_id == cn_registros.c.id
                )
            )
            .where(cn_notificacoes.c.destinatario == usuario.casefold())
            .order_by(cn_notificacoes.c.data_hora.desc()),
            conexao,
        )


def carregar_todas_notificacoes_administracao(
    usuario: str,
    database_url: str | None = None,
) -> pd.DataFrame:
    inicializar_banco(database_url)
    with _engine(database_url).connect() as conexao:
        _exigir_perfil(conexao, usuario, {"ADMIN"})
        return pd.read_sql(
            select(
                cn_notificacoes.c.id,
                cn_notificacoes.c.destinatario.label("Destinatário"),
                cn_notificacoes.c.tipo.label("Tipo"),
                cn_notificacoes.c.mensagem.label("Notificação"),
                cn_notificacoes.c.data_hora.label("Data/Hora"),
                cn_notificacoes.c.lida.label("Lida"),
                cn_registros.c.numero_cn.label("CN"),
            )
            .select_from(
                cn_notificacoes.join(
                    cn_registros, cn_notificacoes.c.cn_id == cn_registros.c.id
                )
            )
            .order_by(cn_notificacoes.c.data_hora.desc()),
            conexao,
        )


def marcar_notificacao_lida(
    notificacao_id: int,
    usuario: str,
    database_url: str | None = None,
) -> None:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        conexao.execute(
            update(cn_notificacoes)
            .where(
                cn_notificacoes.c.id == notificacao_id,
                cn_notificacoes.c.destinatario == usuario.casefold(),
            )
            .values(lida=1)
        )


def processar_escalonamentos(database_url: str | None = None) -> int:
    inicializar_banco(database_url)
    engine: Engine = _engine(database_url)
    agora = _agora()
    criadas = 0
    with engine.begin() as conexao:
        pendentes = conexao.execute(
            select(
                cn_registros.c.id,
                cn_registros.c.numero_cn,
                cn_registros.c.responsavel_email,
                cn_confirmacoes.c.data_hora,
                cn_confirmacoes.c.solicitada_em,
                cn_confirmacoes.c.status,
                cn_registros.c.data_criacao,
            )
            .join(cn_confirmacoes, cn_registros.c.id == cn_confirmacoes.c.cn_id)
            .where(cn_confirmacoes.c.status == "PENDENTE")
        ).mappings()
        for item in pendentes:
            inicio = item["solicitada_em"] or item["data_hora"] or item["data_criacao"]
            if inicio.tzinfo is None:
                inicio = inicio.replace(tzinfo=timezone.utc)
            horas = (agora - inicio).total_seconds() / 3600
            for limite, tipo in (
                (48, "CN sem confirmação — lembrete"),
                (72, "CN sem confirmação — alerta"),
            ):
                if horas < limite:
                    continue
                existe = conexao.execute(
                    select(cn_notificacoes.c.id).where(
                        cn_notificacoes.c.cn_id == item["id"],
                        cn_notificacoes.c.tipo == tipo,
                    )
                ).first()
                if existe is None:
                    _notificar(
                        conexao,
                        int(item["id"]),
                        item["responsavel_email"],
                        tipo,
                        f"A CN {item['numero_cn']} aguarda confirmação de recebimento há {int(horas)} horas.",
                    )
                    criadas += 1
        etapa_rows = conexao.execute(
            select(
                cn_registros.c.id,
                cn_registros.c.numero_cn,
                cn_registros.c.responsavel_email,
                cn_etapas.c.etapa,
                cn_etapas.c.data_planejada,
                cn_etapas.c.data_conclusao,
                cn_etapas.c.data_recebimento,
            ).join(cn_etapas, cn_registros.c.id == cn_etapas.c.cn_id)
        ).mappings()
        hoje = date.today()
        for etapa in etapa_rows:
            mensagens: list[tuple[str, str]] = []
            if _status_etapa(
                etapa["data_planejada"],
                etapa["data_conclusao"],
                hoje,
                etapa["data_recebimento"],
            ) == "ATRASADO":
                mensagens.append(
                    (
                        "Etapa vencida",
                        f"A etapa {etapa['etapa']} da CN {etapa['numero_cn']} está atrasada.",
                    )
                )
            if (
                etapa["etapa"] == "PPAP"
                and not etapa["data_conclusao"]
                and not etapa["data_recebimento"]
            ):
                mensagens.append(
                    ("PPAP pendente", f"O PPAP da CN {etapa['numero_cn']} está pendente.")
                )
            if (
                etapa["etapa"] == "DATA BP"
                and not etapa["data_conclusao"]
                and not etapa["data_recebimento"]
            ):
                mensagens.append(
                    (
                        "Data BP pendente",
                        f"A data de BP da CN {etapa['numero_cn']} está pendente.",
                    )
                )
            for tipo, mensagem in mensagens:
                chave_tipo = f"{tipo}: {etapa['etapa']}" if tipo == "Etapa vencida" else tipo
                existe = conexao.execute(
                    select(cn_notificacoes.c.id).where(
                        cn_notificacoes.c.cn_id == etapa["id"],
                        cn_notificacoes.c.tipo == chave_tipo,
                    )
                ).first()
                if existe is None:
                    _notificar(
                        conexao,
                        int(etapa["id"]),
                        etapa["responsavel_email"],
                        chave_tipo,
                        mensagem,
                    )
                    criadas += 1
    return criadas


def processar_notificacoes_cns(database_url: str | None = None) -> int:
    inicializar_banco(database_url)
    registros = carregar_cns(database_url)
    notificacoes_atraso = 0
    with _engine(database_url).begin() as conexao:
        atrasadas = registros[registros["status"] == "ATRASADO"]
        for registro in atrasadas.to_dict("records"):
            existente = conexao.execute(
                select(cn_notificacoes.c.id).where(
                    cn_notificacoes.c.cn_id == int(registro["id"]),
                    cn_notificacoes.c.tipo == "CN Atrasada",
                )
            ).first()
            if existente is None:
                _notificar(
                    conexao,
                    int(registro["id"]),
                    registro["responsavel_email"],
                    "CN Atrasada",
                    f"A CN {registro['numero_cn']} possui etapa(s) vencida(s).",
                )
                notificacoes_atraso += 1
    return notificacoes_atraso + processar_escalonamentos(database_url)


def registrar_alerta_importacao(
    usuario: str,
    numero_cn: str,
    coluna: str,
    linha: int,
    detalhe: str,
    database_url: str | None = None,
) -> None:
    inicializar_banco(database_url)
    with _engine(database_url).begin() as conexao:
        _exigir_perfil(conexao, usuario, {"ADMIN", "QUALIDADE"})
        _auditar(
            conexao,
            usuario,
            numero_cn,
            coluna,
            "",
            detalhe,
            "ALERTA DE IMPORTAÇÃO",
            f"Linha {linha} — {detalhe}",
        )


def data_para_iso(valor: Any, coluna: str, linha: int) -> tuple[str, str | None]:
    if valor is None or (isinstance(valor, str) and valor.strip().upper() in {"", "-", "_", "?", "NULL", "NA"}):
        return "NA", None
    if isinstance(valor, datetime):
        return valor.date().isoformat(), None
    if isinstance(valor, date):
        return valor.isoformat(), None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return "NA", f"Linha {linha}, coluna {coluna}: valor numérico não reconhecido como data."
    texto = str(valor).strip()
    candidatos = re.findall(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{4}\b", texto)
    if len(candidatos) > 1:
        return "NA", f"Linha {linha}, coluna {coluna}: mais de uma data encontrada: {texto}"
    if len(candidatos) == 1:
        try:
            return datetime.strptime(candidatos[0].replace("-", "/"), "%d/%m/%Y").date().isoformat(), None
        except ValueError:
            return "NA", f"Linha {linha}, coluna {coluna}: data inválida: {candidatos[0]}"
    try:
        return datetime.strptime(texto, "%d/%m/%Y").date().isoformat(), None
    except ValueError:
        return "NA", f"Linha {linha}, coluna {coluna}: conteúdo não é uma data válida: {texto}"
