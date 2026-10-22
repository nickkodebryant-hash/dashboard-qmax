import logging
from typing import Any

import pandas as pd
import streamlit as st
from sqlalchemy.exc import SQLAlchemyError

from core.authentication import render_autenticacao
from database.calibration_store import (
    PERFIS_USUARIO,
    alterar_perfil_usuario,
    aprovar_solicitacao,
    carregar_auditoria_alteracoes,
    carregar_auditoria_eventos,
    listar_usuarios,
    rejeitar_solicitacao,
)


logger = logging.getLogger(__name__)


def _gerar_tela_solicitacoes_acesso(usuario: dict[str, Any]) -> None:
    if usuario["perfil"] != "ADMIN":
        st.error("Somente o perfil ADMIN pode administrar acessos.")
        return
    st.subheader("Solicitações de acesso")
    try:
        pendentes = listar_usuarios("PENDENTE")
        if pendentes.empty:
            st.info("Não há solicitações pendentes.")
        else:
            st.dataframe(
                pendentes[["nome", "email", "criado_em"]],
                hide_index=True,
                use_container_width=True,
            )
            rotulos = {
                f"{linha.id} — {linha.nome} ({linha.email})": int(linha.id)
                for linha in pendentes.itertuples(index=False)
            }
            escolha = st.selectbox(
                "Solicitação",
                list(rotulos),
                key="solicitacao_cal_pendente",
            )
            perfil = st.selectbox(
                "Perfil a conceder",
                ["QUALIDADE", "LEITOR", "ADMIN"],
                key="perfil_solicitacao_cal",
            )
            aprovar, rejeitar = st.columns(2)
            if aprovar.button(
                "Aprovar solicitação",
                type="primary",
                key="aprovar_usuario_cal",
                use_container_width=True,
            ):
                aprovar_solicitacao(
                    rotulos[escolha],
                    usuario["email"],
                    perfil,
                )
                st.success("Usuário aprovado.")
                st.rerun()
            if rejeitar.button(
                "Rejeitar solicitação",
                key="rejeitar_usuario_cal",
                use_container_width=True,
            ):
                rejeitar_solicitacao(rotulos[escolha], usuario["email"])
                st.success("Solicitação rejeitada.")
                st.rerun()
    except (SQLAlchemyError, ValueError, PermissionError) as erro:
        logger.exception("Falha ao administrar solicitações de acesso.")
        st.error(f"Não foi possível concluir a operação de solicitações: {erro}")


def _gerar_tela_usuarios(usuario: dict[str, Any]) -> None:
    if usuario["perfil"] != "ADMIN":
        st.error("Somente o perfil ADMIN pode administrar acessos.")
        return
    st.subheader("Usuários")
    try:
        ativos = listar_usuarios("ATIVO")
        if ativos.empty:
            st.info("Não há usuários ativos.")
            return
        st.dataframe(
            ativos[["nome", "email", "perfil", "aprovado_em"]],
            hide_index=True,
            use_container_width=True,
        )
    except (SQLAlchemyError, ValueError, PermissionError) as erro:
        logger.exception("Falha ao consultar usuários ativos.")
        st.error(f"Não foi possível carregar os usuários: {erro}")


def _gerar_tela_perfis_permissoes(usuario: dict[str, Any]) -> None:
    if usuario["perfil"] != "ADMIN":
        st.error("Somente o perfil ADMIN pode administrar permissões.")
        return
    st.subheader("Perfis e permissões")
    try:
        ativos = listar_usuarios("ATIVO")
        if ativos.empty:
            st.info("Não há usuários ativos.")
            return
        alteraveis = ativos[ativos["email"] != usuario["email"]]
        if alteraveis.empty:
            st.caption("Não há outros usuários ativos para alterar.")
            return
        opcoes = {
            f"{linha.id} — {linha.nome} ({linha.email})": linha
            for linha in alteraveis.itertuples(index=False)
        }
        selecionado = st.selectbox(
            "Usuário",
            list(opcoes),
            key="usuario_cal_alterar_perfil",
        )
        novo_perfil = st.selectbox(
            "Perfil",
            list(PERFIS_USUARIO),
            index=list(PERFIS_USUARIO).index(opcoes[selecionado].perfil),
            key="novo_perfil_cal",
        )
        if st.button("Salvar perfil", key="salvar_perfil_cal"):
            alterar_perfil_usuario(
                int(opcoes[selecionado].id),
                usuario["email"],
                novo_perfil,
            )
            st.success("Perfil atualizado.")
            st.rerun()
    except (SQLAlchemyError, ValueError, PermissionError) as erro:
        logger.exception("Falha ao alterar perfil de usuário.")
        st.error(f"Não foi possível atualizar o perfil: {erro}")


def render_administracao_global() -> None:
    usuario = render_autenticacao()
    if usuario is None:
        return
    if usuario["perfil"] != "ADMIN":
        st.error("Somente o perfil ADMIN pode acessar a Administração.")
        return
    titulo_usuario, botao_sair = st.columns([4, 1])
    titulo_usuario.caption(f"Sessão: {usuario['nome']} · Perfil {usuario['perfil']}")
    if botao_sair.button(
        "Sair",
        key="sair_administracao_global",
        use_container_width=True,
    ):
        st.session_state.pop("_usuario_gestao_calibracao", None)
        st.rerun()

    st.markdown("## ⚙ Administração")
    usuarios, solicitacoes, permissoes, auditoria, notificacoes, logs, configuracoes, cadastros_cn = (
        st.tabs(
            [
                "Usuários",
                "Solicitações de Acesso",
                "Perfis e Permissões",
                "Auditoria Geral",
                "Notificações",
                "Logs",
                "Configurações",
                "Cadastros CN",
            ]
        )
    )
    with usuarios:
        _gerar_tela_usuarios(usuario)
    with solicitacoes:
        _gerar_tela_solicitacoes_acesso(usuario)
    with permissoes:
        _gerar_tela_perfis_permissoes(usuario)
    with auditoria:
        st.subheader("Auditoria Geral")
        try:
            calibracao = carregar_auditoria_alteracoes()
            eventos = carregar_auditoria_eventos()
            from database.cn_store import carregar_auditoria_administracao

            cns = carregar_auditoria_administracao(usuario["email"])
        except (SQLAlchemyError, PermissionError) as erro:
            logger.exception("Falha ao carregar auditoria geral.")
            st.error(f"Não foi possível carregar a auditoria geral: {erro}")
        else:
            auditoria_geral = pd.concat(
                [
                    calibracao.assign(Módulo="Calibração"),
                    eventos.assign(Módulo="Eventos"),
                    cns.assign(Módulo="CN"),
                ],
                ignore_index=True,
                sort=False,
            )
            st.dataframe(
                auditoria_geral,
                hide_index=True,
                use_container_width=True,
            )
    with notificacoes:
        st.subheader("Notificações")
        try:
            from database.cn_store import carregar_todas_notificacoes_administracao

            notificacoes_cns = carregar_todas_notificacoes_administracao(
                usuario["email"]
            )
        except (SQLAlchemyError, PermissionError) as erro:
            logger.exception("Falha ao carregar notificações globais.")
            st.error(f"Não foi possível carregar as notificações: {erro}")
        else:
            if notificacoes_cns.empty:
                st.info("Não há notificações administrativas registradas.")
            else:
                st.dataframe(
                    notificacoes_cns,
                    hide_index=True,
                    use_container_width=True,
                )
    with logs:
        st.subheader("Logs")
        st.info("Não há um repositório de logs de aplicação configurado.")
    with configuracoes:
        st.subheader("Configurações")
        st.info("Não há configurações globais persistidas para administrar.")
    with cadastros_cn:
        from core.cn_management import render_administracao_cadastros_mestres

        render_administracao_cadastros_mestres(usuario)
