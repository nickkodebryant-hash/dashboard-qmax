import logging
from typing import Any

import streamlit as st
from sqlalchemy.exc import SQLAlchemyError

from database.calibration_store import (
    autenticar_usuario,
    criar_usuario,
    listar_usuarios,
    usuario_ativo,
)


logger = logging.getLogger(__name__)


def render_autenticacao() -> dict[str, Any] | None:
    usuario_sessao = st.session_state.get("_usuario_gestao_calibracao")
    if isinstance(usuario_sessao, dict):
        try:
            usuario_atual = usuario_ativo(usuario_sessao["email"])
        except SQLAlchemyError as erro:
            logger.exception("Falha ao validar sessão da gestão de calibração.")
            st.error(f"Não foi possível validar sua sessão: {erro}")
            return None
        if usuario_atual is not None:
            st.session_state["_usuario_gestao_calibracao"] = usuario_atual
            return usuario_atual
        st.session_state.pop("_usuario_gestao_calibracao", None)

    st.subheader("Acesso à Gestão de Calibração")
    entrar, solicitar = st.tabs(["Entrar", "Solicitar acesso"])
    with entrar:
        with st.form("login_gestao_calibracao"):
            email = st.text_input("E-mail", key="login_cal_email")
            senha = st.text_input("Senha", type="password", key="login_cal_senha")
            entrar_enviado = st.form_submit_button("Entrar", type="primary")
        if entrar_enviado:
            try:
                usuario = autenticar_usuario(email, senha)
            except (ValueError, SQLAlchemyError) as erro:
                logger.exception("Falha na autenticação da gestão de calibração.")
                st.error(f"Não foi possível autenticar: {erro}")
            else:
                if usuario is None:
                    st.error("Credenciais inválidas ou solicitação ainda não aprovada.")
                else:
                    st.session_state["_usuario_gestao_calibracao"] = usuario
                    st.rerun()
    with solicitar:
        with st.form("solicitar_acesso_calibracao"):
            nome = st.text_input("Nome completo", key="solicitar_cal_nome")
            email_novo = st.text_input("E-mail corporativo", key="solicitar_cal_email")
            senha_nova = st.text_input(
                "Senha (mínimo 12 caracteres)",
                type="password",
                key="solicitar_cal_senha",
            )
            senha_confirmacao = st.text_input(
                "Confirme a senha",
                type="password",
                key="solicitar_cal_confirmar_senha",
            )
            solicitar_enviado = st.form_submit_button("Enviar solicitação")
        if solicitar_enviado:
            if senha_nova != senha_confirmacao:
                st.error("As senhas não coincidem.")
            else:
                try:
                    criar_usuario(nome, email_novo, senha_nova)
                except (ValueError, SQLAlchemyError) as erro:
                    logger.exception("Falha ao criar solicitação de acesso.")
                    st.error(f"Não foi possível solicitar acesso: {erro}")
                else:
                    st.success(
                        "Solicitação registrada. O acesso ficará disponível após aprovação."
                    )
    try:
        usuarios_cadastrados = listar_usuarios()
    except SQLAlchemyError as erro:
        logger.exception("Falha ao consultar a configuração de usuários.")
        st.error(f"Não foi possível consultar o banco de usuários: {erro}")
    else:
        if usuarios_cadastrados.empty:
            st.info(
                "Ainda não existe administrador. No ambiente local, crie o primeiro "
                "administrador com `python -m database.provision_calibration_admin`."
            )
    return None
