import calendar
import io
import math
import re
import unicodedata
from datetime import date

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from config.settings import MONTH_ORDER


def normalizar_texto_dashboard(valor):
    texto = "" if pd.isna(valor) else str(valor)
    return "".join(
        caractere
        for caractere in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(caractere)
    ).strip().upper()


def mes_padrao_dashboard(meses_disponiveis, hoje=None):
    if not meses_disponiveis:
        raise ValueError("Não há meses válidos para exibir no painel de ações.")
    hoje = hoje or date.today()
    return next(
        (
            mes
            for mes in meses_disponiveis
            if MONTH_ORDER.get(normalizar_texto_dashboard(mes)) == hoje.month
        ),
        meses_disponiveis[0],
    )


def calcular_situacao_prazo(dataframe, coluna_status, coluna_data_prazo, hoje=None):
    hoje = hoje or date.today()
    status = dataframe[coluna_status].map(normalizar_texto_dashboard)
    valores_prazo = dataframe[coluna_data_prazo]
    datas_iso = valores_prazo.astype(str).str.match(r"^\d{4}-\d{2}-\d{2}")
    prazos = pd.Series(pd.NaT, index=dataframe.index, dtype="datetime64[ns]")
    prazos.loc[datas_iso] = pd.to_datetime(
        valores_prazo.loc[datas_iso],
        errors="coerce",
        dayfirst=False,
    )
    prazos.loc[~datas_iso] = pd.to_datetime(
        valores_prazo.loc[~datas_iso],
        errors="coerce",
        dayfirst=True,
    )
    situacoes = pd.Series("Sem prazo", index=dataframe.index, dtype="object")
    situacoes.loc[status == "CONCLUIDO"] = "Concluída"
    abertas = status == "ABERTO"
    situacoes.loc[abertas & prazos.notna() & (prazos.dt.date < hoje)] = "Vencida"
    situacoes.loc[abertas & prazos.notna() & (prazos.dt.date >= hoje)] = "No prazo"
    return situacoes


def calcular_meta_esperada(meta_mensal, mes, hoje=None):
    hoje = hoje or date.today()
    mes_normalizado = normalizar_texto_dashboard(mes)
    numero_mes = MONTH_ORDER.get(mes_normalizado)
    if numero_mes is None:
        raise ValueError(f"Mês não reconhecido para cálculo da meta: {mes!r}.")
    if numero_mes < hoje.month:
        return float(meta_mensal)
    if numero_mes > hoje.month:
        return 0.0
    dias_no_mes = calendar.monthrange(hoje.year, numero_mes)[1]
    return float(meta_mensal) * hoje.day / dias_no_mes


def calcular_indicadores_acoes(
    dataframe,
    coluna_status,
    situacoes_prazo,
    meta_mensal,
    mes,
    hoje=None,
):
    hoje = hoje or date.today()
    status = dataframe[coluna_status].map(normalizar_texto_dashboard)
    abertas = status == "ABERTO"
    total_realizado = len(dataframe)
    meta_esperada = calcular_meta_esperada(meta_mensal, mes, hoje)
    return {
        "meta_total": float(meta_mensal),
        "total_realizado": total_realizado,
        "falta_para_meta": max(0, float(meta_mensal) - total_realizado),
        "percentual_atingimento": (
            total_realizado / float(meta_mensal) if meta_mensal else 0.0
        ),
        "concluidas": int((status == "CONCLUIDO").sum()),
        "em_aberto": int(abertas.sum()),
        "vencidas": int((abertas & situacoes_prazo.eq("Vencida")).sum()),
        "no_prazo": int((abertas & situacoes_prazo.eq("No prazo")).sum()),
        "meta_esperada": meta_esperada,
        "diferenca_meta": total_realizado - meta_esperada,
    }


def _numero_target(valor):
    if pd.isna(valor) or isinstance(valor, bool):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if not texto or texto.startswith("="):
        return 0.0
    texto = texto.replace(" ", "")
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def _normalizar_cabecalho_target(valor):
    return re.sub(
        r"\s+",
        " ",
        re.sub(
            r"[^A-Z0-9]+",
            " ",
            normalizar_texto_dashboard(valor),
        ),
    ).strip()


def calcular_distribuicao_target(linhas, meta_mensal, denominador=None):
    dataframe = linhas.copy()
    if dataframe.empty:
        dataframe["Target Mensal Calculado"] = pd.Series(dtype=float)
        dataframe["Total Realizado Calculado"] = pd.Series(dtype=float)
        dataframe["Pendente Calculado"] = pd.Series(dtype=float)
        dataframe["Atingimento"] = pd.Series(dtype=float)
        return dataframe

    if denominador is None:
        denominador = float(
            (dataframe["Peso"] * dataframe["Qtd. M.O"]).sum()
        )
    fatores = dataframe["Peso"] * dataframe["Qtd. M.O"]
    if denominador > 0:
        dataframe["Target Mensal Calculado"] = (
            float(meta_mensal) * fatores / denominador
        )
    else:
        dataframe["Target Mensal Calculado"] = 0.0
    dataframe["Total Realizado Calculado"] = dataframe[
        [coluna for coluna in dataframe.columns if coluna.startswith("Realizado WK")]
    ].sum(axis=1)
    dataframe["Pendente Calculado"] = (
        dataframe["Target Mensal Calculado"]
        - dataframe["Total Realizado Calculado"]
    ).clip(lower=0).map(lambda valor: math.floor(valor + 0.5))
    dataframe["Atingimento"] = dataframe.apply(
        lambda linha: (
            linha["Total Realizado Calculado"] / linha["Target Mensal Calculado"]
            if linha["Target Mensal Calculado"] > 0
            else 0.0
        ),
        axis=1,
    )
    return dataframe


def ler_target_gerencia_excel(arquivo, hoje=None):
    if arquivo is None:
        raise ValueError("Selecione o arquivo Excel com a aba Target Gerência.")

    workbook = openpyxl.load_workbook(
        io.BytesIO(arquivo.getvalue()),
        read_only=True,
        data_only=False,
    )
    try:
        worksheet = None
        linhas = []
        cabecalhos_requeridos = {"GERENTE", "PESO", "QTD M O", "TARGET MENSAL"}
        for sheet in workbook.worksheets:
            if not normalizar_texto_dashboard(sheet.title).startswith("TARGET GER"):
                continue
            linhas_candidatas = list(sheet.iter_rows(values_only=True))
            if any(
                cabecalhos_requeridos.issubset(
                    {
                        _normalizar_cabecalho_target(valor)
                        for valor in linha
                    }
                )
                for linha in linhas_candidatas
            ):
                worksheet = sheet
                linhas = linhas_candidatas
                break
        if worksheet is None:
            raise ValueError("A aba Target Gerência não foi encontrada no arquivo.")

        regioes = []
        for indice, linha in enumerate(linhas):
            cabecalhos = [_normalizar_cabecalho_target(valor) for valor in linha]
            if not {"GERENTE", "PESO", "QTD M O", "TARGET MENSAL"}.issubset(
                set(cabecalhos)
            ):
                continue
            indices = {
                "responsavel": cabecalhos.index("GERENTE"),
                "turno": cabecalhos.index("TURNO") if "TURNO" in cabecalhos else None,
                "peso": cabecalhos.index("PESO"),
                "pessoas": cabecalhos.index("QTD M O"),
            }
            indices["target_mensal"] = cabecalhos.index("TARGET MENSAL")
            semanas = []
            for coluna, cabecalho in enumerate(cabecalhos):
                correspondencia = re.fullmatch(r"TARGET WEEK\s*(.+)", cabecalho)
                if correspondencia:
                    semana = correspondencia.group(1).strip()
                    coluna_realizado = next(
                        (
                            indice_coluna
                            for indice_coluna, titulo in enumerate(cabecalhos)
                            if (
                                titulo == "REALIZADO"
                                or re.fullmatch(r"REALIZADO\s*\d+", titulo)
                            )
                            and indice_coluna > coluna
                        ),
                        None,
                    )
                    if coluna_realizado is not None:
                        semanas.append((semana, coluna, coluna_realizado))

            registros = []
            for linha_dados in linhas[indice + 1 :]:
                if len(linha_dados) <= max(
                    indices["responsavel"],
                    indices["peso"],
                    indices["pessoas"],
                ):
                    break
                valores_normalizados = {
                    normalizar_texto_dashboard(valor) for valor in linha_dados
                }
                if "TOTAL PROGRAMADO MES" in valores_normalizados:
                    break
                nome = linha_dados[indices["responsavel"]]
                if not nome:
                    continue
                if normalizar_texto_dashboard(nome) in {"GERENTE", "REG"}:
                    break
                peso = _numero_target(linha_dados[indices["peso"]])
                pessoas = _numero_target(linha_dados[indices["pessoas"]])
                if not peso or not pessoas:
                    continue
                registro = {
                    "Responsável": str(nome).strip(),
                    "Turno": (
                        str(linha_dados[indices["turno"]]).strip()
                        if indices["turno"] is not None
                        and linha_dados[indices["turno"]] is not None
                        else "(Não informado)"
                    ),
                    "Peso": peso,
                    "Qtd. M.O": pessoas,
                }
                for semana, indice_target, indice_realizado in semanas:
                    registro[f"Target WK {semana}"] = _numero_target(
                        linha_dados[indice_target]
                        if indice_target < len(linha_dados)
                        else None
                    )
                    registro[f"Realizado WK {semana}"] = _numero_target(
                        linha_dados[indice_realizado]
                        if indice_realizado < len(linha_dados)
                        else None
                    )
                registros.append(registro)
            if registros:
                nome_secao = "Gerentes" if not regioes else "Líderes"
                regioes.append(
                    {
                        "nome": nome_secao,
                        "registros": pd.DataFrame(registros),
                    }
                )

        if not regioes:
            raise ValueError(
                "Não foi possível localizar os blocos de gerente e líder na aba Target Gerência."
            )

        hoje = hoje or date.today()
        meta_mensal = None
        meta_ano = hoje.year
        for linha in linhas:
            if len(linha) <= 25:
                continue
            ano = _numero_target(linha[24])
            meta = _numero_target(linha[25])
            if int(ano) == hoje.year and meta > 0:
                meta_mensal = meta
                meta_ano = int(ano)
                break
        if meta_mensal is None:
            candidatos = [
                (int(_numero_target(linha[24])), _numero_target(linha[25]))
                for linha in linhas
                if len(linha) > 25
                and _numero_target(linha[24]) >= 2000
                and _numero_target(linha[25]) > 0
            ]
            if candidatos:
                meta_ano, meta_mensal = max(candidatos)
        if meta_mensal is None:
            meta_mensal = 560.0

        denominador = float(
            (
                regioes[0]["registros"]["Peso"]
                * regioes[0]["registros"]["Qtd. M.O"]
            ).sum()
        )
        for regiao in regioes:
            regiao["registros"] = calcular_distribuicao_target(
                regiao["registros"],
                meta_mensal,
                denominador,
            )
        return {
            "ano": meta_ano,
            "meta_mensal": meta_mensal,
            "grupos": regioes,
        }
    finally:
        workbook.close()


def exibir_painel_target_gerencia(target):
    grupos = target["grupos"]
    nome_grupo = st.radio(
        "Visão do time",
        [grupo["nome"] for grupo in grupos],
        horizontal=True,
        key="target_gerencia_grupo",
    )
    grupo = next(item for item in grupos if item["nome"] == nome_grupo)
    dados = grupo["registros"].copy()
    turnos = sorted(dados["Turno"].dropna().unique().tolist())
    if len(turnos) > 1:
        turno = st.selectbox("Turno", ["Todos", *turnos], key="target_gerencia_turno")
        if turno != "Todos":
            dados = dados[dados["Turno"] == turno]

    meta = float(dados["Target Mensal Calculado"].sum())
    realizado = float(dados["Total Realizado Calculado"].sum())
    pendente = max(0.0, meta - realizado)
    atingimento = realizado / meta if meta > 0 else 0.0
    if atingimento < 0.8:
        cor, situacao = "#dc2626", "DESEMPENHO CRÍTICO"
    elif atingimento < 1:
        cor, situacao = "#d97706", "EM PROGRESSO"
    else:
        cor, situacao = "#15803d", "META ATINGIDA"

    st.caption(
        f"Distribuição ponderada por peso × quantidade de mão de obra · Meta de referência {target['ano']}"
    )
    metricas = [
        ("Meta mensal", f"{meta:,.0f}", "#1d4ed8"),
        ("Realizado", f"{realizado:,.0f}", "#0f766e"),
        ("Faltam para a meta", f"{pendente:,.0f}", "#b45309"),
        ("% de atingimento", f"{atingimento:.1%}", cor),
    ]
    st.markdown(
        """
        <style>
        .target-kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin:12px 0; }
        .target-kpi { background:linear-gradient(145deg,#fff,#f8fafc); border:1px solid #e2e8f0; border-radius:16px; padding:18px 20px; box-shadow:0 5px 18px rgba(15,23,42,.08); }
        .target-kpi-title { font-size:13px; color:#64748b; font-weight:650; margin-bottom:12px; }
        .target-kpi-value { font-size:32px; line-height:1.1; font-weight:760; }
        .target-status { display:flex; align-items:center; gap:10px; border-radius:12px; padding:12px 16px; color:#fff; font-weight:750; letter-spacing:.04em; margin:12px 0 20px; }
        .target-status-dot { background:#fff; width:12px; height:12px; border-radius:50%; box-shadow:0 0 0 4px rgba(255,255,255,.2); }
        .target-week-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:8px 0 22px; }
        .target-week { border:1px solid #e2e8f0; background:#fff; border-radius:12px; padding:14px; box-shadow:0 3px 12px rgba(15,23,42,.05); }
        .target-week-title { color:#64748b; font-size:12px; font-weight:700; }
        .target-week-value { color:#0f172a; font-size:20px; font-weight:750; margin:6px 0; }
        .target-week-note { color:#64748b; font-size:11px; }
        .target-person { background:#fff; border:1px solid #e2e8f0; border-radius:12px; padding:13px 15px; margin:0 0 10px; box-shadow:0 2px 10px rgba(15,23,42,.04); }
        .target-person-head { display:flex; justify-content:space-between; gap:10px; align-items:center; }
        .target-person-name { font-size:13px; font-weight:700; color:#0f172a; }
        .target-person-value { font-size:12px; color:#475569; white-space:nowrap; }
        .target-track { height:8px; background:#e2e8f0; border-radius:99px; overflow:hidden; margin:9px 0 5px; }
        .target-fill { height:100%; border-radius:99px; }
        .target-person-foot { display:flex; justify-content:space-between; color:#64748b; font-size:11px; }
        @media(max-width:800px) { .target-kpis,.target-week-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } }
        @media(max-width:460px) { .target-kpis,.target-week-grid { grid-template-columns:1fr; } .target-person-head { align-items:flex-start; flex-direction:column; } }
        </style>
        """,
        unsafe_allow_html=True,
    )
    kpis_html = "".join(
        "<div class='target-kpi'>"
        f"<div class='target-kpi-title'>{titulo}</div>"
        f"<div class='target-kpi-value' style='color:{cor}'>{valor}</div>"
        "</div>"
        for titulo, valor, cor in metricas
    )
    st.markdown(f"<div class='target-kpis'>{kpis_html}</div>", unsafe_allow_html=True)
    st.progress(min(1.0, max(0.0, atingimento)), text=f"Progresso mensal · {atingimento:.1%}")
    st.markdown(
        f"<div class='target-status' style='background:{cor}'>"
        f"<span class='target-status-dot'></span>{situacao}</div>",
        unsafe_allow_html=True,
    )

    semanas = [
        coluna.removeprefix("Target ")
        for coluna in dados.columns
        if coluna.startswith("Target WK ")
    ]
    st.subheader("Ritmo semanal")
    semanas_html = []
    for semana in semanas:
        target_semana = float(dados[f"Target {semana}"].sum())
        realizado_semana = float(dados[f"Realizado {semana}"].sum())
        percentual_semana = realizado_semana / target_semana if target_semana > 0 else 0.0
        semanas_html.append(
            "<div class='target-week'>"
            f"<div class='target-week-title'>{semana}</div>"
            f"<div class='target-week-value'>{realizado_semana:,.0f} / {target_semana:,.0f}</div>"
            f"<div class='target-week-note'>{percentual_semana:.0%} do target semanal realizado</div>"
            "</div>"
        )
    st.markdown(
        f"<div class='target-week-grid'>{''.join(semanas_html)}</div>",
        unsafe_allow_html=True,
    )

    st.subheader(f"Distribuição por {nome_grupo[:-1] if nome_grupo.endswith('s') else nome_grupo}")
    dados = dados.sort_values(["Atingimento", "Responsável"], ascending=[True, True])
    coluna_esquerda, coluna_direita = st.columns(2)
    for indice, (_, linha) in enumerate(dados.iterrows()):
        percentual = float(linha["Atingimento"])
        cor_pessoa = "#dc2626" if percentual < 0.8 else "#d97706" if percentual < 1 else "#15803d"
        largura = min(100.0, max(0.0, percentual * 100))
        card_html = (
            "<div class='target-person'><div class='target-person-head'>"
            f"<span class='target-person-name'>{linha['Responsável']}</span>"
            f"<span class='target-person-value'>{linha['Total Realizado Calculado']:,.0f} / "
            f"{linha['Target Mensal Calculado']:,.1f} · {percentual:.0%}</span>"
            "</div><div class='target-track'>"
            f"<div class='target-fill' style='width:{largura:.1f}%;background:{cor_pessoa}'></div>"
            "</div><div class='target-person-foot'>"
            f"<span>Turno {linha['Turno']}</span>"
            f"<span>Restante {max(0.0, linha['Pendente Calculado']):,.0f}</span>"
            "</div></div>"
        )
        with (coluna_esquerda if indice % 2 == 0 else coluna_direita):
            st.markdown(card_html, unsafe_allow_html=True)


def exibir_dashboard_operacional(dataframe, colunas, meses_disponiveis):
    mes_selecionado = st.selectbox(
        "📅 Mês de análise",
        meses_disponiveis,
        index=meses_disponiveis.index(mes_padrao_dashboard(meses_disponiveis)),
        key="mes_analise_compartilhado",
    )
    meta_mensal = st.number_input(
        "Target Mensal (configurado manualmente)",
        min_value=1,
        value=560,
        step=1,
        key="target_mensal_operacional",
    )
    df_mes = dataframe[dataframe[colunas["mes"]] == mes_selecionado].copy()
    df_mes["_situacao_prazo_dashboard"] = calcular_situacao_prazo(
        df_mes, colunas["status"], colunas["data_prazo"]
    )
    status_mes = df_mes[colunas["status"]].map(normalizar_texto_dashboard)
    total_realizado = len(df_mes)
    total_aberto = int((status_mes == "ABERTO").sum())
    total_nao_reconhecido = int(
        (~status_mes.isin(["CONCLUIDO", "ABERTO"])).sum()
    )
    percentual_realizado = total_realizado / meta_mensal
    faltam_para_meta = max(0, meta_mensal - total_realizado)

    st.caption(
        f"Conferência da aba FORMS — {mes_selecionado}: {len(df_mes)} lançamento(s) "
        f"= {int((status_mes == 'CONCLUIDO').sum())} concluído(s) + {total_aberto} aberto(s) "
        f"+ {total_nao_reconhecido} com status vazio ou não reconhecido."
    )
    if total_nao_reconhecido:
        st.warning(
            f"{total_nao_reconhecido} lançamento(s) não foram somados como concluídos "
            "nem abertos porque o STATUS está vazio ou não reconhecido."
        )

    st.markdown(
        "<div class='header-blue'>Gestão avançada de indicadores da qualidade</div>",
        unsafe_allow_html=True,
    )
    col_esquerda, col_direita = st.columns([3, 1.5])
    with col_esquerda:
        card1, card2, card3 = st.columns(3)
        with card1:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-title title-programado'>🎯 Total Programado Mês</div>"
                f"<div class='kpi-value val-programado'>{meta_mensal}</div></div>",
                unsafe_allow_html=True,
            )
        with card2:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-title title-realizado'>📋 Total Realizado Mês (qtd.)</div>"
                f"<div class='kpi-value val-realizado'>{total_realizado}</div></div>",
                unsafe_allow_html=True,
            )
        with card3:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-title title-porcentagem'>📊 Total Realizado Mês (%)</div>"
                f"<div class='kpi-value val-porcentagem'>{int(percentual_realizado * 100)}%</div></div>",
                unsafe_allow_html=True,
            )

        if percentual_realizado < 0.8:
            st.markdown(
                "<div class='banner-critico'><span>⚠️ ATENÇÃO: DESEMPENHO CRÍTICO</span>"
                "<span>Meta de atingimento: ≥ 80%</span></div>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                "<div class='banner-critico' style='background-color:#e6f4ea; "
                "border-color:#b7e1cd; color:#137333;'><span>✅ DESEMPENHO DENTRO DA META</span>"
                "<span>Meta de atingimento: ≥ 80%</span></div>",
                unsafe_allow_html=True,
            )

        st.metric(label="📌 Ações em aberto", value=total_aberto)
        historico = dataframe[
            dataframe[colunas["mes"]].astype(str).str.strip().ne("")
        ].copy()
        historico["_mes_normalizado"] = historico[colunas["mes"]].apply(
            normalizar_texto_dashboard
        )
        historico["_situacao"] = calcular_situacao_prazo(
            historico, colunas["status"], colunas["data_prazo"]
        ).replace(
            {
                "Concluída": "Concluído",
                "Vencida": "Aberto e vencido",
                "No prazo": "Aberto no prazo",
            }
        )
        ordem_meses = sorted(
            historico["_mes_normalizado"].unique(),
            key=lambda mes: MONTH_ORDER.get(mes, 999),
        )
        ordem_situacoes = ["Concluído", "Aberto e vencido", "Aberto no prazo"]
        st.subheader("📈 Ações por mês e situação")
        mes_grafico = st.selectbox(
            "Filtrar gráfico por mês",
            ["Todos os meses", *ordem_meses],
            key="historico_mes_grafico",
        )
        historico_grafico = historico
        if mes_grafico != "Todos os meses":
            historico_grafico = historico[historico["_mes_normalizado"] == mes_grafico]
        meses_grafico = (
            [mes_grafico] if mes_grafico != "Todos os meses" else ordem_meses
        )
        historico_por_mes = (
            historico_grafico.groupby(["_mes_normalizado", "_situacao"])
            .size()
            .unstack(fill_value=0)
            .reindex(index=meses_grafico, columns=ordem_situacoes, fill_value=0)
        )
        historico_fig = go.Figure()
        cores_situacoes = {
            "Concluído": "#00a65a",
            "Aberto e vencido": "#d9534f",
            "Aberto no prazo": "#f39c12",
        }
        for situacao in ordem_situacoes:
            historico_fig.add_trace(
                go.Bar(
                    name=situacao,
                    x=historico_por_mes.index,
                    y=historico_por_mes[situacao],
                    marker_color=cores_situacoes[situacao],
                    text=historico_por_mes[situacao],
                    textposition="outside",
                    cliponaxis=False,
                    hovertemplate=(
                        "Mês: %{x}<br>Situação: "
                        + situacao
                        + "<br>Quantidade: %{y}<extra></extra>"
                    ),
                )
            )
        historico_fig.update_layout(
            barmode="group",
            xaxis=dict(categoryorder="array", categoryarray=meses_grafico),
            yaxis_title="Quantidade de ações",
            legend_title="Situação",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#f7f9fc"),
            margin=dict(t=20, b=20, l=10, r=10),
            height=400,
        )
        st.plotly_chart(historico_fig, use_container_width=True)

    with col_direita:
        st.markdown(
            "<div class='chart-box-header'>Desempenho do Mês</div>",
            unsafe_allow_html=True,
        )
        fig_rosca = go.Figure(
            data=[
                go.Pie(
                    labels=["Atingimento", "Restante"],
                    values=[
                        min(percentual_realizado, 1.0),
                        max(0.0, 1.0 - percentual_realizado),
                    ],
                    hole=0.75,
                    marker_colors=["#f39c12", "#e0e0e0"],
                    textinfo="none",
                    showlegend=False,
                )
            ]
        )
        fig_rosca.update_layout(
            height=180,
            margin=dict(t=0, b=0, l=0, r=0),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#f7f9fc"),
            annotations=[
                dict(
                    text=f"<b>{int(percentual_realizado * 100)}%</b><br>"
                    "<span style='font-size:10px;'>ATINGIMENTO</span>",
                    x=0.5,
                    y=0.5,
                    font_size=16,
                    showarrow=False,
                )
            ],
        )
        st.plotly_chart(fig_rosca, use_container_width=True)
        st.markdown(
            f"<div class='meta-box'><span style='font-size:10px; font-weight:bold;'>"
            f"FALTAM PARA A META</span><br><span style='font-size:22px; font-weight:bold; "
            f"color:#d93025;'>{faltam_para_meta}</span></div>",
            unsafe_allow_html=True,
        )
    return mes_selecionado, df_mes


def exibir_visao_executiva_acoes(dataframe, colunas, meses_disponiveis):
    hoje = date.today()
    mes_selecionado = mes_padrao_dashboard(meses_disponiveis, hoje)
    df_mes = dataframe[dataframe[colunas["mes"]] == mes_selecionado].copy()
    situacoes = calcular_situacao_prazo(
        df_mes, colunas["status"], colunas["data_prazo"], hoje
    )
    indicadores = calcular_indicadores_acoes(
        df_mes,
        colunas["status"],
        situacoes,
        560,
        mes_selecionado,
        hoje,
    )
    atingimento = indicadores["percentual_atingimento"]
    if atingimento < 0.8:
        semaforo_cor = "#dc2626"
        semaforo_texto = "ABAIXO DA META"
    elif atingimento < 1:
        semaforo_cor = "#d97706"
        semaforo_texto = "ATENÇÃO"
    else:
        semaforo_cor = "#15803d"
        semaforo_texto = "META ATINGIDA"

    kpis = [
        ("Meta Mês", f"{indicadores['meta_total']:,.0f}", "#2563eb"),
        ("Realizado", f"{indicadores['total_realizado']:,}", "#0f766e"),
        ("Faltam para Meta", f"{indicadores['falta_para_meta']:,.0f}", "#b45309"),
        ("% de Atingimento", f"{atingimento:.0%}", semaforo_cor),
        ("Ações Concluídas", f"{indicadores['concluidas']:,}", "#15803d"),
        ("Ações em Aberto", f"{indicadores['em_aberto']:,}", "#475569"),
        ("Ações Vencidas", f"{indicadores['vencidas']:,}", "#dc2626"),
        ("Ações no Prazo", f"{indicadores['no_prazo']:,}", "#0284c7"),
    ]
    st.markdown(
        """
        <style>
        .qmax-executive-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin:10px 0 20px; }
        .qmax-executive-card { background:#fff; border:1px solid #e2e8f0; border-radius:14px; padding:18px 20px; box-shadow:0 4px 14px rgba(15,23,42,.07); min-height:112px; }
        .qmax-executive-label { color:#64748b; font-size:13px; font-weight:650; margin-bottom:12px; }
        .qmax-executive-value { font-size:32px; line-height:1.1; font-weight:750; }
        .qmax-executive-traffic { display:flex; align-items:center; gap:10px; padding:14px 18px; border-radius:12px; color:#fff; font-weight:750; letter-spacing:.04em; }
        .qmax-executive-light { width:13px; height:13px; border-radius:50%; background:#fff; box-shadow:0 0 0 4px rgba(255,255,255,.2); }
        @media(max-width:850px) { .qmax-executive-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } }
        @media(max-width:480px) { .qmax-executive-grid { grid-template-columns:1fr; } }
        </style>
        """,
        unsafe_allow_html=True,
    )
    cards_html = "".join(
        "<div class='qmax-executive-card'>"
        f"<div class='qmax-executive-label'>{rotulo}</div>"
        f"<div class='qmax-executive-value' style='color:{cor}'>{valor}</div>"
        "</div>"
        for rotulo, valor, cor in kpis
    )
    st.markdown(
        f"<div class='qmax-executive-grid'>{cards_html}</div>",
        unsafe_allow_html=True,
    )
    st.progress(min(1.0, max(0.0, atingimento)), text=f"Progresso da meta mensal — {atingimento:.0%}")
    st.markdown(
        f"<div class='qmax-executive-traffic' style='background:{semaforo_cor}'>"
        f"<span class='qmax-executive-light'></span> {semaforo_texto}"
        "</div>",
        unsafe_allow_html=True,
    )
    return mes_selecionado, df_mes, indicadores


def exibir_relatorios_acoes(dataframe, colunas, meses_disponiveis):
    mes_selecionado = st.selectbox(
        "Mês do relatório",
        meses_disponiveis,
        index=meses_disponiveis.index(mes_padrao_dashboard(meses_disponiveis)),
        key="mes_relatorio_acoes",
    )
    df_mes = dataframe[dataframe[colunas["mes"]] == mes_selecionado].copy()
    situacoes = calcular_situacao_prazo(
        df_mes, colunas["status"], colunas["data_prazo"]
    )
    df_vencidos = df_mes[situacoes == "Vencida"].copy()
    st.subheader(f"Ações em aberto e vencidas por setor — {mes_selecionado}")
    if df_vencidos.empty:
        st.success("Nenhuma ação em aberto e vencida foi identificada.")
        return

    setor_normalizado = df_vencidos[colunas["setor"]].apply(
        lambda setor: " ".join(str(setor).split()).casefold() or "(Sem setor)"
    )
    df_vencidos = df_vencidos.assign(_setor_normalizado=setor_normalizado)
    agrupado = (
        df_vencidos.groupby("_setor_normalizado", as_index=False)
        .agg(**{"Ações abertas e vencidas": (colunas["status"], "size")})
    )
    nomes_setores = (
        df_vencidos.groupby("_setor_normalizado")[colunas["setor"]].first().to_dict()
    )
    agrupado["Setor"] = agrupado["_setor_normalizado"].map(
        lambda setor: str(nomes_setores.get(setor, setor)).strip() or "(Sem setor)"
    )
    email_enviado = df_vencidos[colunas["email"]].apply(
        lambda valor: normalizar_texto_dashboard(valor).replace("-", " ") == "E MAIL ENVIADO"
    )
    notificados_por_setor = (
        email_enviado.groupby(df_vencidos["_setor_normalizado"]).sum().to_dict()
    )
    agrupado["% Notificado via E-mail"] = agrupado.apply(
        lambda linha: 100
        * notificados_por_setor.get(linha["_setor_normalizado"], 0)
        / linha["Ações abertas e vencidas"],
        axis=1,
    )
    agrupado = agrupado.sort_values("Ações abertas e vencidas", ascending=True)
    total_vencidas = len(df_vencidos)
    taxa_email = 100 * int(email_enviado.sum()) / total_vencidas
    metrica_total, metrica_email = st.columns(2)
    metrica_total.metric("Ações abertas e vencidas", total_vencidas)
    metrica_email.metric("% notificadas por e-mail", f"{taxa_email:.1f}%")

    figura = go.Figure(
        data=[
            go.Bar(
                x=agrupado["Ações abertas e vencidas"],
                y=agrupado["Setor"],
                orientation="h",
                marker_color="#d9534f",
                text=agrupado["Ações abertas e vencidas"],
                textposition="outside",
                customdata=agrupado["% Notificado via E-mail"],
                hovertemplate=(
                    "Setor: %{y}<br>Ações abertas e vencidas: %{x}"
                    "<br>Notificadas por e-mail: %{customdata:.1f}%<extra></extra>"
                ),
            )
        ]
    )
    figura.update_layout(
        title="Ações em aberto e vencidas por setor",
        xaxis_title="Quantidade de ações",
        yaxis_title="Setor",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#f7f9fc"),
        margin=dict(t=50, b=20, l=10, r=40),
        height=max(350, 48 * len(agrupado)),
    )
    st.plotly_chart(figura, use_container_width=True)
    st.caption(
        "Uma ação é considerada notificada quando a célula de e-mail contém "
        "“E-MAIL ENVIADO”; qualquer outro texto é desconsiderado."
    )
