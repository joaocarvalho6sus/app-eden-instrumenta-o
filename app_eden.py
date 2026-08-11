"""
=============================================================================
 THE EDEN, ESTORIL — APLICACAO DE ANALISE DE INSTRUMENTACAO GEOTECNICA
 Dissertacao de Mestrado — back-analysis da contencao periferica
=============================================================================

 O QUE E ESTE FICHEIRO
 ---------------------
 E uma aplicacao Streamlit. NAO se corre colando no Python normal nem num
 "Python online". O Streamlit le este ficheiro e transforma-o numa pagina
 web interativa. Ver instrucoes em "COMO CORRER" mais abaixo.

 O QUE A APLICACAO FAZ
 ---------------------
   - Inclinometros:
       - perfil deformado por leitura (profundidade no eixo vertical);
       - evolucao do deslocamento a PROFUNDIDADE FIXA e no MAXIMO GLOBAL;
       - velocidade de deslocamento (mm/dia) entre leituras;
       - detecao automatica de aceleracoes / sinais precursores.
   - Celulas de carga:
       - carga atual vs. blocagem;
       - variacao (%) face a blocagem, com limiares de alerta e alarme.
   - Piezometros:
       - evolucao da cota da agua.

 COMO CORRER (uma so vez, a instalacao)
 --------------------------------------
   1) Instalar o Python de python.org — MARCAR "Add Python to PATH".
   2) Abrir o Prompt de Comando e instalar as bibliotecas:
          pip install streamlit pandas numpy plotly openpyxl
   3) Colocar ESTE ficheiro e o Excel na MESMA pasta.

 COMO CORRER (sempre que quiseres usar)
 --------------------------------------
   No Prompt de Comando, dentro da pasta:
          streamlit run app_eden.py
   A aplicacao abre sozinha no navegador.

 COMO ADAPTAR
 ------------
 Toda a configuracao — nome do ficheiro, nomes das folhas, nomes das
 colunas e criterios de alerta — esta reunida no bloco CONFIGURACAO logo
 a seguir aos imports. Para adaptar a app a outro ficheiro ou a colunas
 com outros nomes, muda-se SO esse bloco; a logica por baixo nao precisa
 de ser tocada.
=============================================================================
"""

# -------------------------------------------------------------------------
# IMPORTS
# -------------------------------------------------------------------------
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# =========================================================================
# CONFIGURACAO  —  e aqui que se adapta a app; nao e preciso mexer no resto
# =========================================================================

# Nome do ficheiro Excel que fica ao lado deste script.
FICHEIRO_EXCEL = "Modelo_Dados_Instrumentacao_The_Eden_2.xlsx"

# Nomes das folhas dentro do Excel.
FOLHAS = {
    "resumo": "Inclinometros_Resumo",
    "perfis": "Inclinometros_Perfis",
    "celulas": "Celulas_Carga",
    "piezo": "Piezometros",
}

# Nomes das colunas usadas em cada folha.
# Se um dia os relatorios vierem com cabecalhos ligeiramente diferentes,
# corrige-se aqui — uma vez — e a app toda passa a funcionar.
COLS = {
    "data": "Data",
    "inclinometro": "Inclinómetro",
    "profundidade": "Profundidade (m)",
    "desl_total": "Desl. acumulado total (mm)",
    "desl_max_global": "Máx. desloc. acumulado total (mm)",
    "prof_do_max": "Profundidade do máximo (m)",
    # celulas de carga
    "celula": "Célula",
    "ancoragem": "Ancoragem",
    "carga_atual": "Carga atual (kN)",
    "blocagem": "Blocagem (kN)",
    "variacao": "Variação calculada",
    "estado": "Estado",
    # piezometros
    "piezometro": "Piezómetro",
    "cota_agua": "Cota da água calculada (m)",
    "cota_boca": "Cota da boca (m)",
    "prof_abaixo_boca": "Profundidade abaixo da boca (m)",
}

# Criterios das celulas de carga (variacao face a blocagem), retirados do
# dicionario de dados do proprio Excel: Regular <15% ; Alerta 15-25% ; Alarme >25%.
CC_ALERTA = 0.15
CC_ALARME = 0.25

# Valores por defeito da detecao de precursores (ajustaveis na app).
LIMIAR_VEL_DEFEITO = 0.5   # mm/dia
FATOR_ACEL_DEFEITO = 1.8   # multiplo da velocidade anterior


# =========================================================================
# CARREGAMENTO DE DADOS
# =========================================================================
@st.cache_data(show_spinner="A carregar o Excel...")
def carregar_dados(fonte):
    """
    Le as quatro folhas relevantes do Excel e converte a coluna de data.

    'fonte' pode ser:
       - um caminho/nome de ficheiro (str), ou
       - um ficheiro carregado pelo utilizador (file uploader do Streamlit).

    Devolve um dicionario de DataFrames: resumo, perfis, celulas, piezo.
    """
    excel = pd.ExcelFile(fonte)
    dados = {}
    for chave, nome_folha in FOLHAS.items():
        df = pd.read_excel(excel, nome_folha)
        if COLS["data"] in df.columns:
            df[COLS["data"]] = pd.to_datetime(df[COLS["data"]], errors="coerce")
        dados[chave] = df
    return dados


def validar_colunas(df, nomes, contexto):
    """
    Verifica se as colunas esperadas existem no DataFrame.
    Se faltar alguma, mostra um aviso claro em vez de rebentar com um
    erro tecnico — ajuda a perceber que o Excel tem outra estrutura.
    """
    em_falta = [n for n in nomes if n not in df.columns]
    if em_falta:
        st.error(
            f"Na seccao '{contexto}' faltam as colunas: {em_falta}. "
            f"Confirma os nomes no bloco CONFIGURACAO (dicionario COLS) "
            f"ou no proprio Excel."
        )
        return False
    return True


# =========================================================================
# CALCULO DE VELOCIDADE E DETECAO DE PRECURSORES
# =========================================================================
def calcular_velocidade(datas, valores, limiar_vel, fator_acel):
    """
    Calcula a velocidade de deslocamento entre leituras consecutivas e
    identifica sinais precursores.

    Parametros
    ----------
    datas      : sequencia de datas das leituras.
    valores    : sequencia de deslocamentos (mm), alinhada com 'datas'.
    limiar_vel : velocidade (mm/dia) acima da qual se marca precursor.
    fator_acel : se a velocidade for este multiplo da anterior, marca-se
                 precursor mesmo que abaixo do limiar (deteta aceleracoes).

    Devolve
    -------
    DataFrame com colunas: data, valor, dias, delta, velocidade, precursor.
    """
    df = pd.DataFrame({"data": pd.to_datetime(list(datas)),
                       "valor": list(valores)})
    df = df.sort_values("data").reset_index(drop=True)

    df["dias"] = df["data"].diff().dt.days
    df["delta"] = df["valor"].diff()

    df["velocidade"] = np.where(
        df["dias"] > 0, (df["delta"] / df["dias"]).round(3), np.nan
    )

    df["precursor"] = False
    for i in range(1, len(df)):
        v = df.loc[i, "velocidade"]
        v_ant = df.loc[i - 1, "velocidade"]

        acima_do_limiar = pd.notna(v) and v >= limiar_vel
        acelerou = (pd.notna(v) and pd.notna(v_ant)
                    and v_ant > 0 and v >= fator_acel * v_ant)

        df.loc[i, "precursor"] = bool(acima_do_limiar or acelerou)

    return df


# =========================================================================
# COMPONENTES DA INTERFACE  (uma funcao por separador, para ficar legivel)
# =========================================================================
def separador_inclinometros(dados, limiar_vel, fator_acel):
    perfis = dados["perfis"]
    resumo = dados["resumo"]

    ok = validar_colunas(
        perfis,
        [COLS["data"], COLS["inclinometro"], COLS["profundidade"], COLS["desl_total"]],
        "Inclinometros / perfis",
    ) and validar_colunas(
        resumo,
        [COLS["data"], COLS["inclinometro"], COLS["desl_max_global"], COLS["prof_do_max"]],
        "Inclinometros / resumo",
    )
    if not ok:
        return

    lista_inc = sorted(perfis[COLS["inclinometro"]].dropna().unique())
    inc = st.selectbox("Inclinometro", lista_inc)

    p_inc = perfis[perfis[COLS["inclinometro"]] == inc].copy()
    r_inc = resumo[resumo[COLS["inclinometro"]] == inc].copy().sort_values(COLS["data"])
    datas_inc = sorted(p_inc[COLS["data"]].dropna().unique())

    if not datas_inc:
        st.warning("Este inclinometro nao tem leituras com data valida.")
        return

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Perfil deformado")
        st.caption("Deslocamento acumulado ao longo da profundidade. "
                   "A base da sonda e assumida fixa.")
        indices_default = [0, len(datas_inc) // 2, len(datas_inc) - 1]
        datas_default = [datas_inc[i] for i in sorted(set(indices_default))]
        datas_sel = st.multiselect(
            "Leituras a mostrar",
            options=datas_inc,
            default=datas_default,
            format_func=lambda d: pd.to_datetime(d).strftime("%d/%m/%Y"),
        )
        fig = go.Figure()
        for d in datas_sel:
            sub = p_inc[p_inc[COLS["data"]] == d].sort_values(COLS["profundidade"])
            fig.add_trace(go.Scatter(
                x=sub[COLS["desl_total"]],
                y=sub[COLS["profundidade"]],
                mode="lines+markers",
                name=pd.to_datetime(d).strftime("%d/%m/%Y"),
            ))
        fig.update_yaxes(autorange="reversed", title="Profundidade (m)")
        fig.update_xaxes(title="Deslocamento acumulado total (mm)")
        fig.update_layout(height=560, legend_title="Leitura")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Evolucao do deslocamento")
        st.caption("Comparacao entre seguir o maximo global e seguir uma "
                   "profundidade fixa — uteis para justificar a escolha na tese.")

        profs = sorted(p_inc[COLS["profundidade"]].dropna().unique())
        moda_prof = r_inc[COLS["prof_do_max"]].mode()
        prof_default = float(moda_prof.iloc[0]) if len(moda_prof) else profs[0]
        if prof_default not in profs:
            prof_default = profs[0]

        prof_fixa = st.select_slider(
            "Profundidade fixa a seguir (m)", options=profs, value=prof_default
        )

        serie_fixa = p_inc[p_inc[COLS["profundidade"]] == prof_fixa].sort_values(COLS["data"])
        serie_max = r_inc.sort_values(COLS["data"])

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=serie_max[COLS["data"]], y=serie_max[COLS["desl_max_global"]],
            mode="lines+markers", name="Maximo global",
        ))
        fig2.add_trace(go.Scatter(
            x=serie_fixa[COLS["data"]], y=serie_fixa[COLS["desl_total"]],
            mode="lines+markers", name=f"A {prof_fixa:.1f} m",
        ))
        fig2.update_xaxes(title="Data de leitura")
        fig2.update_yaxes(title="Deslocamento acumulado (mm)")
        fig2.update_layout(height=560, legend_title="Serie")
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()
    st.subheader("Velocidade e sinais precursores")

    escolha_base = st.radio(
        "Serie de base para a velocidade",
        ["Maximo global", f"Profundidade fixa ({prof_fixa:.1f} m)"],
        horizontal=True,
    )
    if escolha_base == "Maximo global":
        vdf = calcular_velocidade(serie_max[COLS["data"]],
                                  serie_max[COLS["desl_max_global"]],
                                  limiar_vel, fator_acel)
    else:
        vdf = calcular_velocidade(serie_fixa[COLS["data"]],
                                  serie_fixa[COLS["desl_total"]],
                                  limiar_vel, fator_acel)

    fig3 = go.Figure()
    fig3.add_trace(go.Bar(
        x=vdf["data"], y=vdf["velocidade"],
        marker_color=["crimson" if p else "steelblue" for p in vdf["precursor"]],
    ))
    fig3.add_hline(y=limiar_vel, line_dash="dash", line_color="crimson",
                   annotation_text="Limiar")
    fig3.update_xaxes(title="Data")
    fig3.update_yaxes(title="Velocidade (mm/dia)")
    fig3.update_layout(height=380)
    st.plotly_chart(fig3, use_container_width=True)

    n_prec = int(vdf["precursor"].sum())
    c1, c2, c3 = st.columns(3)
    ultimo_max = serie_max[COLS["desl_max_global"]].iloc[-1] if len(serie_max) else np.nan
    c1.metric("Ult. desloc. max. (mm)", f"{ultimo_max:.2f}")
    vel_max = vdf["velocidade"].max()
    c2.metric("Velocidade max. (mm/dia)", f"{vel_max:.3f}" if pd.notna(vel_max) else "-")
    c3.metric("Precursores", n_prec)

    if n_prec:
        st.warning(f"{n_prec} leitura(s) com aceleracao acima dos criterios definidos.")
    else:
        st.success("Nenhuma aceleracao acima dos criterios definidos.")

    tabela = vdf.rename(columns={
        "data": "Data", "valor": "Desloc. (mm)", "dias": "Dias",
        "delta": "Delta (mm)", "velocidade": "Vel. (mm/dia)", "precursor": "Precursor",
    })
    st.dataframe(
        tabela.style.apply(
            lambda linha: ["background-color:#ffe0e0" if linha["Precursor"] else ""
                           for _ in linha],
            axis=1,
        ),
        use_container_width=True,
    )


def separador_celulas(dados):
    cc = dados["celulas"]
    if not validar_colunas(
        cc, [COLS["data"], COLS["celula"], COLS["carga_atual"], COLS["variacao"]],
        "Celulas de carga",
    ):
        return

    cc = cc.sort_values(COLS["data"])
    st.subheader("Celulas de carga — evolucao")

    cel = st.selectbox("Celula", sorted(cc[COLS["celula"]].dropna().unique()))
    sub = cc[cc[COLS["celula"]] == cel].sort_values(COLS["data"])

    col1, col2 = st.columns(2)
    with col1:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=sub[COLS["data"]], y=sub[COLS["carga_atual"]],
                                 mode="lines+markers", name="Carga atual"))
        if COLS["blocagem"] in sub.columns:
            fig.add_trace(go.Scatter(x=sub[COLS["data"]], y=sub[COLS["blocagem"]],
                                     mode="lines", line_dash="dot", name="Blocagem"))
        fig.update_xaxes(title="Data")
        fig.update_yaxes(title="Carga (kN)")
        fig.update_layout(height=460)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=sub[COLS["data"]], y=sub[COLS["variacao"]] * 100,
                                  mode="lines+markers", name="Variacao (%)"))
        fig2.add_hline(y=CC_ALERTA * 100, line_dash="dash", line_color="orange",
                       annotation_text="Alerta 15%")
        fig2.add_hline(y=CC_ALARME * 100, line_dash="dash", line_color="red",
                       annotation_text="Alarme 25%")
        fig2.update_xaxes(title="Data")
        fig2.update_yaxes(title="Variacao face a blocagem (%)")
        fig2.update_layout(height=460)
        st.plotly_chart(fig2, use_container_width=True)

    colunas_tabela = [c for c in [COLS["data"], COLS["ancoragem"], COLS["carga_atual"],
                                  COLS["blocagem"], COLS["variacao"], COLS["estado"]]
                      if c in sub.columns]
    st.dataframe(sub[colunas_tabela], use_container_width=True)


def separador_piezometros(dados):
    pz = dados["piezo"]
    if not validar_colunas(
        pz, [COLS["data"], COLS["piezometro"], COLS["cota_agua"]], "Piezometros"
    ):
        return

    pz = pz.sort_values(COLS["data"])
    st.subheader("Piezometros — cota da agua")

    p = st.selectbox("Piezometro", sorted(pz[COLS["piezometro"]].dropna().unique()))
    sub = pz[pz[COLS["piezometro"]] == p].sort_values(COLS["data"])

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=sub[COLS["data"]], y=sub[COLS["cota_agua"]],
                             mode="lines+markers", name="Cota da agua"))
    fig.update_xaxes(title="Data")
    fig.update_yaxes(title="Cota da agua (m)")
    fig.update_layout(height=460)
    st.plotly_chart(fig, use_container_width=True)

    st.caption("Sugestao para a tese: sobrepor a precipitacao diaria a esta serie "
               "para avaliar a resposta do nivel freatico a pluviosidade.")
    colunas_tabela = [c for c in [COLS["data"], COLS["cota_boca"],
                                  COLS["prof_abaixo_boca"], COLS["cota_agua"]]
                      if c in sub.columns]
    st.dataframe(sub[colunas_tabela], use_container_width=True)


# =========================================================================
# PROGRAMA PRINCIPAL
# =========================================================================
def main():
    st.set_page_config(page_title="The Eden — Instrumentacao", layout="wide")

    st.sidebar.title("The Eden - Instrumentacao")

    fonte_opcao = st.sidebar.radio(
        "Ficheiro de dados",
        ["Usar o ficheiro ao lado do script", "Carregar manualmente"],
    )

    if fonte_opcao == "Carregar manualmente":
        carregado = st.sidebar.file_uploader("Excel de instrumentacao", type=["xlsx"])
        if carregado is None:
            st.info("Carrega o Excel na barra lateral para comecar.")
            st.stop()
        fonte = carregado
    else:
        fonte = FICHEIRO_EXCEL
        if not Path(FICHEIRO_EXCEL).exists():
            st.error(
                f"Nao encontrei '{FICHEIRO_EXCEL}' na pasta do script. "
                f"Coloca o Excel na mesma pasta, ou usa 'Carregar manualmente'."
            )
            st.stop()

    try:
        dados = carregar_dados(fonte)
    except Exception as erro:
        st.error(f"Nao consegui ler o Excel. Detalhe tecnico: {erro}")
        st.stop()

    st.sidebar.divider()
    st.sidebar.subheader("Detecao de precursores")
    limiar_vel = st.sidebar.slider("Limiar de velocidade (mm/dia)",
                                   0.1, 3.0, LIMIAR_VEL_DEFEITO, 0.05)
    fator_acel = st.sidebar.slider("Fator de aceleracao (x)",
                                   1.2, 3.0, FATOR_ACEL_DEFEITO, 0.1)

    st.title("The Eden, Estoril — Analise de Instrumentacao")
    st.caption("Back-analysis da contencao periferica - perfis, velocidades e "
               "sinais precursores. A integracao dos perfis assume a base fixa.")

    aba_inc, aba_cc, aba_pz = st.tabs(
        ["Inclinometros", "Celulas de carga", "Piezometros"]
    )
    with aba_inc:
        separador_inclinometros(dados, limiar_vel, fator_acel)
    with aba_cc:
        separador_celulas(dados)
    with aba_pz:
        separador_piezometros(dados)


# Em Streamlit o ficheiro e executado de cima a baixo; esta guarda mantem
# o comportamento correto e e boa pratica em Python.
if __name__ == "__main__":
    main()
