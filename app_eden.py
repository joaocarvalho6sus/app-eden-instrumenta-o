"""
=============================================================================
 THE EDEN, ESTORIL — APLICACAO DE ANALISE DE INSTRUMENTACAO GEOTECNICA
 Dissertacao de Mestrado — back-analysis da contencao periferica
=============================================================================

 O QUE E ESTE FICHEIRO
 ---------------------
 Aplicacao Streamlit. NAO se corre colando no Python normal nem num
 "Python online": o Streamlit le este ficheiro e transforma-o numa pagina
 web. Ver "COMO CORRER".

 SEPARADORES
 -----------
   1. Visao geral 3D  — alvos topograficos no espaco real (M,P,Z), com
                        deslocamento amplificado, cor por magnitude e
                        evolucao por campanha. Superficie interpolada
                        OPCIONAL, sempre com os pontos medidos por cima.
   2. Inclinometros   — perfil deformado, evolucao, velocidade, precursores.
   3. Alvos (2D)      — series temporais de deslocamento horizontal/vertical.
   4. Celulas de carga— carga vs. blocagem e variacao (%) com alerta/alarme.
   5. Piezometros     — cota da agua.

 NOTA METODOLOGICA (importante para a defesa)
 --------------------------------------------
 - Os ALVOS TOPOGRAFICOS tem coordenadas reais (M,P,Z) -> podem ir para 3D.
 - Os INCLINOMETROS nao tem coordenadas em planta neste ficheiro (so azimute
   e profundidade), por isso NAO sao colocados no 3D dos alvos — seria
   inventar a posicao. Ficam na sua vista propria. Se um dia tiveres as
   coordenadas em planta, integram-se com vetores orientados pelo azimute.
 - A superficie 3D e INTERPOLACAO entre alvos: mostra-se so como apoio
   visual e com os pontos medidos sempre visiveis. E uma leitura do campo
   de deslocamentos dos alvos, nao um modelo do macico.

 COMO CORRER
 -----------
   1) Instalar Python (marcar "Add Python to PATH").
   2) pip install streamlit pandas numpy plotly scipy openpyxl ezdxf
   3) Por este ficheiro e o Excel na mesma pasta.
   4) streamlit run app_eden.py

 SEPARADOR DA PLANTA (DXF)
 -------------------------
 O separador "Planta (DXF)" le um desenho DXF (exportado do AutoCAD/Civil 3D)
 e desenha as suas linhas com os alvos por cima. Para a planta encaixar nos
 alvos, o DXF tem de estar no MESMO referencial de coordenadas da obra
 (M/P ~5000). Se nao estiver, ha um ajuste manual de posicao no separador.
=============================================================================
"""

from pathlib import Path
import base64

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# scipy e usado so para a superficie interpolada; se faltar, a app continua
# a funcionar e apenas desativa essa opcao.
try:
    from scipy.interpolate import griddata
    TEM_SCIPY = True
except Exception:
    TEM_SCIPY = False

# ezdxf le os desenhos DXF; se faltar, o separador da planta avisa e
# a restante app funciona na mesma.
try:
    import ezdxf
    TEM_EZDXF = True
except Exception:
    TEM_EZDXF = False


# =========================================================================
# CONFIGURACAO
# =========================================================================
FICHEIRO_EXCEL = "Modelo_Dados_Instrumentacao_The_Eden_2.xlsx"

FOLHAS = {
    "instrumentos": "Instrumentos",
    "alvos": "Alvos_Topograficos",
    "resumo": "Inclinometros_Resumo",
    "perfis": "Inclinometros_Perfis",
    "celulas": "Celulas_Carga",
    "piezo": "Piezometros",
}

COLS = {
    "data": "Data",
    # inclinometros
    "inclinometro": "Inclinómetro",
    "profundidade": "Profundidade (m)",
    "desl_total": "Desl. acumulado total (mm)",
    "desl_max_global": "Máx. desloc. acumulado total (mm)",
    "prof_do_max": "Profundidade do máximo (m)",
    # alvos topograficos
    "alvo": "Alvo",
    "edificio": "Edifício / elemento",
    "M0": "M0 (m)", "P0": "P0 (m)", "Z0": "Z0 (m)",
    "dM": "ΔM acumulado (mm)", "dP": "ΔP acumulado (mm)", "dZ": "ΔZ acumulado (mm)",
    "desl_h": "Desloc. horizontal acumulado (mm)",
    # celulas
    "celula": "Célula", "ancoragem": "Ancoragem",
    "carga_atual": "Carga atual (kN)", "blocagem": "Blocagem (kN)",
    "variacao": "Variação calculada", "estado": "Estado",
    # piezometros
    "piezometro": "Piezómetro", "cota_agua": "Cota da água calculada (m)",
    "cota_boca": "Cota da boca (m)", "prof_abaixo_boca": "Profundidade abaixo da boca (m)",
}

CC_ALERTA = 0.15
CC_ALARME = 0.25
LIMIAR_VEL_DEFEITO = 0.5
FATOR_ACEL_DEFEITO = 1.8

# =========================================================================
# CRITERIOS DE ALERTA / ALARME DOS ALVOS TOPOGRAFICOS
# -------------------------------------------------------------------------
# Transcritos do relatorio de alvos topograficos (33GRADOS). Sao criterios
# de DESLOCAMENTO ACUMULADO, em mm, aplicados por grupo. O estado de cada
# leitura e o MAIS SEVERO entre o nivel horizontal e o vertical.
#
#   (H_alerta, H_alarme, V_alerta, V_alarme)  em mm
#
# EDIFICIOS ADJACENTES (Santa Casa, Cimas, Clinica): H 15/25, V 10/20.
#   -> CONFIRMADO no relatorio (Cap. 4, "Criterios de Alerta e de Alarme
#      (Edificios Adjacentes)"). Reproduz a coluna Estado do Excel a 100%.
# CONTENCAO 17 m (alcados AB, CD, BF, PQ): H 20/40, V 10/15.
# CONTENCAO 24 m (alcados FG, GH, JK, KL, MNO, OP): H 30/40, V 10/15.
#   -> Estes reproduzem a coluna Estado a 100% nos respetivos alcados.
# ALCADO DE: classificado provisoriamente como 17 m. Os deslocamentos
#   observados sao ~0 (tudo Regular), pelo que os dados NAO permitem
#   distinguir 17 de 24 m. A CONFIRMAR com o projeto de contencao.
# =========================================================================
CRIT_VIZINHOS   = (15, 25, 10, 20)   # edificios adjacentes (OFICIAL)
CRIT_CONT_17    = (20, 40, 10, 15)   # contencao 17 m
CRIT_CONT_24    = (30, 40, 10, 15)   # contencao 24 m

ALCADOS_24M = {"FG", "GH", "JK", "KL", "MNO", "OP"}
ALCADOS_17M = {"AB", "CD", "BF", "PQ"}
ALCADOS_A_CONFIRMAR = {"DE"}         # sem deslocamento -> grupo nao distinguivel


def _extrair_alcado(edif):
    """Devolve o codigo do alcado (ex. 'FG') ou None se nao for contencao."""
    import re
    if isinstance(edif, str) and "Alçado" in edif:
        m = re.search(r"Alçado (\w+)", edif)
        return m.group(1) if m else None
    return None


def criterios_do_alvo(edif):
    """
    Devolve (criterio, rotulo, a_confirmar) para uma linha de alvo, a partir
    do nome do edificio/elemento. 'criterio' e o tuplo (Ha,Hm,Va,Vm).
    """
    alc = _extrair_alcado(edif)
    if alc is None:
        return CRIT_VIZINHOS, "Edificio adjacente (15/25 · 10/20)", False
    if alc in ALCADOS_24M:
        return CRIT_CONT_24, f"Contencao 24 m — Alcado {alc} (30/40 · 10/15)", False
    if alc in ALCADOS_17M:
        return CRIT_CONT_17, f"Contencao 17 m — Alcado {alc} (20/40 · 10/15)", False
    # alcado sem classificacao segura
    return CRIT_CONT_17, f"Alcado {alc} (17 m assumido — A CONFIRMAR)", True


def estado_calculado(h, v, criterio):
    """
    Estado a partir do deslocamento horizontal (h) e vertical (v) acumulados,
    dado um criterio (Ha,Hm,Va,Vm). O estado e o mais severo entre H e V.
    Devolve 'Alarme' | 'Alerta' | 'Regular' | 'Sem leitura'.
    """
    ha, hm, va, vm = criterio
    if pd.isna(h) and pd.isna(v):
        return "Sem leitura"
    nh = 2 if (pd.notna(h) and h >= hm) else (1 if (pd.notna(h) and h >= ha) else 0)
    nv = 2 if (pd.notna(v) and abs(v) >= vm) else (1 if (pd.notna(v) and abs(v) >= va) else 0)
    n = max(nh, nv)
    return "Alarme" if n == 2 else ("Alerta" if n == 1 else "Regular")


def anexar_estado_calculado(df):
    """
    Recebe o dataframe de alvos e devolve uma copia com colunas novas:
      'Criterio'          — rotulo legivel do criterio aplicado
      'Estado calculado'  — estado recalculado de ΔH/ΔV com os criterios oficiais
      'Confere'           — True se coincide com a coluna 'Estado' do Excel
      'Fachada SC'        — 'Frente escavacao' | 'Lateral (mar)' | '' (so Santa Casa)
    Nao altera a coluna 'Estado' original: serve de auditoria lado a lado.
    """
    d = df.copy()
    crits, rotulos, estados, confere, fachadas = [], [], [], [], []
    for _, r in d.iterrows():
        crit, rotulo, _ac = criterios_do_alvo(r.get(COLS["edificio"]))
        h = r.get(COLS["desl_h"]); v = r.get(COLS["dZ"])
        ec = estado_calculado(h, v, crit)
        crits.append(crit); rotulos.append(rotulo); estados.append(ec)
        est_excel = r.get(COLS["estado"])
        # so compara quando ambos tem um estado 'real'
        if isinstance(est_excel, str) and est_excel in ("Regular", "Alerta", "Alarme") \
           and ec in ("Regular", "Alerta", "Alarme"):
            confere.append(ec == est_excel)
        else:
            confere.append(None)
        fachadas.append(fachada_santa_casa(r.get(COLS["edificio"]), r.get(COLS["alvo"])))
    d["Criterio"] = rotulos
    d["Estado calculado"] = estados
    d["Confere"] = confere
    d["Fachada SC"] = fachadas
    return d


# =========================================================================
# SANTA CASA — DUAS FACHADAS E SUBSTITUICAO DE ALVOS
# -------------------------------------------------------------------------
# O edificio da Santa Casa da Misericordia tem duas fachadas instrumentadas:
#   Fachada 1 (frente a escavacao): alvos A1, A2, A3, A4
#   Fachada 2 (lateral, virada ao mar): A5/A5b, A6/A6b, A7/A7b, A8/A8b
# Os alvos A5-A8 foram tapados por um painel publicitario (out/2025) e
# substituidos por A5b-A8b, RE-ZERADOS na data da troca (20/10/2025). Por
# isso os "b" arrancam de zero mais tarde: os seus acumulados NAO sao
# comparaveis diretamente com A1-A4. (Fonte: folha Qualidade_Dados do Excel
# e planta de localizacao do relatorio.)
# =========================================================================
SC_FACHADA_1 = {"A1", "A2", "A3", "A4"}
SC_FACHADA_2 = {"A5", "A6", "A7", "A8", "A5b", "A6b", "A7b", "A8b"}
SC_SUBSTITUIDOS = {"A5": "A5b", "A6": "A6b", "A7": "A7b", "A8": "A8b"}


def fachada_santa_casa(edif, alvo):
    """Devolve a fachada da Santa Casa a que o alvo pertence, ou '' se nao aplicar."""
    if not (isinstance(edif, str) and "Santa Casa" in edif):
        return ""
    a = str(alvo)
    if a in SC_FACHADA_1:
        return "Frente escavacao"
    if a in SC_FACHADA_2:
        return "Lateral (mar)"
    return ""

# =========================================================================
# DADOS GEOLOGICOS  (Relatorio Geologico-Geotecnico ENGGEO, processo 220216)
# Transcritos dos Quadros II, III, V, VI, VII e dos logs de sondagem.
# Ficam embutidos porque vêm do relatorio e nao mudam. Nenhum valor inventado.
# =========================================================================
GEO_SONDAGENS = {
    "SC6/Pz": {"cota_terreno": 23.3, "profundidade": 21.05, "nf_prof": 7.17},
    "SC7":    {"cota_terreno": 26.0, "profundidade": 19.89, "nf_prof": None},
    "SC8/Pz": {"cota_terreno": 28.8, "profundidade": 21.00, "nf_prof": 9.47},
    "SC9/Pz": {"cota_terreno": 27.5, "profundidade": 21.41, "nf_prof": 11.18},
}

# camadas (topo, base, unidade) em profundidade (m)
GEO_LITOLOGIA = {
    "SC6/Pz": [(0.0, 0.5, "Aterro"), (0.5, 21.05, "Gres (C1As)")],
    "SC7":    [(0.0, 0.5, "Aterro"), (0.5, 19.89, "Gres (C1As)")],
    "SC8/Pz": [(0.0, 0.5, "Aterro"), (0.5, 19.5, "Gres (C1As)"),
               (19.5, 21.0, "Calcario (C1A)")],
    "SC9/Pz": [(0.0, 0.5, "Aterro"), (0.5, 21.41, "Gres (C1As)")],
}

# cor de cada unidade litologica (para a coluna)
GEO_CORES_LITO = {
    "Aterro": "#d9822b",
    "Gres (C1As)": "#a9c47f",
    "Calcario (C1A)": "#6b8e4e",
}

# ensaios SPT: (profundidade_m, N). N=60 indica nega.
GEO_SPT = {
    "SC6/Pz": [(1.5,60),(3.0,60),(4.5,46),(6.0,60),(7.5,60),(9.0,38),(10.5,25),
               (12.0,56),(13.5,35),(15.0,25),(16.5,22),(18.0,60),(19.5,60)],
    "SC7":    [(1.5,32),(3.0,36),(4.5,60),(6.0,60),(7.5,60),(9.0,60),(10.5,60),
               (12.0,60),(13.5,60),(15.0,60),(16.5,60),(18.0,60)],
    "SC8/Pz": [(1.5,11),(3.0,52),(4.5,60),(6.0,19),(7.5,24),(9.0,36),(10.5,32),
               (12.0,47),(13.5,60),(15.0,60),(16.5,60),(18.0,22),(19.5,60)],
    "SC9/Pz": [(1.5,60),(3.0,60),(4.5,60),(6.0,60),(7.5,60),(9.0,60),(10.5,40),
               (12.0,60),(13.5,60),(15.0,35),(16.5,37),(18.0,60),(19.5,60)],
}

# zonamento geotecnico (Quadros V, VI, VII)
GEO_ZONAMENTO = [
    {"Zona":"ZG6","Descricao":"Aterro heterogeneo de origem nao selectiva",
     "gama (kN/m3)":"17-18","c' (kPa)":"<5","fi' (graus)":"<26","E'":"<5 MPa"},
    {"Zona":"ZG5","Descricao":"Gres pouco consolidado, SPT 11-30, RQD 0%",
     "gama (kN/m3)":"19-20","c' (kPa)":"5-15","fi' (graus)":"28-33","E'":"8-30 MPa"},
    {"Zona":"ZG4","Descricao":"Gres pouco consolidado, SPT 31-56, RQD 0%",
     "gama (kN/m3)":"20-21","c' (kPa)":"5-30","fi' (graus)":"30-36","E'":"25-50 MPa"},
    {"Zona":"ZG3","Descricao":"Gres/calcario irreg. consolidado, SPT>=60, RQD 0-25%",
     "gama (kN/m3)":"22-24","c' (kPa)":"0.04-0.40 MPa","fi' (graus)":"29-32","E'":"0.07-0.22 GPa"},
    {"Zona":"ZG2","Descricao":"Gres irreg. consolidado a consolidado, SPT>=60, RQD 45-75%",
     "gama (kN/m3)":"24-25","c' (kPa)":"1.26-3.14 MPa","fi' (graus)":"32-36","E'":"0.80-2.81 GPa"},
    {"Zona":"ZG1","Descricao":"Gres consolidado, SPT>=60, RQD 76-100%",
     "gama (kN/m3)":"25-26","c' (kPa)":"2.70-8.69 MPa","fi' (graus)":"38-42","E'":"7.94-14.13 GPa"},
]

# =========================================================================
# CRONOGRAMA DA OBRA  (Plano de Trabalhos Alves Ribeiro/HCI, 05/05/2025)
# Datas PREVISTAS transcritas do PDF do plano. Sao o planeado, nao o real.
# =========================================================================
# (nome, inicio ISO, conclusao ISO)
FASES_OBRA = [
    ("Contencao periferica",                 "2025-05-13", "2026-02-03"),
    ("Estacas Poente e Norte",               "2025-05-13", "2025-06-23"),
    ("Estacas Central e Nascente",           "2025-06-24", "2025-08-18"),
    ("Escavacao + bandas de laje + ancoragens", "2025-07-08", "2026-03-02"),
    ("Fundacao e laje de fundo",             "2026-01-20", "2026-03-16"),
    ("Microestacas",                         "2026-02-03", "2026-03-02"),
    ("Piso -3", "2026-02-24", "2026-04-06"),
    ("Piso -2", "2026-03-17", "2026-04-27"),
    ("Piso -1", "2026-03-31", "2026-05-11"),
]

CORES_FASES = ["#8dd3c7", "#ffffb3", "#bebada", "#fb8072", "#80b1d3",
               "#fdb462", "#b3de69", "#fccde5", "#d9d9d9"]


def adicionar_fases_obra(fig, dt_min, dt_max, faixas=True, marcos=True):
    """
    Sobrepoe as fases da obra a um grafico com o tempo no eixo X.
    So desenha as fases que se sobrepoem a janela [dt_min, dt_max] dos dados,
    para nao encher o grafico com fases de 2026 quando os dados sao de 2025.
    """
    dt_min = pd.to_datetime(dt_min)
    dt_max = pd.to_datetime(dt_max)
    # margem para as fases que comecam pouco antes/depois
    margem = pd.Timedelta(days=20)

    for i, (nome, ini, fim) in enumerate(FASES_OBRA):
        t0 = pd.to_datetime(ini)
        t1 = pd.to_datetime(fim)
        # sobrepoe a janela dos dados?
        if t1 < dt_min - margem or t0 > dt_max + margem:
            continue
        cor = CORES_FASES[i % len(CORES_FASES)]
        # recortar a faixa a janela visivel
        vt0 = max(t0, dt_min - margem)
        vt1 = min(t1, dt_max + margem)
        if faixas:
            fig.add_vrect(x0=vt0, x1=vt1, fillcolor=cor, opacity=0.18,
                          line_width=0, layer="below",
                          annotation_text=nome, annotation_position="top left",
                          annotation=dict(font_size=9, textangle=0))
        if marcos:
            # linha no inicio da fase, se cair na janela
            if dt_min - margem <= t0 <= dt_max + margem:
                fig.add_vline(x=t0, line=dict(color=cor, width=1.5, dash="dot"))


st.set_page_config(page_title="IMS — Instrumentation Monitoring System",
                   layout="wide")


# =========================================================================
# CARREGAMENTO
# =========================================================================
@st.cache_data(show_spinner="A carregar o Excel...")
def carregar_dados(fonte):
    excel = pd.ExcelFile(fonte)
    dados = {}
    for chave, folha in FOLHAS.items():
        try:
            df = pd.read_excel(excel, folha)
        except Exception:
            df = pd.DataFrame()
        if COLS["data"] in df.columns:
            df[COLS["data"]] = pd.to_datetime(df[COLS["data"]], errors="coerce")
        dados[chave] = df
    return dados


def validar_colunas(df, nomes, contexto):
    if df.empty:
        st.warning(f"A folha de '{contexto}' esta vazia ou nao foi encontrada.")
        return False
    faltam = [n for n in nomes if n not in df.columns]
    if faltam:
        st.error(f"Em '{contexto}' faltam colunas: {faltam}. "
                 f"Verifica o bloco CONFIGURACAO (COLS) ou o Excel.")
        return False
    return True


# =========================================================================
# CALCULO DE VELOCIDADE / PRECURSORES  (inclinometros)
# =========================================================================
def calcular_velocidade(datas, valores, limiar_vel, fator_acel):
    df = pd.DataFrame({"data": pd.to_datetime(list(datas)), "valor": list(valores)})
    df = df.sort_values("data").reset_index(drop=True)
    df["dias"] = df["data"].diff().dt.days
    df["delta"] = df["valor"].diff()
    df["velocidade"] = np.where(df["dias"] > 0,
                                (df["delta"] / df["dias"]).round(3), np.nan)
    df["precursor"] = False
    for i in range(1, len(df)):
        v, va = df.loc[i, "velocidade"], df.loc[i - 1, "velocidade"]
        acima = pd.notna(v) and v >= limiar_vel
        acel = pd.notna(v) and pd.notna(va) and va > 0 and v >= fator_acel * va
        df.loc[i, "precursor"] = bool(acima or acel)
    return df


# =========================================================================
# LEITURA DE DESENHOS DXF
# =========================================================================
@st.cache_data(show_spinner="A ler o desenho DXF...")
def dxf_listar_layers(conteudo_bytes):
    """Devolve a lista de nomes de layers de um DXF (recebido como bytes)."""
    import io
    from ezdxf.recover import read as recover_read
    doc, _ = recover_read(io.BytesIO(conteudo_bytes))
    return [l.dxf.name for l in doc.layers]


@st.cache_data(show_spinner="A extrair geometria do DXF...")
def dxf_extrair_segmentos(conteudo_bytes, layers_incluir=None):
    """
    Le um DXF e devolve uma lista de polilinhas para desenhar:
        [(xs, ys, layer), ...]
    Le os tipos mais comuns em plantas: LINE, LWPOLYLINE, POLYLINE, ARC, CIRCLE.
    (ARC e CIRCLE sao essenciais: as estacas da cortina de contencao aparecem
    desenhadas como pequenos arcos/circulos.)
    Se 'layers_incluir' for dado, so devolve entidades dessas layers.
    """
    import io
    import numpy as np
    from ezdxf.recover import read as recover_read
    doc, _ = recover_read(io.BytesIO(conteudo_bytes))
    msp = doc.modelspace()

    segmentos = []
    for e in msp:
        t = e.dxftype()
        lay = e.dxf.layer
        if layers_incluir and lay not in layers_incluir:
            continue
        try:
            if t == "LINE":
                segmentos.append(([e.dxf.start[0], e.dxf.end[0]],
                                  [e.dxf.start[1], e.dxf.end[1]], lay))
            elif t == "LWPOLYLINE":
                pts = e.get_points()
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                if e.closed and len(xs) > 2:
                    xs = xs + [xs[0]]
                    ys = ys + [ys[0]]
                segmentos.append((xs, ys, lay))
            elif t == "POLYLINE":
                xs = [v.dxf.location[0] for v in e.vertices]
                ys = [v.dxf.location[1] for v in e.vertices]
                if xs:
                    segmentos.append((xs, ys, lay))
            elif t == "ARC":
                a0 = np.radians(e.dxf.start_angle)
                a1 = np.radians(e.dxf.end_angle)
                if a1 < a0:
                    a1 += 2 * np.pi
                ang = np.linspace(a0, a1, 16)
                cx, cy, r = e.dxf.center[0], e.dxf.center[1], e.dxf.radius
                segmentos.append((list(cx + r * np.cos(ang)),
                                  list(cy + r * np.sin(ang)), lay))
            elif t == "CIRCLE":
                ang = np.linspace(0, 2 * np.pi, 20)
                cx, cy, r = e.dxf.center[0], e.dxf.center[1], e.dxf.radius
                segmentos.append((list(cx + r * np.cos(ang)),
                                  list(cy + r * np.sin(ang)), lay))
        except Exception:
            pass
    return segmentos


def dxf_gama_coordenadas(segmentos):
    """Devolve (xmin, xmax, ymin, ymax) do conjunto de segmentos, ou None."""
    xs, ys = [], []
    for sx, sy, _ in segmentos:
        xs.extend(sx)
        ys.extend(sy)
    if not xs:
        return None
    return min(xs), max(xs), min(ys), max(ys)


# =========================================================================
# SEPARADOR 1 — VISAO GERAL 3D (ALVOS TOPOGRAFICOS)
# =========================================================================
def classificar_grupo(nome_edificio):
    """Devolve ('contencao', alcado) ou ('edificio', nome) para dar cor/forma."""
    import re
    if isinstance(nome_edificio, str) and "Alçado" in nome_edificio:
        m = re.search(r"Alçado (\w+)", nome_edificio)
        return ("contencao", m.group(1) if m else "?")
    return ("edificio", nome_edificio)


def contorno_recinto(campanha):
    """
    Constroi o contorno do recinto como a ENVOLVENTE CONVEXA (convex hull)
    dos alvos da contencao periferica. Ao contrario de ligar centroides, o
    convex hull envolve sempre os pontos por fora — nunca passa por dentro
    da nuvem, evitando a falsa impressao de alvos 'interiores'.
    Devolve (Ms, Ps, Zs) ja fechado, ou None se nao houver pontos/scipy.
    Tudo no sistema dos alvos — nao precisa de DXF nem de alinhamento.
    """
    import numpy as np
    cont = campanha[campanha[COLS["edificio"]].astype(str).str.contains("Alçado", na=False)].copy()
    if len(cont) < 3:
        return None
    pts = cont[[COLS["M0"], COLS["P0"]]].to_numpy()
    z_med = float(cont[COLS["Z0"]].mean())
    try:
        from scipy.spatial import ConvexHull
        hull = ConvexHull(pts)
        poly = pts[hull.vertices]
    except Exception:
        # sem scipy: cai para ordenacao angular (menos bom, mas funcional)
        cx, cy = pts[:, 0].mean(), pts[:, 1].mean()
        ang = np.arctan2(pts[:, 1] - cy, pts[:, 0] - cx)
        poly = pts[np.argsort(ang)]
    Ms = list(poly[:, 0]) + [poly[0, 0]]
    Ps = list(poly[:, 1]) + [poly[0, 1]]
    Zs = [z_med] * len(Ms)
    return Ms, Ps, Zs


# cores fixas por edificio vizinho (as restantes recebem cor automatica)
CORES_EDIFICIO = {
    "Edifício Santa Casa da Misericórdia": "#d62728",
    "Restaurante Cimas": "#2ca02c",
    "Clínica Abreu Loureiro": "#1f77b4",
}


def separador_3d(dados):
    alvos = dados["alvos"]
    ok = validar_colunas(
        alvos,
        [COLS["data"], COLS["alvo"], COLS["edificio"], COLS["M0"], COLS["P0"],
         COLS["Z0"], COLS["dM"], COLS["dP"], COLS["dZ"], COLS["desl_h"]],
        "Alvos topograficos",
    )
    if not ok:
        return

    st.subheader("Movimento dos alvos no espaco, com a geometria da obra")
    st.caption("O contorno castanho e a envolvente dos alvos da contencao "
               "(convex hull), que aproxima o limite do recinto de escavacao. "
               "Os alvos dos edificios vizinhos aparecem agrupados e "
               "identificados por cor. As setas mostram a direcao e magnitude "
               "do deslocamento acumulado (amplificado). Todos os alvos sao de "
               "periferia — na cortina de contencao ou nas fachadas vizinhas; "
               "nao ha instrumentos dentro da escavacao. Tudo no sistema de "
               "coordenadas dos alvos, sem necessidade de DXF.")

    datas = sorted(alvos[COLS["data"]].dropna().unique())
    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        data_sel = st.select_slider(
            "Campanha", options=datas, value=datas[-1],
            format_func=lambda d: pd.to_datetime(d).strftime("%d/%m/%Y"))
    with col_b:
        fator = st.slider("Amplificacao do deslocamento", 50, 2000, 500, 50,
                          help="Os deslocamentos sao milimetricos e as "
                               "coordenadas em metros; amplifica-se para ver.")
    with col_c:
        mostrar_contorno = st.checkbox("Contorno do recinto", value=True)
        mostrar_paredes = st.checkbox("Paredes dos alcados", value=False,
                                      help="Desenha planos verticais nos "
                                           "alcados da contencao, para dar "
                                           "volume ao recinto.")
        destacar_alarmes = st.checkbox(
            "Destacar alarmes/alertas", value=True,
            help="Contorna a vermelho os alvos em alarme e a laranja os "
                 "em alerta, segundo os criterios oficiais recalculados.")

    campanha = alvos[alvos[COLS["data"]] == data_sel].copy()
    # recalcular estado de cada alvo da campanha com os criterios oficiais
    campanha = anexar_estado_calculado(campanha)

    fig = go.Figure()

    # ---- contorno do recinto (a partir dos alcados) ----------------------
    cont = contorno_recinto(campanha)
    if mostrar_contorno and cont is not None:
        Ms, Ps, Zs = cont
        fig.add_trace(go.Scatter3d(
            x=Ms, y=Ps, z=Zs, mode="lines",
            line=dict(color="saddlebrown", width=6),
            name="Contorno do recinto (envolvente)", hoverinfo="skip",
        ))
        # paredes verticais opcionais (da cota do contorno para baixo)
        if mostrar_paredes:
            base_z = min(Zs) - 8  # profundidade visual da escavacao
            for i in range(len(Ms) - 1):
                fig.add_trace(go.Scatter3d(
                    x=[Ms[i], Ms[i+1], Ms[i+1], Ms[i], Ms[i]],
                    y=[Ps[i], Ps[i+1], Ps[i+1], Ps[i], Ps[i]],
                    z=[Zs[i], Zs[i+1], base_z, base_z, Zs[i]],
                    mode="lines", line=dict(color="peru", width=2),
                    surfaceaxis=2, surfacecolor="rgba(210,180,140,0.25)",
                    showlegend=False, hoverinfo="skip",
                ))

    # ---- alvos por grupo (cor por edificio; contencao a laranja) ---------
    grupos = campanha.groupby(campanha[COLS["edificio"]].apply(
        lambda s: classificar_grupo(s)[0] if classificar_grupo(s)[0] == "edificio"
        else "Contencao periferica"))

    # primeiro a contencao (laranja), depois cada edificio com a sua cor
    for chave, grp in campanha.groupby(COLS["edificio"]):
        tipo, etiqueta = classificar_grupo(chave)
        x0 = grp[COLS["M0"]].to_numpy()
        y0 = grp[COLS["P0"]].to_numpy()
        z0 = grp[COLS["Z0"]].to_numpy()
        dx = grp[COLS["dM"]].to_numpy() / 1000.0 * fator
        dy = grp[COLS["dP"]].to_numpy() / 1000.0 * fator
        dz = grp[COLS["dZ"]].to_numpy() / 1000.0 * fator
        dh = grp[COLS["desl_h"]].to_numpy()
        nomes = grp[COLS["alvo"]].astype(str).to_numpy()

        estados = grp["Estado calculado"].to_numpy()
        fachadas = grp["Fachada SC"].to_numpy()

        if tipo == "edificio":
            cor = CORES_EDIFICIO.get(chave, None)
            nome_leg = chave
        else:
            cor = "#ff7f0e"
            nome_leg = f"Contencao — Alcado {etiqueta}"

        # contorno do marcador por estado (destaque de alarme/alerta).
        # NOTA: em Scatter3d, marker.line.width tem de ser um escalar (nao
        # aceita lista por ponto, ao contrario do Scatter 2D). A cor da borda
        # PODE ser uma lista por ponto — e o que usamos para o destaque.
        if destacar_alarmes:
            cor_borda = ["#c0140f" if e == "Alarme" else
                         ("#e67e00" if e == "Alerta" else "rgba(0,0,0,0.2)")
                         for e in estados]
            # largura unica: mais grossa se o grupo tiver algum alarme/alerta
            larg_borda = 4 if any(e in ("Alarme", "Alerta") for e in estados) else 1
        else:
            cor_borda = "rgba(0,0,0,0.2)"
            larg_borda = 1

        marker = dict(size=6, color=cor if cor else "#7f7f7f",
                      line=dict(color=cor_borda, width=larg_borda))

        # hover com estado e fachada
        cd = np.column_stack([dh, estados, fachadas])
        fig.add_trace(go.Scatter3d(
            x=x0 + dx, y=y0 + dy, z=z0 + dz, mode="markers+text",
            marker=marker,
            text=nomes, textposition="top center", textfont=dict(size=8),
            name=nome_leg,
            customdata=cd,
            hovertemplate="Alvo %{text}<br>Desl. h: %{customdata[0]:.1f} mm"
                          "<br>Estado: %{customdata[1]}"
                          "<br>%{customdata[2]}"
                          "<extra>" + nome_leg + "</extra>",
        ))
        # setas de deslocamento
        for i in range(len(x0)):
            fig.add_trace(go.Scatter3d(
                x=[x0[i], x0[i] + dx[i]], y=[y0[i], y0[i] + dy[i]],
                z=[z0[i], z0[i] + dz[i]],
                mode="lines", line=dict(color="crimson", width=3),
                showlegend=False, hoverinfo="skip",
            ))

    fig.update_layout(
        height=720,
        scene=dict(xaxis_title="M (m)", yaxis_title="P (m)", zaxis_title="Z (m)",
                   aspectmode="data"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(size=9)),
        margin=dict(l=0, r=0, t=30, b=0),
    )
    st.plotly_chart(fig, use_container_width=True)

    # metricas
    desl_h_all = campanha[COLS["desl_h"]].to_numpy()
    nomes_all = campanha[COLS["alvo"]].astype(str).to_numpy()
    edif_all = campanha[COLS["edificio"]].to_numpy()
    n_alarme = int((campanha["Estado calculado"] == "Alarme").sum())
    n_alerta = int((campanha["Estado calculado"] == "Alerta").sum())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alvos na campanha", len(campanha))
    c2.metric("Desl. horizontal max. (mm)", f"{np.nanmax(desl_h_all):.1f}")
    c3.metric("Em alarme", n_alarme)
    c4.metric("Em alerta", n_alerta)
    idx = int(np.nanargmax(desl_h_all))
    st.caption(f"O alvo mais afetado ({nomes_all[idx]}, "
               f"{np.nanmax(desl_h_all):.1f} mm) pertence a: {edif_all[idx]}. "
               f"Contorno vermelho = alarme; laranja = alerta (criterios "
               f"oficiais recalculados de ΔH/ΔV).")


# =========================================================================
# SEPARADOR 2 — INCLINOMETROS
# =========================================================================
def separador_inclinometros(dados, limiar_vel, fator_acel):
    perfis, resumo = dados["perfis"], dados["resumo"]
    ok = validar_colunas(perfis, [COLS["data"], COLS["inclinometro"],
                                  COLS["profundidade"], COLS["desl_total"]],
                         "Inclinometros / perfis") \
        and validar_colunas(resumo, [COLS["data"], COLS["inclinometro"],
                                     COLS["desl_max_global"], COLS["prof_do_max"]],
                            "Inclinometros / resumo")
    if not ok:
        return

    inc = st.selectbox("Inclinometro",
                       sorted(perfis[COLS["inclinometro"]].dropna().unique()))
    p_inc = perfis[perfis[COLS["inclinometro"]] == inc].copy()
    r_inc = resumo[resumo[COLS["inclinometro"]] == inc].copy().sort_values(COLS["data"])
    datas_inc = sorted(p_inc[COLS["data"]].dropna().unique())
    if not datas_inc:
        st.warning("Sem leituras com data valida.")
        return

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Perfil deformado")
        st.caption("Deslocamento acumulado ao longo da profundidade. Base fixa.")
        idx = sorted(set([0, len(datas_inc) // 2, len(datas_inc) - 1]))
        sel = st.multiselect("Leituras", datas_inc,
                             default=[datas_inc[i] for i in idx],
                             format_func=lambda d: pd.to_datetime(d).strftime("%d/%m/%Y"))

        # opcao de sobrepor a geologia de uma sondagem (peca-chave do back-analysis)
        geo_on = st.checkbox("Sobrepor geologia (sondagem)", value=False,
                             help="Mostra a coluna litologica e o nivel "
                                  "freatico de uma sondagem ao lado do perfil, "
                                  "para relacionar a deformacao com o terreno.")
        sond_sel = None
        if geo_on:
            sond_sel = st.selectbox("Sondagem de referencia",
                                    list(GEO_LITOLOGIA.keys()))

        fig = go.Figure()

        # se geologia ligada, desenhar faixas litologicas de fundo (a toda a largura)
        if geo_on and sond_sel:
            # usar profundidade do perfil para a extensao horizontal das faixas
            xmax = float(p_inc[COLS["desl_total"]].abs().max()) * 1.1 + 1
            for topo, base, unidade in GEO_LITOLOGIA[sond_sel]:
                cor = GEO_CORES_LITO.get(unidade, "#cccccc")
                fig.add_shape(type="rect", x0=-xmax, x1=xmax, y0=topo, y1=base,
                              fillcolor=cor, opacity=0.25,
                              line=dict(width=0), layer="below")
            # nivel freatico
            nf = GEO_SONDAGENS[sond_sel]["nf_prof"]
            if nf is not None:
                fig.add_hline(y=nf, line=dict(color="blue", width=2, dash="dash"),
                              annotation_text=f"NF ({sond_sel})",
                              annotation_position="right")
            # entradas de legenda para as unidades
            for unidade, cor in GEO_CORES_LITO.items():
                if any(u == unidade for _, _, u in GEO_LITOLOGIA[sond_sel]):
                    fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
                                             marker=dict(size=10, color=cor, symbol="square"),
                                             name=unidade))

        for d in sel:
            s = p_inc[p_inc[COLS["data"]] == d].sort_values(COLS["profundidade"])
            fig.add_trace(go.Scatter(x=s[COLS["desl_total"]], y=s[COLS["profundidade"]],
                                     mode="lines+markers",
                                     name=pd.to_datetime(d).strftime("%d/%m/%Y")))
        fig.update_yaxes(autorange="reversed", title="Profundidade (m)")
        fig.update_xaxes(title="Deslocamento acumulado (mm)")
        fig.update_layout(height=560, legend_title="Leitura / geologia")
        st.plotly_chart(fig, use_container_width=True)
        if geo_on and sond_sel:
            st.caption(f"Geologia da sondagem {sond_sel} sobreposta. Repara se "
                       f"a maior curvatura do perfil coincide com uma mudanca "
                       f"de camada ou com o nivel freatico — e a leitura central "
                       f"do back-analysis.")

    with col2:
        st.subheader("Evolucao do deslocamento")
        st.caption("Maximo global vs. profundidade fixa.")
        profs = sorted(p_inc[COLS["profundidade"]].dropna().unique())
        moda = r_inc[COLS["prof_do_max"]].mode()
        pdef = float(moda.iloc[0]) if len(moda) else profs[0]
        if pdef not in profs:
            pdef = profs[0]
        prof_fixa = st.select_slider("Profundidade fixa (m)", options=profs, value=pdef)
        s_fix = p_inc[p_inc[COLS["profundidade"]] == prof_fixa].sort_values(COLS["data"])
        s_max = r_inc.sort_values(COLS["data"])
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=s_max[COLS["data"]], y=s_max[COLS["desl_max_global"]],
                                  mode="lines+markers", name="Maximo global"))
        fig2.add_trace(go.Scatter(x=s_fix[COLS["data"]], y=s_fix[COLS["desl_total"]],
                                  mode="lines+markers", name=f"A {prof_fixa:.1f} m"))
        if st.session_state.get("mostrar_obra") and len(s_max):
            adicionar_fases_obra(fig2, s_max[COLS["data"]].min(),
                                 s_max[COLS["data"]].max())
        fig2.update_xaxes(title="Data")
        fig2.update_yaxes(title="Deslocamento (mm)")
        fig2.update_layout(height=560, legend_title="Serie")
        st.plotly_chart(fig2, use_container_width=True)
        if st.session_state.get("mostrar_obra"):
            st.caption("Faixas coloridas = fases do plano de trabalhos "
                       "(previstas). Repara se a aceleracao do deslocamento "
                       "coincide com o avanco de uma fase de escavacao.")

    st.divider()
    st.subheader("Velocidade e sinais precursores")
    base = st.radio("Serie de base", ["Maximo global", f"Profundidade fixa ({prof_fixa:.1f} m)"],
                    horizontal=True)
    if base == "Maximo global":
        vdf = calcular_velocidade(s_max[COLS["data"]], s_max[COLS["desl_max_global"]],
                                  limiar_vel, fator_acel)
    else:
        vdf = calcular_velocidade(s_fix[COLS["data"]], s_fix[COLS["desl_total"]],
                                  limiar_vel, fator_acel)
    fig3 = go.Figure(go.Bar(
        x=vdf["data"], y=vdf["velocidade"],
        marker_color=["crimson" if p else "steelblue" for p in vdf["precursor"]]))
    fig3.add_hline(y=limiar_vel, line_dash="dash", line_color="crimson",
                   annotation_text="Limiar")
    fig3.update_xaxes(title="Data")
    fig3.update_yaxes(title="Velocidade (mm/dia)")
    fig3.update_layout(height=380)
    st.plotly_chart(fig3, use_container_width=True)

    n = int(vdf["precursor"].sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("Ult. desloc. max. (mm)",
              f"{s_max[COLS['desl_max_global']].iloc[-1]:.2f}" if len(s_max) else "-")
    vmax = vdf["velocidade"].max()
    c2.metric("Velocidade max. (mm/dia)", f"{vmax:.3f}" if pd.notna(vmax) else "-")
    c3.metric("Precursores", n)
    (st.warning if n else st.success)(
        f"{n} leitura(s) com aceleracao acima dos criterios." if n
        else "Nenhuma aceleracao acima dos criterios.")


# =========================================================================
# SEPARADOR 3 — ALVOS (2D, series temporais)
# =========================================================================
def separador_alvos_2d(dados):
    alvos = dados["alvos"]
    if not validar_colunas(alvos, [COLS["data"], COLS["alvo"], COLS["edificio"],
                                   COLS["desl_h"], COLS["dZ"]], "Alvos topograficos"):
        return
    st.subheader("Alvos topograficos — evolucao temporal")

    # recalcular estado de ΔH/ΔV com os criterios oficiais (auditoria)
    alvos = anexar_estado_calculado(alvos)

    # ---- painel de estado da ultima campanha ----------------------------
    ult = alvos[alvos[COLS["data"]] == alvos[COLS["data"]].max()].copy()
    data_ult = pd.to_datetime(alvos[COLS["data"]].max()).strftime("%d/%m/%Y")
    n_alarme = int((ult["Estado calculado"] == "Alarme").sum())
    n_alerta = int((ult["Estado calculado"] == "Alerta").sum())
    n_reg = int((ult["Estado calculado"] == "Regular").sum())
    n_sl = int((ult["Estado calculado"] == "Sem leitura").sum())

    st.markdown(f"**Estado na ultima campanha ({data_ult})** — recalculado dos "
                f"deslocamentos ΔH/ΔV com os criterios oficiais:")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Em ALARME", n_alarme)
    m2.metric("Em ALERTA", n_alerta)
    m3.metric("Regular", n_reg)
    m4.metric("Sem leitura", n_sl)

    # auditoria: o recalculo confere com a coluna Estado do Excel?
    comp = alvos[alvos["Confere"].notna()]
    n_conf = int(comp["Confere"].sum())
    n_tot = int(len(comp))
    if n_tot:
        if n_conf == n_tot:
            st.success(f"Auditoria: o estado recalculado coincide com a coluna "
                       f"'Estado' do relatorio em {n_conf}/{n_tot} leituras (100%). "
                       f"Os criterios estao corretamente reproduzidos.")
        else:
            st.warning(f"Auditoria: divergencia em {n_tot - n_conf}/{n_tot} leituras "
                       f"entre o estado recalculado e a coluna 'Estado' do relatorio. "
                       f"Ver tabela de divergencias abaixo.")
            with st.expander("Ver divergencias estado recalculado vs. relatorio"):
                div = comp[~comp["Confere"]]
                st.dataframe(
                    div[[COLS["data"], COLS["alvo"], COLS["edificio"],
                         COLS["desl_h"], COLS["dZ"], COLS["estado"],
                         "Estado calculado", "Criterio"]],
                    use_container_width=True, hide_index=True)

    # alvos em alarme/alerta agora, para leitura rapida
    crit_now = ult[ult["Estado calculado"].isin(["Alarme", "Alerta"])].copy()
    if len(crit_now):
        crit_now = crit_now.sort_values("Estado calculado")
        st.caption("Alvos em alerta ou alarme na ultima campanha:")
        st.dataframe(
            crit_now[[COLS["alvo"], COLS["edificio"], "Fachada SC",
                      COLS["desl_h"], COLS["dZ"], "Estado calculado", "Criterio"]]
            .rename(columns={COLS["desl_h"]: "Desl. H (mm)", COLS["dZ"]: "ΔZ (mm)"}),
            use_container_width=True, hide_index=True)
    st.divider()

    edificios = sorted(alvos[COLS["edificio"]].dropna().unique())
    edi = st.selectbox("Edificio / elemento", edificios)
    sub = alvos[alvos[COLS["edificio"]] == edi]

    # se for a Santa Casa, permitir filtrar por fachada e avisar da substituicao
    e_santa_casa = isinstance(edi, str) and "Santa Casa" in edi
    if e_santa_casa:
        st.info(
            "A Santa Casa tem duas fachadas instrumentadas: **Frente a escavacao** "
            "(A1–A4) e **Lateral, virada ao mar** (A5–A8). Os alvos A5–A8 foram "
            "tapados por um painel publicitario e substituidos por **A5b–A8b**, "
            "RE-ZERADOS em 20/10/2025 — por isso os acumulados dos 'b' nao sao "
            "comparaveis diretamente com A1–A4 (arrancam de zero mais tarde).")
        fach = st.radio("Fachada", ["Ambas", "Frente escavacao", "Lateral (mar)"],
                        horizontal=True)
        if fach != "Ambas":
            sub = sub[sub["Fachada SC"] == fach]

    lista = sorted(sub[COLS["alvo"]].dropna().unique())
    sel = st.multiselect("Alvos", lista, default=lista[:min(5, len(lista))])

    # criterio aplicavel a este edificio (para desenhar as linhas de limiar)
    crit_edi, rotulo_edi, ac_edi = criterios_do_alvo(edi)
    Ha, Hm, Va, Vm = crit_edi
    st.caption(f"Criterio aplicado: **{rotulo_edi}**. As linhas tracejadas nos "
               f"graficos marcam os limiares de alerta e alarme.")
    if ac_edi:
        st.warning("Este alcado esta com criterio ASSUMIDO (a confirmar com o "
                   "projeto de contencao).")
    data_rezerag = pd.to_datetime("2025-10-20")

    col1, col2 = st.columns(2)
    with col1:
        st.caption("Deslocamento horizontal acumulado (mm)")
        fig = go.Figure()
        for a in sel:
            s = sub[sub[COLS["alvo"]] == a].sort_values(COLS["data"])
            fig.add_trace(go.Scatter(x=s[COLS["data"]], y=s[COLS["desl_h"]],
                                     mode="lines+markers", name=a))
        if st.session_state.get("mostrar_obra") and len(sub):
            adicionar_fases_obra(fig, sub[COLS["data"]].min(), sub[COLS["data"]].max())
        fig.add_hline(y=Ha, line_dash="dash", line_color="orange",
                      annotation_text=f"Alerta {Ha}", annotation_position="right")
        fig.add_hline(y=Hm, line_dash="dash", line_color="red",
                      annotation_text=f"Alarme {Hm}", annotation_position="right")
        if e_santa_casa:
            fig.add_vline(x=data_rezerag, line=dict(color="gray", width=1.5, dash="dot"),
                          annotation_text="Re-zeragem A5b–A8b", annotation_position="top")
        fig.update_xaxes(title="Data")
        fig.update_yaxes(title="Desl. horizontal (mm)")
        fig.update_layout(height=460)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.caption("Assentamento vertical acumulado, ΔZ (mm)")
        fig2 = go.Figure()
        for a in sel:
            s = sub[sub[COLS["alvo"]] == a].sort_values(COLS["data"])
            fig2.add_trace(go.Scatter(x=s[COLS["data"]], y=s[COLS["dZ"]],
                                      mode="lines+markers", name=a))
        if st.session_state.get("mostrar_obra") and len(sub):
            adicionar_fases_obra(fig2, sub[COLS["data"]].min(), sub[COLS["data"]].max())
        # limiares verticais: o assentamento e negativo -> desenhar em -Va e -Vm
        fig2.add_hline(y=-Va, line_dash="dash", line_color="orange",
                       annotation_text=f"Alerta -{Va}", annotation_position="right")
        fig2.add_hline(y=-Vm, line_dash="dash", line_color="red",
                       annotation_text=f"Alarme -{Vm}", annotation_position="right")
        if e_santa_casa:
            fig2.add_vline(x=data_rezerag, line=dict(color="gray", width=1.5, dash="dot"),
                           annotation_text="Re-zeragem A5b–A8b", annotation_position="top")
        fig2.update_xaxes(title="Data")
        fig2.update_yaxes(title="ΔZ (mm)")
        fig2.update_layout(height=460)
        st.plotly_chart(fig2, use_container_width=True)


# =========================================================================
# SEPARADOR 4 — CELULAS DE CARGA
# =========================================================================
def separador_celulas(dados):
    cc = dados["celulas"]
    if not validar_colunas(cc, [COLS["data"], COLS["celula"], COLS["carga_atual"],
                                COLS["variacao"]], "Celulas de carga"):
        return
    cc = cc.sort_values(COLS["data"])
    st.subheader("Celulas de carga")
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
        fig.update_xaxes(title="Data"); fig.update_yaxes(title="Carga (kN)")
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
        fig2.update_xaxes(title="Data"); fig2.update_yaxes(title="Variacao (%)")
        fig2.update_layout(height=460)
        st.plotly_chart(fig2, use_container_width=True)
    cols = [c for c in [COLS["data"], COLS["ancoragem"], COLS["carga_atual"],
                        COLS["blocagem"], COLS["variacao"], COLS["estado"]]
            if c in sub.columns]
    st.dataframe(sub[cols], use_container_width=True)


# =========================================================================
# SEPARADOR 5 — PIEZOMETROS
# =========================================================================
def separador_piezometros(dados):
    pz = dados["piezo"]
    if not validar_colunas(pz, [COLS["data"], COLS["piezometro"], COLS["cota_agua"]],
                           "Piezometros"):
        return
    pz = pz.sort_values(COLS["data"])
    st.subheader("Piezometros — cota da agua")
    p = st.selectbox("Piezometro", sorted(pz[COLS["piezometro"]].dropna().unique()))
    sub = pz[pz[COLS["piezometro"]] == p].sort_values(COLS["data"])
    fig = go.Figure(go.Scatter(x=sub[COLS["data"]], y=sub[COLS["cota_agua"]],
                               mode="lines+markers", name="Cota da agua"))
    if st.session_state.get("mostrar_obra") and len(sub):
        adicionar_fases_obra(fig, sub[COLS["data"]].min(), sub[COLS["data"]].max())
    fig.update_xaxes(title="Data"); fig.update_yaxes(title="Cota da agua (m)")
    fig.update_layout(height=460)
    st.plotly_chart(fig, use_container_width=True)
    if st.session_state.get("mostrar_obra"):
        st.caption("Faixas = fases da obra (previstas). A descida do nivel de "
                   "agua durante a escavacao e coerente com rebaixamento "
                   "induzido pela propria escavacao.")
    else:
        st.caption("Sugestao: sobrepor a precipitacao diaria para avaliar a "
                   "resposta do nivel freatico a pluviosidade.")


# =========================================================================
# SEPARADOR 6 — PLANTA / DXF
# =========================================================================
# Layers estruturais tipicas (nomenclatura AIA/ISO) que costumam conter o
# contorno do recinto e a estrutura; usadas para pre-selecao inteligente.
LAYERS_ESTRUTURAIS_SUGERIDAS = [
    "S-BEAM", "S-COLS", "S-GRID", "A-FLOR", "S-WALL", "S-PILE",
]


def transformacao_semelhanca(p1_src, p1_dst, p2_src, p2_dst):
    """
    Transformacao de semelhanca 2D (translacao + rotacao + escala uniforme)
    que leva pontos do sistema do DESENHO (src) para o sistema dos ALVOS (dst),
    a partir de 2 pares de pontos correspondentes.

    Dois pontos chegam para fixar as 4 incognitas (2 translacao, 1 rotacao,
    1 escala). Devolve (aplicar, escala, angulo_graus), onde aplicar(x, y)
    converte uma coordenada do desenho para o sistema dos alvos.
    """
    import numpy as np
    x1, y1 = p1_src; X1, Y1 = p1_dst
    x2, y2 = p2_src; X2, Y2 = p2_dst
    dxs, dys = x2 - x1, y2 - y1        # vetor no sistema do desenho
    dXs, dYs = X2 - X1, Y2 - Y1        # vetor no sistema dos alvos
    Ls = np.hypot(dxs, dys)
    Ld = np.hypot(dXs, dYs)
    if Ls == 0:
        raise ValueError("Os dois pontos do desenho coincidem.")
    escala = Ld / Ls
    theta = np.arctan2(dYs, dXs) - np.arctan2(dys, dxs)
    c, s = np.cos(theta), np.sin(theta)

    def aplicar(x, y):
        xr, yr = np.asarray(x) - x1, np.asarray(y) - y1
        X = X1 + escala * (c * xr - s * yr)
        Y = Y1 + escala * (s * xr + c * yr)
        return X, Y

    return aplicar, escala, np.degrees(theta)


def separador_planta(dados):
    st.subheader("Planta do projeto (DXF)")

    if not TEM_EZDXF:
        st.error("Falta a biblioteca 'ezdxf'. Acrescenta 'ezdxf' ao "
                 "requirements.txt (e 'pip install ezdxf' se correres "
                 "localmente) para ativar a leitura de desenhos.")
        return

    st.caption("Carrega uma planta em DXF (por exemplo a planta de escavacao "
               "e contencao de um piso). Podes ve-la sozinha, ou ALINHA-LA com "
               "os alvos topograficos indicando 2 pontos de referencia — a app "
               "calcula a transformacao e sobrepoe os alvos coloridos pelo "
               "deslocamento.")

    dxf = st.file_uploader("Ficheiro DXF", type=["dxf"])
    if dxf is None:
        st.info("Carrega um DXF para ver a planta. Sugestao: a planta de "
                "escavacao do piso -1, -2 ou -3 mostra bem o contorno do "
                "recinto de contencao.")
        return

    conteudo = dxf.getvalue()

    try:
        layers = dxf_listar_layers(conteudo)
    except Exception as e:
        st.error(f"Nao consegui ler as layers do DXF. Detalhe: {e}")
        return

    sugeridas = [l for l in layers
                 if any(l.upper().startswith(p) for p in LAYERS_ESTRUTURAIS_SUGERIDAS)]

    st.write(f"O desenho tem **{len(layers)} layers**. Estao pre-selecionadas "
             f"as estruturais (contorno, estacas, grelha). Ajusta se quiseres:")
    layers_sel = st.multiselect("Layers a desenhar", sorted(layers),
                                default=sorted(sugeridas) if sugeridas else [])
    if not layers_sel:
        st.warning("Escolhe pelo menos uma layer para desenhar.")
        return

    try:
        segmentos = dxf_extrair_segmentos(conteudo, set(layers_sel))
    except Exception as e:
        st.error(f"Nao consegui extrair a geometria. Detalhe: {e}")
        return
    if not segmentos:
        st.warning("Nao encontrei geometria nas layers escolhidas. Tenta outras.")
        return

    # ---------------------------------------------------------------------
    # ALINHAMENTO OPCIONAL COM OS ALVOS
    # ---------------------------------------------------------------------
    alvos = dados["alvos"]
    tem_alvos = (not alvos.empty and COLS["M0"] in alvos.columns
                 and COLS["alvo"] in alvos.columns)

    alinhar = False
    aplicar = None
    if tem_alvos:
        alinhar = st.checkbox(
            "Alinhar a planta com os alvos (sobreposicao)", value=False,
            help="Precisas de indicar, para 2 alvos, onde eles estao no "
                 "desenho. A coordenada no sistema dos alvos ja vem do Excel.")

    if alinhar:
        # coordenadas dos alvos (sistema dos alvos) — uma linha por alvo (usar campanha mais recente)
        ult = alvos[alvos[COLS["data"]] == alvos[COLS["data"]].max()]
        lista_alvos = sorted(ult[COLS["alvo"]].astype(str).unique())

        st.markdown("**Ponto de referencia 1**")
        c1, c2, c3 = st.columns(3)
        a1 = c1.selectbox("Alvo 1", lista_alvos, key="a1")
        x1d = c2.number_input("X no desenho", value=0.0, key="x1d", format="%.2f")
        y1d = c3.number_input("Y no desenho", value=0.0, key="y1d", format="%.2f")

        st.markdown("**Ponto de referencia 2** (escolhe um bem afastado do 1)")
        d1, d2, d3 = st.columns(3)
        a2 = d1.selectbox("Alvo 2", lista_alvos,
                          index=min(len(lista_alvos) - 1, 1), key="a2")
        x2d = d2.number_input("X no desenho", value=0.0, key="x2d", format="%.2f")
        y2d = d3.number_input("Y no desenho", value=0.0, key="y2d", format="%.2f")

        st.caption("Como obter o X,Y no desenho: abre o DXF no AutoCAD ou num "
                   "visualizador, aponta o cursor ao sitio onde o alvo esta, e "
                   "le as coordenadas. Quanto mais afastados os 2 alvos, melhor.")

        def coord_alvo(nome):
            linha = ult[ult[COLS["alvo"]].astype(str) == nome]
            return float(linha[COLS["M0"]].iloc[0]), float(linha[COLS["P0"]].iloc[0])

        if a1 == a2:
            st.warning("Escolhe dois alvos diferentes.")
        elif (x1d, y1d) == (0.0, 0.0) or (x2d, y2d) == (0.0, 0.0):
            st.info("Preenche as coordenadas dos 2 pontos no desenho para "
                    "calcular o alinhamento.")
        else:
            try:
                P1 = coord_alvo(a1)
                P2 = coord_alvo(a2)
                aplicar, escala, ang = transformacao_semelhanca(
                    (x1d, y1d), P1, (x2d, y2d), P2)
                m1, m2, m3 = st.columns(3)
                m1.metric("Escala desenho->alvos", f"{escala:.4f}")
                m2.metric("Rotacao", f"{ang:.1f}°")
                m3.metric("Estado", "Alinhado")
                if not (0.5 < escala < 2.0):
                    st.warning("A escala calculada e invulgar. Confirma as "
                               "coordenadas dos pontos no desenho — pode haver "
                               "troca de X/Y ou de ponto.")
            except Exception as e:
                st.error(f"Nao consegui calcular o alinhamento: {e}")

    # ---------------------------------------------------------------------
    # DESENHAR
    # ---------------------------------------------------------------------
    fig = go.Figure()
    cores = {}
    paleta = ["#333333", "#1f77b4", "#d62728", "#2ca02c", "#9467bd",
              "#8c564b", "#e377c2", "#ff7f0e"]
    for i, lay in enumerate(sorted(set(s[2] for s in segmentos))):
        cores[lay] = paleta[i % len(paleta)]

    mostrados = set()
    for xs, ys, lay in segmentos:
        if aplicar is not None:
            X, Y = aplicar(xs, ys)
            xs, ys = list(X), list(Y)
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines",
            line=dict(color=cores[lay], width=1),
            name=lay, legendgroup=lay,
            showlegend=(lay not in mostrados), hoverinfo="skip",
        ))
        mostrados.add(lay)

    # sobrepor alvos se estiver alinhado
    if aplicar is not None and tem_alvos:
        ult = alvos[alvos[COLS["data"]] == alvos[COLS["data"]].max()]
        fig.add_trace(go.Scatter(
            x=ult[COLS["M0"]], y=ult[COLS["P0"]],
            mode="markers+text",
            marker=dict(size=10, color=ult[COLS["desl_h"]], colorscale="YlOrRd",
                        colorbar=dict(title="Desl. h (mm)"), cmin=0,
                        line=dict(width=1, color="black")),
            text=ult[COLS["alvo"]].astype(str), textposition="top center",
            textfont=dict(size=8), name="Alvos",
            customdata=ult[COLS["desl_h"]],
            hovertemplate="Alvo %{text}<br>Desl. h: %{customdata:.1f} mm<extra></extra>",
        ))

    eixo = "sistema dos alvos (M, P)" if aplicar is not None else "X, Y local do desenho"
    fig.update_layout(
        height=700, xaxis_title=eixo, yaxis_title=eixo,
        yaxis=dict(scaleanchor="x", scaleratio=1),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(l=0, r=0, t=30, b=0),
    )
    st.plotly_chart(fig, use_container_width=True)

    if aplicar is not None:
        st.success("Planta alinhada com os alvos. Verifica visualmente se o "
                   "contorno bate certo com a nuvem de alvos; se nao, ajusta os "
                   "pontos de referencia.")
    else:
        st.caption("Planta em coordenadas locais do desenho. Ativa o "
                   "alinhamento acima para a sobrepor aos alvos.")


# =========================================================================
# SEPARADOR 7 — GEOLOGIA
# =========================================================================
def desenhar_coluna_litologica(fig, sondagem, x_centro=0, largura=0.8,
                               em_cota=False, cota_terreno=0):
    """Desenha a coluna litologica de uma sondagem como retangulos coloridos.
    Se em_cota=True, converte profundidade em cota (cota_terreno - prof)."""
    for topo, base, unidade in GEO_LITOLOGIA[sondagem]:
        y0 = cota_terreno - topo if em_cota else topo
        y1 = cota_terreno - base if em_cota else base
        cor = GEO_CORES_LITO.get(unidade, "#cccccc")
        fig.add_shape(type="rect",
                      x0=x_centro - largura/2, x1=x_centro + largura/2,
                      y0=y0, y1=y1, fillcolor=cor, opacity=0.7,
                      line=dict(color="black", width=0.5), layer="below")


def _zona_por_N(n):
    """
    Ponte de LEITURA N -> zona geotecnica, segundo os intervalos do proprio
    relatorio (Quadros V-VII): ZG5 (SPT 11-30), ZG4 (31-56), ZG3/ZG2/ZG1 (>=60,
    distinguidas pelo RQD que NAO temos por ponto). Serve so para colorir o
    ponto SPT e dar leitura rapida; NAO define fronteiras de camada.
    """
    if n < 11:
        return "ZG6/ZG5"
    if n <= 30:
        return "ZG5"
    if n <= 56:
        return "ZG4"
    return "ZG3–ZG1 (nega)"


ZONA_CORES = {
    "ZG6/ZG5": "#d9822b",
    "ZG5": "#c9a227",
    "ZG4": "#7fa650",
    "ZG3–ZG1 (nega)": "#3f6f3f",
}


def separador_geologia(dados):
    st.subheader("Geologia do terreno (Relatorio ENGGEO, proc. 220216)")
    st.caption("Leitura integrada por sondagem: a coluna litologica, os ensaios "
               "SPT e o zonamento geotecnico partilham o eixo de profundidade, "
               "para se lerem em conjunto. Sob ~0,5 m de aterro, o terreno e "
               "essencialmente grés dos 'Grés Superiores' (C1As), com calcario "
               "(C1A) apenas no fundo do SC8. A deformacao nao se explica por "
               "uma camada mole — nao existe — mas pelo grau de consolidacao do "
               "grés, que cresce com a profundidade (o SPT sobe de ~11-30 para "
               "nega).")

    sonds = list(GEO_LITOLOGIA.keys())

    col_sel, col_info = st.columns([1, 2])
    with col_sel:
        modo = st.radio("Vista", ["Uma sondagem (detalhe)", "As quatro (comparar)"])
        if modo == "Uma sondagem (detalhe)":
            sond = st.selectbox("Sondagem", sonds)
        else:
            sond = None
    with col_info:
        st.caption("A coluna colorida a esquerda de cada sondagem e a litologia; "
                   "os pontos e a linha sao o SPT (N pancadas), no mesmo eixo de "
                   "profundidade. A cor do ponto SPT indica a zona geotecnica "
                   "provavel pelo valor de N (ver tabela em baixo). A linha azul "
                   "tracejada e o nivel freatico.")

    alvo_sonds = [sond] if sond else sonds
    n_col = len(alvo_sonds)

    fig = go.Figure()
    LARG_LITO = 0.12
    SPT_MAX = 65.0

    for i, s in enumerate(alvo_sonds):
        x_base = i
        x0_lito = x_base - 0.45
        x1_lito = x_base - 0.45 + LARG_LITO

        # coluna litologica
        for topo, base, unidade in GEO_LITOLOGIA[s]:
            cor = GEO_CORES_LITO.get(unidade, "#cccccc")
            fig.add_shape(type="rect", x0=x0_lito, x1=x1_lito, y0=topo, y1=base,
                          fillcolor=cor, opacity=0.85,
                          line=dict(color="black", width=0.4), layer="below")

        # SPT reescalado a direita da coluna litologica
        x_spt0 = x1_lito + 0.03
        x_spt1 = x_base + 0.45

        def _xN(n, a=x_spt0, b=x_spt1):
            return a + (n / SPT_MAX) * (b - a)

        ensaios = GEO_SPT[s]
        profs = [e[0] for e in ensaios]
        ns = [e[1] for e in ensaios]
        xs = [_xN(n) for n in ns]
        zonas = [_zona_por_N(n) for n in ns]
        cores_pt = [ZONA_CORES[z] for z in zonas]

        fig.add_trace(go.Scatter(
            x=xs, y=profs, mode="lines",
            line=dict(color="rgba(90,90,90,0.55)", width=1.5),
            showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(
            x=xs, y=profs, mode="markers",
            marker=dict(size=7, color=cores_pt,
                        line=dict(color="black", width=0.4)),
            customdata=list(zip(ns, zonas)),
            hovertemplate=(f"{s}<br>Prof: %{{y:.1f}} m<br>"
                           "N: %{customdata[0]}<br>Zona: %{customdata[1]}"
                           "<extra></extra>"),
            showlegend=False))

        # nega (N=60)
        fig.add_shape(type="line", x0=_xN(60), x1=_xN(60),
                      y0=0, y1=GEO_LITOLOGIA[s][-1][1],
                      line=dict(color="gray", width=1, dash="dot"), layer="below")

        # nivel freatico
        nf = GEO_SONDAGENS[s]["nf_prof"]
        if nf is not None:
            fig.add_shape(type="line", x0=x0_lito, x1=x_spt1, y0=nf, y1=nf,
                          line=dict(color="blue", width=2, dash="dash"))

        fig.add_annotation(x=x_base, y=1.0, yref="paper", showarrow=False,
                           text=f"<b>{s}</b>", font=dict(size=12))
        for nval in (0, 30, 60):
            fig.add_annotation(x=_xN(nval), y=-0.6, showarrow=False,
                               text=str(nval), font=dict(size=8, color="gray"))

    # legendas fantasma
    for unidade, cor in GEO_CORES_LITO.items():
        if any(any(u == unidade for _, _, u in GEO_LITOLOGIA[s]) for s in alvo_sonds):
            fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
                                     marker=dict(size=12, color=cor, symbol="square"),
                                     name=f"Litologia: {unidade}"))
    for zona, cor in ZONA_CORES.items():
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
                                 marker=dict(size=10, color=cor),
                                 name=f"SPT→{zona}"))
    fig.add_trace(go.Scatter(x=[None], y=[None], mode="lines",
                             line=dict(color="blue", dash="dash"),
                             name="Nivel freatico"))
    fig.add_trace(go.Scatter(x=[None], y=[None], mode="lines",
                             line=dict(color="gray", dash="dot"),
                             name="Nega (N=60)"))

    fig.update_yaxes(autorange="reversed", title="Profundidade (m)")
    fig.update_xaxes(showticklabels=False, showgrid=False, zeroline=False,
                     range=[-0.6, n_col - 0.4])
    fig.update_layout(height=640, legend_title="Legenda",
                      margin=dict(l=0, r=0, t=30, b=10),
                      legend=dict(font=dict(size=10)))
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Como ler: para cada sondagem, a barra colorida a esquerda e a "
               "litologia; a curva a direita e o SPT (escala 0–60, com a linha "
               "pontilhada na nega). Onde o SPT sobe, o grés esta mais "
               "consolidado — e e ai que a rigidez do macico aumenta. "
               "Profundidades de cada zona ZG NAO estao definidas no relatorio "
               "por ponto; a cor do SPT e a zona PROVAVEL pelo valor de N.")

    st.divider()

    st.markdown("#### Zonamento geotecnico e parametros de projeto")
    st.caption("Parametros propostos (Quadros V-VII do relatorio) que alimentam "
               "a modelacao numerica da contencao. A coluna 'SPT tipico' e a "
               "ponte para o perfil acima: e por ela que se le em que zona esta "
               "cada troco de terreno.")

    faixa_spt = {
        "ZG6": "aterro", "ZG5": "11–30", "ZG4": "31–56",
        "ZG3": "≥60 (RQD 0–25%)", "ZG2": "≥60 (RQD 45–75%)",
        "ZG1": "≥60 (RQD 76–100%)",
    }
    zt = pd.DataFrame(GEO_ZONAMENTO)
    zt.insert(2, "SPT tipico (N)", zt["Zona"].map(faixa_spt))
    st.dataframe(zt, use_container_width=True, hide_index=True)
    st.caption("gama: peso volumico | c': coesao | fi': angulo de atrito | "
               "E': modulo de deformabilidade. Zonas ZG3-ZG1 (rocha) com c' e E' "
               "em MPa/GPa; ZG6-ZG4 (solo/grés brando) em kPa/MPa. Nota: as tres "
               "zonas de nega (ZG3-ZG1) distinguem-se pelo RQD, que o SPT sozinho "
               "nao mede — por isso o perfil agrupa-as como 'nega'.")



# =========================================================================
# SEPARADOR 8 — CRONOGRAMA DA OBRA
# =========================================================================
def separador_obra(dados):
    st.subheader("Cronograma da obra (Plano de Trabalhos)")
    st.caption("Plano de trabalhos da empreitada (Alves Ribeiro / HCI, "
               "05/05/2025). Sao datas PREVISTAS — o planeado, que pode diferir "
               "do executado. A janela de instrumentacao (out-dez 2025) esta "
               "assinalada para veres que fases estavam ativas durante a "
               "monitorizacao.")

    # Gantt simples com barras horizontais
    fig = go.Figure()
    for i, (nome, ini, fim) in enumerate(FASES_OBRA):
        t0 = pd.to_datetime(ini)
        t1 = pd.to_datetime(fim)
        cor = CORES_FASES[i % len(CORES_FASES)]
        fig.add_trace(go.Scatter(
            x=[t0, t1], y=[nome, nome], mode="lines",
            line=dict(color=cor, width=16),
            hovertemplate=f"{nome}<br>{ini} a {fim}<extra></extra>",
            showlegend=False,
        ))

    # faixa da janela de instrumentacao
    alvos = dados.get("alvos")
    if alvos is not None and not alvos.empty and COLS["data"] in alvos.columns:
        d0 = alvos[COLS["data"]].min()
        d1 = alvos[COLS["data"]].max()
        fig.add_vrect(x0=d0, x1=d1, fillcolor="crimson", opacity=0.12,
                      line_width=0,
                      annotation_text="Instrumentacao (dados)",
                      annotation_position="top left")

    fig.update_xaxes(title="Data")
    fig.update_layout(height=460, margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig, use_container_width=True)

    st.caption("Nota para a tese: apresentar estas datas como PLANEADAS. Se "
               "tiveres os autos de obra (datas reais de execucao), o cruzamento "
               "com a instrumentacao passa a ser rigoroso; sem eles, e uma "
               "aproximacao defensavel desde que assinalada como tal.")

    # tabela do plano
    tabela = pd.DataFrame(
        [{"Fase": n, "Inicio": i, "Conclusao": f} for n, i, f in FASES_OBRA])
    st.dataframe(tabela, use_container_width=True, hide_index=True)


# =========================================================================
# SEPARADOR 0 — PAGINA INICIAL (HOME)
# =========================================================================
def separador_home(dados):
    # ---- BANNER no topo: gradiente azul ---------------------------------
    banner = (
        "<div style='border-radius:12px; height:190px; margin-bottom:18px; "
        "background:linear-gradient(120deg, #0f2037 0%, #1f3a5f 55%, "
        "#2e5c8a 100%); display:flex; align-items:center; padding-left:36px;'>"
        "<div style='color:white;'>"
        "<div style='font-size:3.2em; font-weight:800; letter-spacing:3px; "
        "line-height:1;'>IMS</div>"
        "<div style='font-size:1.2em; opacity:0.92; margin-top:8px;'>"
        "Instrumentation Monitoring System</div>"
        "<div style='font-size:0.95em; opacity:0.8; margin-top:12px; "
        "max-width:560px;'>Analise e visualizacao de instrumentacao geotecnica "
        "— deslocamentos, velocidades, sinais precursores, geologia e "
        "sequencia de obra.</div>"
        "</div></div>"
    )
    st.markdown(banner, unsafe_allow_html=True)

    # ---- identificacao da obra + numeros-chave --------------------------
    col_id, col_num = st.columns([1.3, 2])
    with col_id:
        st.markdown("#### Obra")
        st.markdown(
            "**Hotel Eden, Estoril**  \n"
            "Reformulacao — escavacao e contencao periferica  \n"
            "Monte Estoril, Cascais  \n"
            "_Back-analysis de instrumentacao_")

    with col_num:
        st.markdown("#### Instrumentacao monitorizada")
        # calcular numeros reais a partir dos dados
        n_inc = dados["perfis"][COLS["inclinometro"]].nunique() if not dados["perfis"].empty else 0
        n_alvos = dados["alvos"][COLS["alvo"]].nunique() if not dados["alvos"].empty else 0
        n_camp = dados["alvos"][COLS["data"]].nunique() if not dados["alvos"].empty else 0
        n_cel = dados["celulas"][COLS["celula"]].nunique() if not dados["celulas"].empty else 0
        n_pz = dados["piezo"][COLS["piezometro"]].nunique() if not dados["piezo"].empty else 0
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Inclinometros", n_inc)
        c2.metric("Alvos topogr.", n_alvos)
        c3.metric("Campanhas", n_camp)
        c4.metric("Celulas carga", n_cel)
        c5.metric("Piezometros", n_pz)

    st.divider()

    # ---- cartoes das areas de analise -----------------------------------
    st.markdown("#### O que podes explorar")
    st.caption("Usa os separadores no topo para navegar. Aqui fica o mapa geral "
               "de cada area.")

    cartoes = [
        ("Visao geral 3D", "Alvos no espaco com a geometria da obra: contorno do "
         "recinto, edificios vizinhos e vetores de deslocamento amplificados."),
        ("Inclinometros", "Perfil deformado em profundidade, evolucao no tempo, "
         "velocidade e detecao de sinais precursores. Sobreposicao da geologia."),
        ("Alvos (2D)", "Series temporais de deslocamento horizontal e "
         "assentamento vertical, por edificio e por alvo."),
        ("Celulas de carga", "Carga nas ancoragens vs. blocagem, com limiares "
         "de alerta (15%) e alarme (25%)."),
        ("Piezometros", "Evolucao da cota da agua subterranea, para relacionar "
         "com a escavacao e a pluviosidade."),
        ("Geologia", "Colunas litologicas das sondagens, ensaios SPT em "
         "profundidade e zonamento geotecnico (ZG1-ZG6)."),
        ("Obra", "Cronograma do plano de trabalhos, com a janela de "
         "instrumentacao assinalada. Sobrepoe-se aos graficos temporais."),
        ("Planta (DXF)", "Leitura de plantas de escavacao em DXF, com opcao de "
         "alinhamento aos alvos por pontos de referencia."),
    ]
    # desenhar em grelha de 2 colunas
    for i in range(0, len(cartoes), 2):
        cols = st.columns(2)
        for j, col in enumerate(cols):
            if i + j < len(cartoes):
                titulo, desc = cartoes[i + j]
                with col:
                    st.markdown(
                        f"<div style='border:1px solid #e0e4e8; border-radius:8px; "
                        f"padding:14px 16px; margin-bottom:10px; background:#fafbfc;'>"
                        f"<div style='font-weight:600; color:#1f3a5f; "
                        f"font-size:1.05em; margin-bottom:4px;'>{titulo}</div>"
                        f"<div style='color:#5a6b7b; font-size:0.92em;'>{desc}</div>"
                        f"</div>",
                        unsafe_allow_html=True)

    st.divider()
    st.caption("Nota metodologica: a integracao dos perfis inclinometricos "
               "assume a base fixa. Datas da sequencia de obra sao as previstas "
               "no plano de trabalhos. Superficie 3D e interpolada entre alvos "
               "medidos — apoio visual, nao modelo do macico.")


# =========================================================================
# PRINCIPAL
# =========================================================================
def main():
    st.sidebar.title("The Eden - Instrumentacao")
    fonte_op = st.sidebar.radio("Ficheiro de dados",
                                ["Usar o ficheiro ao lado do script", "Carregar manualmente"])
    if fonte_op == "Carregar manualmente":
        up = st.sidebar.file_uploader("Excel de instrumentacao", type=["xlsx"])
        if up is None:
            st.info("Carrega o Excel na barra lateral para comecar.")
            st.stop()
        fonte = up
    else:
        fonte = FICHEIRO_EXCEL
        if not Path(FICHEIRO_EXCEL).exists():
            st.error(f"Nao encontrei '{FICHEIRO_EXCEL}'. Poe o Excel na pasta do "
                     f"script ou usa 'Carregar manualmente'.")
            st.stop()
    try:
        dados = carregar_dados(fonte)
    except Exception as e:
        st.error(f"Nao consegui ler o Excel. Detalhe: {e}")
        st.stop()

    st.sidebar.divider()
    st.sidebar.subheader("Detecao de precursores")
    limiar_vel = st.sidebar.slider("Limiar de velocidade (mm/dia)", 0.1, 3.0,
                                   LIMIAR_VEL_DEFEITO, 0.05)
    fator_acel = st.sidebar.slider("Fator de aceleracao (x)", 1.2, 3.0,
                                   FATOR_ACEL_DEFEITO, 0.1)

    st.sidebar.divider()
    st.sidebar.subheader("Sequencia da obra")
    mostrar_obra = st.sidebar.checkbox(
        "Sobrepor fases da obra aos graficos temporais", value=True,
        help="Mostra, nos graficos com data no eixo, as fases do plano de "
             "trabalhos (Alves Ribeiro). Datas PREVISTAS — nao as reais.")
    st.session_state["mostrar_obra"] = mostrar_obra

    if not TEM_SCIPY:
        st.sidebar.info("Instala 'scipy' para ativar a superficie 3D interpolada.")

    if not TEM_EZDXF:
        st.sidebar.info("Instala 'ezdxf' para ativar a leitura de plantas DXF.")

    thome, t3d, tinc, talv, tcc, tpz, tgeo, tobra, tplan = st.tabs(
        ["Inicio", "Visao geral 3D", "Inclinometros", "Alvos (2D)",
         "Celulas de carga", "Piezometros", "Geologia", "Obra", "Planta (DXF)"])
    with thome:
        separador_home(dados)
    with t3d:
        separador_3d(dados)
    with tinc:
        separador_inclinometros(dados, limiar_vel, fator_acel)
    with talv:
        separador_alvos_2d(dados)
    with tcc:
        separador_celulas(dados)
    with tpz:
        separador_piezometros(dados)
    with tgeo:
        separador_geologia(dados)
    with tobra:
        separador_obra(dados)
    with tplan:
        separador_planta(dados)


if __name__ == "__main__":
    main()
