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
# CRITERIOS DE ALERTA E ALARME DOS ALVOS TOPOGRAFICOS
# Fonte: relatorio de alvos topograficos da obra (tabela de criterios).
# Dependem do alcado (altura de contencao). Deslocamentos em mm.
# =========================================================================
CRIT_ALVO_17 = {"h_alerta": 20, "h_alarme": 40, "v_alerta": 10, "v_alarme": 15}
CRIT_ALVO_24 = {"h_alerta": 30, "h_alarme": 40, "v_alerta": 10, "v_alarme": 15}

# alcados de contencao por grupo de criterio (da tabela)
ALCADOS_17 = {"AB", "CD", "BF", "PQ"}
ALCADOS_24 = {"FG", "GH", "JK", "KL", "MNO", "OP"}

# Prefixos dos alvos dos edificios vizinhos (sem criterio explicito na tabela).
# Adota-se o criterio conservador (17m) — ASSUMIDO, a assinalar na tese.
PREFIXOS_EDIFICIO = {"A", "B", "C"}


def criterio_do_alvo(nome_alvo):
    """
    Devolve (dict_criterio, origem_texto, assumido_bool) para um alvo,
    a partir do prefixo de letras do seu nome.
    Alvos de contencao: criterio real do alcado.
    Alvos de edificios vizinhos: criterio 17m conservador, marcado como assumido.
    """
    import re
    m = re.match(r"^([A-Z]+)", str(nome_alvo))
    pref = m.group(1) if m else ""
    if pref in ALCADOS_17:
        return CRIT_ALVO_17, f"Alcado {pref} (17 m)", False
    if pref in ALCADOS_24:
        return CRIT_ALVO_24, f"Alcado {pref} (24 m)", False
    # edificio vizinho (A/B/C) ou desconhecido -> conservador, assumido
    return CRIT_ALVO_17, "Edificio vizinho (17 m, assumido)", True


def estado_alvo(desl_h, delta_z, criterio):
    """Devolve 'ALARME', 'ALERTA' ou 'OK' comparando desl. horizontal e
    vertical (em modulo) com os limiares do criterio."""
    import numpy as np
    dh = abs(desl_h) if pd.notna(desl_h) else 0
    dz = abs(delta_z) if pd.notna(delta_z) else 0
    if dh >= criterio["h_alarme"] or dz >= criterio["v_alarme"]:
        return "ALARME"
    if dh >= criterio["h_alerta"] or dz >= criterio["v_alerta"]:
        return "ALERTA"
    return "OK"


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


def _bloco_edificio(fig, grp, cor, nome):
    """Desenha um bloco/casa esquematico (caixa 3D) na zona dos alvos de um
    edificio vizinho. Posicao real (assente nos alvos); forma ILUSTRATIVA."""
    import numpy as np
    x0, x1 = grp[COLS["M0"]].min(), grp[COLS["M0"]].max()
    y0, y1 = grp[COLS["P0"]].min(), grp[COLS["P0"]].max()
    z0 = grp[COLS["Z0"]].min()
    # dar alguma margem e uma altura simbolica
    mx = max((x1 - x0) * 0.15, 1.5); my = max((y1 - y0) * 0.15, 1.5)
    x0 -= mx; x1 += mx; y0 -= my; y1 += my
    altura = 6.0
    zt = z0 + altura
    # 8 vertices da caixa
    xs = [x0, x1, x1, x0, x0, x1, x1, x0]
    ys = [y0, y0, y1, y1, y0, y0, y1, y1]
    zs = [z0, z0, z0, z0, zt, zt, zt, zt]
    # faces da caixa (indices dos vertices) via Mesh3d
    fig.add_trace(go.Mesh3d(
        x=xs, y=ys, z=zs,
        i=[0,0,0,4,4,1,1,2,3,0,4,3],
        j=[1,2,4,5,7,2,5,3,7,3,5,2],
        k=[2,3,5,7,6,6,6,7,4,4,6,6],
        color=cor, opacity=0.25, name=nome, hoverinfo="name",
        showlegend=False,
    ))
    # telhado simples: uma linha ao cimo para dar ar de "casa"
    xm = (x0 + x1) / 2
    fig.add_trace(go.Scatter3d(
        x=[x0, xm, x1], y=[(y0+y1)/2]*3, z=[zt, zt+2.5, zt],
        mode="lines", line=dict(color=cor, width=3),
        showlegend=False, hoverinfo="skip",
    ))


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
    st.caption("Contorno castanho = envolvente dos alvos da contencao (limite "
               "aproximado do recinto). Alvos coloridos por estado (verde OK, "
               "laranja alerta, vermelho alarme). Blocos = edificios vizinhos "
               "(posicao real, forma esquematica). Setas = deslocamento "
               "acumulado amplificado.")

    datas = sorted(alvos[COLS["data"]].dropna().unique())

    # ---- filtro por fase de escavacao ------------------------------------
    # fases da obra que se sobrepoem ao periodo dos alvos
    d_ini, d_fim = pd.to_datetime(datas[0]), pd.to_datetime(datas[-1])
    fases_disp = [("Todas as campanhas", None, None)]
    for nome, ini, fim in FASES_OBRA:
        t0, t1 = pd.to_datetime(ini), pd.to_datetime(fim)
        if not (t1 < d_ini or t0 > d_fim):
            fases_disp.append((nome, t0, t1))

    col_top1, col_top2 = st.columns([1, 1])
    with col_top1:
        fase_nome = st.selectbox(
            "Filtrar por fase da obra (planeamento real)",
            [f[0] for f in fases_disp],
            help="Mostra so as campanhas dentro da fase escolhida. Datas do "
                 "plano de trabalhos.")
    fase = next(f for f in fases_disp if f[0] == fase_nome)

    # campanhas dentro da fase
    if fase[1] is None:
        datas_fase = datas
    else:
        datas_fase = [d for d in datas
                      if fase[1] <= pd.to_datetime(d) <= fase[2]]
        if not datas_fase:
            st.warning("Nenhuma campanha de alvos dentro desta fase.")
            datas_fase = datas

    with col_top2:
        data_sel = st.select_slider(
            "Campanha", options=datas_fase, value=datas_fase[-1],
            format_func=lambda d: pd.to_datetime(d).strftime("%d/%m/%Y"))

    col_b, col_c = st.columns([1, 1])
    with col_b:
        fator = st.slider("Amplificacao do deslocamento", 50, 2000, 500, 50)
    with col_c:
        mostrar_contorno = st.checkbox("Contorno do recinto", value=True)
        mostrar_edificios = st.checkbox("Blocos dos edificios vizinhos", value=True)

    campanha = alvos[alvos[COLS["data"]] == data_sel].copy()

    fig = go.Figure()

    # ---- contorno do recinto ---------------------------------------------
    cont = contorno_recinto(campanha)
    if mostrar_contorno and cont is not None:
        Ms, Ps, Zs = cont
        fig.add_trace(go.Scatter3d(
            x=Ms, y=Ps, z=Zs, mode="lines",
            line=dict(color="saddlebrown", width=6),
            name="Contorno do recinto", hoverinfo="skip",
        ))

    # ---- alvos por grupo, coloridos por ESTADO (alerta/alarme) -----------
    COR_ESTADO = {"OK": "#2ca02c", "ALERTA": "#ff9900", "ALARME": "#d62728"}
    for chave, grp in campanha.groupby(COLS["edificio"]):
        tipo, etiqueta = classificar_grupo(chave)
        x0 = grp[COLS["M0"]].to_numpy()
        y0 = grp[COLS["P0"]].to_numpy()
        z0 = grp[COLS["Z0"]].to_numpy()
        dx = grp[COLS["dM"]].to_numpy() / 1000.0 * fator
        dy = grp[COLS["dP"]].to_numpy() / 1000.0 * fator
        dz = grp[COLS["dZ"]].to_numpy() / 1000.0 * fator
        dh = grp[COLS["desl_h"]].to_numpy()
        dzv = grp[COLS["dZ"]].to_numpy()
        nomes = grp[COLS["alvo"]].astype(str).to_numpy()

        # estado de cada alvo
        cores_alvo = []
        estados = []
        for k, nome_a in enumerate(nomes):
            crit, _, _ = criterio_do_alvo(nome_a)
            est = estado_alvo(dh[k], dzv[k], crit)
            estados.append(est)
            cores_alvo.append(COR_ESTADO[est])

        nome_leg = chave if tipo == "edificio" else f"Contencao {etiqueta}"
        fig.add_trace(go.Scatter3d(
            x=x0 + dx, y=y0 + dy, z=z0 + dz, mode="markers+text",
            marker=dict(size=6, color=cores_alvo,
                        line=dict(width=1, color="black")),
            text=nomes, textposition="top center", textfont=dict(size=8),
            name=nome_leg, customdata=np.array(list(zip(dh, estados)), dtype=object),
            hovertemplate="Alvo %{text}<br>Desl. h: %{customdata[0]:.1f} mm"
                          "<br>Estado: %{customdata[1]}<extra></extra>",
            showlegend=False,
        ))
        for i in range(len(x0)):
            fig.add_trace(go.Scatter3d(
                x=[x0[i], x0[i] + dx[i]], y=[y0[i], y0[i] + dy[i]],
                z=[z0[i], z0[i] + dz[i]],
                mode="lines", line=dict(color="crimson", width=3),
                showlegend=False, hoverinfo="skip",
            ))

        # bloco esquematico do edificio vizinho
        if mostrar_edificios and tipo == "edificio":
            cor_ed = CORES_EDIFICIO.get(chave, "#888888")
            _bloco_edificio(fig, grp, cor_ed, chave)

    # legenda manual do estado (cores) + edificios
    for est, cor in COR_ESTADO.items():
        fig.add_trace(go.Scatter3d(x=[None], y=[None], z=[None], mode="markers",
                                   marker=dict(size=8, color=cor), name=f"Alvo {est}"))

    fig.update_layout(
        height=740,
        scene=dict(xaxis_title="M (m)", yaxis_title="P (m)", zaxis_title="Z (m)",
                   aspectmode="data"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(size=9)),
        margin=dict(l=0, r=0, t=30, b=0),
    )
    st.plotly_chart(fig, use_container_width=True)

    # ---- metricas + resumo de alarmes ------------------------------------
    desl_h_all = campanha[COLS["desl_h"]].to_numpy()
    nomes_all = campanha[COLS["alvo"]].astype(str).to_numpy()
    edif_all = campanha[COLS["edificio"]].to_numpy()
    # contar estados
    n_alarme = n_alerta = 0
    alarmados = []
    for k, nm in enumerate(nomes_all):
        crit, _, _ = criterio_do_alvo(nm)
        est = estado_alvo(desl_h_all[k], campanha[COLS["dZ"]].to_numpy()[k], crit)
        if est == "ALARME":
            n_alarme += 1; alarmados.append(nm)
        elif est == "ALERTA":
            n_alerta += 1

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alvos na campanha", len(campanha))
    c2.metric("Desl. horiz. max (mm)", f"{np.nanmax(desl_h_all):.1f}")
    c3.metric("Em alerta", n_alerta)
    c4.metric("Em alarme", n_alarme)
    if n_alarme:
        st.error(f"{n_alarme} alvo(s) em ALARME nesta campanha: "
                 f"{', '.join(alarmados)}.")
    elif n_alerta:
        st.warning(f"{n_alerta} alvo(s) em alerta nesta campanha.")
    else:
        st.success("Nenhum alvo ultrapassa criterios nesta campanha.")
    idx = int(np.nanargmax(desl_h_all))
    st.caption(f"Mais afetado: {nomes_all[idx]} ({np.nanmax(desl_h_all):.1f} mm), "
               f"em {edif_all[idx]}. Estados de edificios vizinhos usam o "
               f"criterio 17 m assumido.")


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

    # ---- painel de estado: verificar TODOS os alvos na ultima campanha ----
    ult = alvos[alvos[COLS["data"]] == alvos[COLS["data"]].max()].copy()
    linhas_estado = []
    for _, r in ult.iterrows():
        crit, origem, assumido = criterio_do_alvo(r[COLS["alvo"]])
        est = estado_alvo(r[COLS["desl_h"]], r[COLS["dZ"]], crit)
        linhas_estado.append({
            "Alvo": str(r[COLS["alvo"]]), "Estado": est,
            "Desl. H (mm)": round(r[COLS["desl_h"]], 1),
            "ΔZ (mm)": round(r[COLS["dZ"]], 1),
            "Criterio": origem, "assumido": assumido,
        })
    df_est = pd.DataFrame(linhas_estado)
    n_alarme = (df_est["Estado"] == "ALARME").sum()
    n_alerta = (df_est["Estado"] == "ALERTA").sum()

    data_ult = pd.to_datetime(alvos[COLS["data"]].max()).strftime("%d/%m/%Y")
    c1, c2, c3 = st.columns(3)
    c1.metric(f"Alvos OK ({data_ult})", int((df_est["Estado"] == "OK").sum()))
    c2.metric("Em ALERTA", int(n_alerta))
    c3.metric("Em ALARME", int(n_alarme))

    if n_alarme:
        st.error(f"{n_alarme} alvo(s) em ALARME na ultima campanha — "
                 f"deslocamento acima do criterio de alarme.")
    elif n_alerta:
        st.warning(f"{n_alerta} alvo(s) em ALERTA na ultima campanha.")
    else:
        st.success("Nenhum alvo ultrapassa os criterios na ultima campanha.")

    with st.expander("Ver criterios e estado de todos os alvos"):
        st.caption("Criterios (relatorio da obra): contencao 17 m (alcados AB, "
                   "CD, BF, PQ) -> H alerta 20 / alarme 40 mm; contencao 24 m "
                   "(FG, GH, JK, KL, MNO, OP) -> H alerta 30 / alarme 40 mm; "
                   "vertical alerta 10 / alarme 15 mm em ambos. Edificios "
                   "vizinhos (A, B, C): criterio 17 m ASSUMIDO (nao consta da "
                   "tabela — confirmar com o projetista).")
        def cor_estado(row):
            c = {"ALARME": "#ffcccc", "ALERTA": "#fff2cc"}.get(row["Estado"], "")
            return [f"background-color:{c}" if c else "" for _ in row]
        st.dataframe(
            df_est.drop(columns=["assumido"]).style.apply(cor_estado, axis=1),
            use_container_width=True, hide_index=True)

    st.divider()

    edificios = sorted(alvos[COLS["edificio"]].dropna().unique())
    edi = st.selectbox("Edificio / elemento", edificios)
    sub = alvos[alvos[COLS["edificio"]] == edi]
    lista = sorted(sub[COLS["alvo"]].dropna().unique())
    sel = st.multiselect("Alvos", lista, default=lista[:min(5, len(lista))])

    # criterio do edificio selecionado (para as linhas de limiar)
    crit_edi, origem_edi, assumido_edi = (CRIT_ALVO_17, "", True)
    if sel:
        crit_edi, origem_edi, assumido_edi = criterio_do_alvo(sel[0])

    col1, col2 = st.columns(2)
    with col1:
        st.caption(f"Deslocamento horizontal acumulado (mm) — criterio: {origem_edi}")
        fig = go.Figure()
        for a in sel:
            s = sub[sub[COLS["alvo"]] == a].sort_values(COLS["data"])
            fig.add_trace(go.Scatter(x=s[COLS["data"]], y=s[COLS["desl_h"]],
                                     mode="lines+markers", name=a))
        # linhas de limiar horizontal
        fig.add_hline(y=crit_edi["h_alerta"], line_dash="dash", line_color="orange",
                      annotation_text=f"Alerta {crit_edi['h_alerta']}")
        fig.add_hline(y=crit_edi["h_alarme"], line_dash="dash", line_color="red",
                      annotation_text=f"Alarme {crit_edi['h_alarme']}")
        if st.session_state.get("mostrar_obra") and len(sub):
            adicionar_fases_obra(fig, sub[COLS["data"]].min(), sub[COLS["data"]].max())
        fig.update_xaxes(title="Data")
        fig.update_yaxes(title="Desl. horizontal (mm)")
        fig.update_layout(height=460)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.caption("Assentamento vertical acumulado, ΔZ (mm) — "
                   f"alerta {crit_edi['v_alerta']} / alarme {crit_edi['v_alarme']}")
        fig2 = go.Figure()
        for a in sel:
            s = sub[sub[COLS["alvo"]] == a].sort_values(COLS["data"])
            fig2.add_trace(go.Scatter(x=s[COLS["data"]], y=s[COLS["dZ"]],
                                      mode="lines+markers", name=a))
        # limiares verticais (ΔZ pode ser negativo — desenhar ambos os sinais)
        for val, cor, txt in [(crit_edi["v_alerta"], "orange", "Alerta"),
                              (crit_edi["v_alarme"], "red", "Alarme")]:
            fig2.add_hline(y=val, line_dash="dash", line_color=cor,
                           annotation_text=f"{txt} +{val}")
            fig2.add_hline(y=-val, line_dash="dash", line_color=cor,
                           annotation_text=f"{txt} -{val}")
        if st.session_state.get("mostrar_obra") and len(sub):
            adicionar_fases_obra(fig2, sub[COLS["data"]].min(), sub[COLS["data"]].max())
        fig2.update_xaxes(title="Data")
        fig2.update_yaxes(title="ΔZ (mm)")
        fig2.update_layout(height=460)
        st.plotly_chart(fig2, use_container_width=True)

    if assumido_edi:
        st.info("Nota: este edificio usa o criterio de 17 m ASSUMIDO (nao "
                "consta explicitamente da tabela de criterios). Confirmar com "
                "o projetista antes de usar na tese como definitivo.")


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


def separador_geologia(dados):
    st.subheader("Geologia do terreno (Relatorio ENGGEO, proc. 220216)")
    st.caption("Dados do relatorio geologico-geotecnico: colunas litologicas "
               "das sondagens, ensaios SPT em profundidade e zonamento "
               "geotecnico. O terreno e, sob 0,5 m de aterro, essencialmente "
               "grés dos 'Grés Superiores' (C1As), com calcario (C1A) apenas "
               "no fundo do SC8. A deformacao nao se explica por uma camada "
               "mole — nao existe — mas pelo grau de consolidacao do grés.")

    sub1, sub2, sub3 = st.tabs(["Sondagens (litologia)", "Ensaios SPT",
                                "Zonamento geotecnico"])

    # ---- litologia lado a lado -------------------------------------------
    with sub1:
        st.caption("Colunas litologicas das quatro sondagens, em profundidade. "
                   "A linha azul marca o nivel freatico.")
        fig = go.Figure()
        sonds = list(GEO_LITOLOGIA.keys())
        for i, s in enumerate(sonds):
            desenhar_coluna_litologica(fig, s, x_centro=i, largura=0.7)
            # nivel freatico
            nf = GEO_SONDAGENS[s]["nf_prof"]
            if nf is not None:
                fig.add_shape(type="line", x0=i-0.35, x1=i+0.35, y0=nf, y1=nf,
                              line=dict(color="blue", width=2, dash="dash"))
        # legenda manual das unidades
        for unidade, cor in GEO_CORES_LITO.items():
            fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
                                     marker=dict(size=12, color=cor, symbol="square"),
                                     name=unidade))
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="lines",
                                 line=dict(color="blue", dash="dash"),
                                 name="Nivel freatico"))
        fig.update_yaxes(autorange="reversed", title="Profundidade (m)")
        fig.update_xaxes(tickmode="array", tickvals=list(range(len(sonds))),
                         ticktext=sonds, range=[-0.6, len(sonds)-0.4])
        fig.update_layout(height=600, legend_title="Unidade")
        st.plotly_chart(fig, use_container_width=True)

    # ---- SPT --------------------------------------------------------------
    with sub2:
        st.caption("Ensaios SPT (N = numero de pancadas) em profundidade. "
                   "N=60 corresponde a nega. Valores altos = terreno mais "
                   "resistente. A variacao dentro do grés reflete o grau de "
                   "consolidacao (zonas ZG5 a ZG1).")
        fig2 = go.Figure()
        for s, ensaios in GEO_SPT.items():
            profs = [e[0] for e in ensaios]
            ns = [e[1] for e in ensaios]
            fig2.add_trace(go.Scatter(x=ns, y=profs, mode="lines+markers", name=s))
        fig2.update_yaxes(autorange="reversed", title="Profundidade (m)")
        fig2.update_xaxes(title="N (pancadas)", range=[0, 65])
        fig2.add_vline(x=60, line_dash="dot", line_color="gray",
                       annotation_text="Nega (60)")
        fig2.update_layout(height=600, legend_title="Sondagem")
        st.plotly_chart(fig2, use_container_width=True)

    # ---- zonamento --------------------------------------------------------
    with sub3:
        st.caption("Zonamento geotecnico e parametros propostos (Quadros V-VII "
                   "do relatorio). Estes sao os parametros que alimentam a "
                   "modelacao numerica da contencao.")
        st.dataframe(pd.DataFrame(GEO_ZONAMENTO), use_container_width=True,
                     hide_index=True)
        st.caption("gama: peso volumico | c': coesao | fi': angulo de atrito | "
                   "E': modulo de deformabilidade. Zonas ZG3-ZG1 (rocha) com c' "
                   "e E' em MPa/GPa; ZG6-ZG4 (solo/grés brando) em kPa/MPa.")


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
# SEPARADOR 9 — CORRELACAO (deslocamento + agua + fases, mesmo eixo temporal)
# =========================================================================
def separador_correlacao(dados):
    st.subheader("Correlacao temporal — deslocamento, agua e fases da obra")
    st.caption("Junta num so eixo temporal o deslocamento de um inclinometro, "
               "a cota da agua do piezometro e as fases da obra. E aqui que se "
               "ve se a aceleracao dos deslocamentos e a descida da agua "
               "acompanham o avanco da escavacao.")

    resumo = dados["resumo"]
    pz = dados["piezo"]
    ok = validar_colunas(resumo, [COLS["data"], COLS["inclinometro"],
                                  COLS["desl_max_global"]], "Inclinometros") \
        and validar_colunas(pz, [COLS["data"], COLS["piezometro"],
                                 COLS["cota_agua"]], "Piezometros")
    if not ok:
        return

    col1, col2 = st.columns(2)
    with col1:
        inc = st.selectbox("Inclinometro",
                           sorted(resumo[COLS["inclinometro"]].dropna().unique()),
                           key="corr_inc")
    with col2:
        piez = st.selectbox("Piezometro",
                            sorted(pz[COLS["piezometro"]].dropna().unique()),
                            key="corr_piez")

    s_inc = resumo[resumo[COLS["inclinometro"]] == inc].sort_values(COLS["data"])
    s_pz = pz[pz[COLS["piezometro"]] == piez].sort_values(COLS["data"])

    # eixo Y duplo: deslocamento (esq) e cota da agua (dir)
    from plotly.subplots import make_subplots
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Scatter(
        x=s_inc[COLS["data"]], y=s_inc[COLS["desl_max_global"]],
        mode="lines+markers", name=f"Desloc. max {inc} (mm)",
        line=dict(color="#d62728", width=2)), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=s_pz[COLS["data"]], y=s_pz[COLS["cota_agua"]],
        mode="lines+markers", name=f"Cota agua {piez} (m)",
        line=dict(color="#1f77b4", width=2, dash="dot")), secondary_y=True)

    # fases da obra por baixo
    todas_datas = pd.concat([s_inc[COLS["data"]], s_pz[COLS["data"]]])
    if len(todas_datas):
        adicionar_fases_obra(fig, todas_datas.min(), todas_datas.max())

    fig.update_xaxes(title="Data")
    fig.update_yaxes(title_text="Deslocamento maximo (mm)", secondary_y=False)
    fig.update_yaxes(title_text="Cota da agua (m)", secondary_y=True)
    fig.update_layout(height=560, legend=dict(orientation="h", y=1.05))
    st.plotly_chart(fig, use_container_width=True)

    st.caption("Leitura: se a linha vermelha (deslocamento) sobe enquanto a "
               "azul (agua) desce, ambas durante a fase de escavacao, isso "
               "sugere que a escavacao esta a induzir tanto o movimento da "
               "contencao como o rebaixamento do nivel freatico. Datas de fase "
               "sao as previstas no plano de trabalhos.")


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

    thome, t3d, tinc, talv, tcc, tpz, tcorr, tgeo, tobra, tplan = st.tabs(
        ["Inicio", "Visao geral 3D", "Inclinometros", "Alvos (2D)",
         "Celulas de carga", "Piezometros", "Correlacao", "Geologia", "Obra",
         "Planta (DXF)"])
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
    with tcorr:
        separador_correlacao(dados)
    with tgeo:
        separador_geologia(dados)
    with tobra:
        separador_obra(dados)
    with tplan:
        separador_planta(dados)


if __name__ == "__main__":
    main()
