"""
=============================================================================
 THE EDEN, ESTORIL — APLICAÇÃO DE ANÁLISE DE INSTRUMENTAÇÃO GEOTÉCNICA
 Dissertação de Mestrado — back-analysis da contenção periférica
=============================================================================

 O QUE E ESTE FICHEIRO
 ---------------------
 Aplicação Streamlit. NÃO se corre colando no Python normal nem num
 "Python online": o Streamlit le este ficheiro e transforma-o numa página
 web. Ver "COMO CORRER".

 SEPARADORES
 -----------
   1. Visão geral 3D  — alvos topográficos no espaço real (M,P,Z), com
                        deslocamento amplificado, cor por magnitude e
                        evolução por campanha. Superfície interpolada
                        OPCIONAL, sempre com os pontos medidos por cima.
   2. Inclinómetros   — perfil deformado, evolução, velocidade, precursores.
   3. Alvos (2D)      — séries temporais de deslocamento horizontal/vertical.
   4. Células de carga— carga vs. blocagem e variação (%) com alerta/alarme.
   5. Piezómetros     — cota da água.

 NOTA METODOLÓGICA (importante para a defesa)
 --------------------------------------------
 - Os ALVOS TOPOGRÁFICOS tem coordenadas reais (M,P,Z) -> podem ir para 3D.
 - Os INCLINÓMETROS não tem coordenadas em planta neste ficheiro (só azimute
   e profundidade), por isso NÃO são colocados no 3D dos alvos — seria
   inventar a posição. Ficam na sua vista própria. Se um dia tiveres as
   coordenadas em planta, integram-se com vetores orientados pelo azimute.
 - A superfície 3D e INTERPOLAÇÃO entre alvos: mostra-se só como apoio
   visual e com os pontos medidos sempre visíveis. E uma leitura do campo
   de deslocamentos dos alvos, não um modelo do maciço.

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
 (M/P ~5000). Se não estiver, há um ajuste manual de posição no separador.
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
    "ancoragens": "Ancoragens",      # registos de execução (opcional)
    "fissuro": "Fissurometros",      # fissurómetros do Edifício 1 (opcional)
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
# EDIFICIOS ADJACENTES (Edifício 1, Edifício 2, Edifício 3): H 15/25, V 10/20.
#   -> Capitulo 4 dos relatorios de alvos, desde o R00 (12/05/2025); os
#      relatorios nao indicam a origem destes valores.
# CONTENCAO 17 m (alcados AB, CD, DE, BF, PQ): H 20/40, V 10/15.
# CONTENCAO 24 m (alcados FG, GH, JK, KL, LM, MNO, OP): H 30/40, V 10/15.
#   -> Valores da memoria descritiva (secao 14.6); atribuicao a cada alcado
#      segundo o capitulo 4 dos relatorios de alvos, desde o R06 (06/10/2025).
#      Os alcados DE (17 m) e LM (24 m) so constam a partir do R36
#      (16/03/2026). Valores inalterados do R00 ao R67.
# NOTA: MNO e OP tem 17 a 20 m de cortina, mas os relatorios classificam-nos
#   como 24 m; com 17 m entrariam em alerta mais cedo (fev. 2026).
# =========================================================================
CRIT_VIZINHOS   = (15, 25, 10, 20)   # edificios adjacentes (OFICIAL)
CRIT_CONT_17    = (20, 40, 10, 15)   # contencao 17 m
CRIT_CONT_24    = (30, 40, 10, 15)   # contencao 24 m

# -------------------------------------------------------------------------
# ESTIMATIVA DE DEFORMACAO DE PROJETO (memoria descritiva JETsj)
# -------------------------------------------------------------------------
# Valor de CALCULO (nao limite): "estima-se para a ultima fase de escavacao
# uma deformacao horizontal acumulada de cerca de 20mm na cortina poente/norte
# e 10mm na cortina nascente/sul" (corresponde a deformacao relativa > H/500).
# Usado para a analise OBSERVADO vs. PREVISTO. E uma estimativa de cortina —
# a associacao de cada grupo de alvos a orientacao e INFERIDA das coordenadas
# (a confirmar com a planta do projeto).
DEFORM_PROJETO = {
    "poente_norte": 20.0,   # mm (cortina poente/norte)
    "nascente_sul": 10.0,   # mm (cortina nascente/sul)
}

# A estimativa de projeto e atribuida pela FRENTE DE ESCAVACAO adjacente,
# a mesma usada nas Correlacoes (zona_escav_de_edificio / ALCADO_ZONA):
#   frente poente (alcados CD, BF, DE, EF, FG, GH; Edifício 1) -> 20 mm
#   frente norte  (alcados HI-KL; Edifício 3)                      -> 20 mm
#   frente nascente/sul (alcados AB, LM-QR; Edifício 2)              -> 10 mm
# Assim, Edifício 2 (frente nascente, alcado PQ a ~8 m) compara com 10 mm.
# NOTA: para edificios vizinhos a comparacao e INDIRETA — a estimativa de
# projeto refere-se ao deslocamento da cortina, nao ao do edificio.
PREVISTO_POR_FRENTE = {"poente_sul": 20.0, "norte": 20.0, "nascente_sul": 10.0}

# ZONA DO PROJETO de cada alcado (memoria descritiva: ~20 mm na cortina
# poente/norte e ~10 mm na nascente/sul; figuras de calculo: ~16 mm a norte e
# a poente e ~4 mm a sul). Atribuicao INFERIDA da posicao de cada alcado (a
# memoria nao lista os alcados de cada zona); os do canto sudoeste (BF, CD,
# DE) sao os de atribuicao menos segura. Difere das FRENTES DE ESCAVACAO do
# plano (ALCADO_ZONA), que servem as Correlacoes: o LM, por exemplo, e da
# frente de escavacao nascente/sul, mas fica no lado norte da cortina.
ZONA_PROJETO_ALCADO = {
    "BF": "poente_norte", "CD": "poente_norte", "DE": "poente_norte",
    "EF": "poente_norte", "FG": "poente_norte", "GH": "poente_norte",
    "HI": "poente_norte", "IJ": "poente_norte", "JK": "poente_norte",
    "KL": "poente_norte", "LM": "poente_norte",
    "AB": "nascente_sul", "MN": "nascente_sul", "MNO": "nascente_sul",
    "OP": "nascente_sul", "PQ": "nascente_sul", "QR": "nascente_sul",
}
NOME_ZONA_PROJETO = {"poente_norte": "poente/norte", "nascente_sul": "nascente/sul"}


def previsto_do_grupo(grupo):
    """(estimativa do projeto em mm, zona, tipo de comparação) de um grupo de
    alvos. Alçados: zona do projeto (comparação direta). Edifícios vizinhos:
    frente de escavação adjacente (comparação indireta, a estimativa é da
    cortina). (None, None, None) se não houver correspondência."""
    alc = _extrair_alcado(grupo)
    if alc:
        z = ZONA_PROJETO_ALCADO.get(alc)
        if z is None:
            return None, None, None
        return DEFORM_PROJETO[z], NOME_ZONA_PROJETO[z], "direta"
    fr = zona_escav_de_edificio(grupo)
    if fr in PREVISTO_POR_FRENTE:
        return PREVISTO_POR_FRENTE[fr], NOME_FRENTE[fr], "indireta"
    return None, None, None


def data_ultrapassagem(serie_datas, serie_valores, limiar):
    """Primeira data em que o máximo dos alvos (por campanha) atinge o limiar."""
    d = pd.DataFrame({"t": pd.to_datetime(serie_datas), "v": serie_valores}).dropna()
    if d.empty:
        return None, None
    m = d.groupby("t")["v"].max().sort_index()
    acima = m[m >= limiar]
    if acima.empty:
        return None, None
    return acima.index[0], float(acima.iloc[0])


def nome_curto_grupo(g):
    alc = _extrair_alcado(g)
    return f"Alçado {alc}" if alc else str(g)
NOME_FRENTE = {"poente_sul": "poente", "norte": "norte",
               "nascente_sul": "nascente/sul"}

# Localizacao das CELULAS DE CARGA (alcado/zona onde esta a ancoragem).
# So faz sentido cruzar a carga de uma celula com o movimento de alvos do
# MESMO alcado/zona. CONFIRMADO em planta e modelo 3D do projeto: ambas as
# celulas estao na cortina sob o Edifício 1 (lado poente,
# junto a ancoragem A26). Nota: o alcado da 33GRADOS assinala "ZG2" nesta zona,
# mas as sondagens SC6/Pz e SC8/Pz e os perfis interpretativos da ENGGEO nao o
# confirmam — por isso a zona geotecnica NAO e indicada na app. Por isso cruzam validamente com os alvos do Edifício 1
# (A1-A8...). Cada entrada: (rotulo da zona, filtro de grupo de alvos).
#   - CC 2501796 / A26: cortina do Edifício 1 (Piso -1, cota 15,90)
#   - CC 200792  / DE : cortina do Edifício 1 (proxima da A26)
LOCALIZACAO_CELULAS = {
    "CC 2501796": ("Cortina sob o Edifício 1 — Piso -1 (cota 15,90), junto a A26", "Edifício 1"),
    "CC 200792":  ("Cortina sob o Edifício 1 — Piso -2 (cota 12,45), perto da A26", "Edifício 1"),
}

# Imagens de localizacao de cada celula (planta/alcado do projeto + contexto).
# Ficheiros em fotos_celulas/ (ao lado do script). Cada entrada: lista de
# (ficheiro, legenda). A "planta_geral" e partilhada como contexto do recinto.
IMAGENS_CELULAS = {
    "CC 2501796": [
        ("cc_2501796_planta.jpg",
         "Planta de localização das células — A26 (Piso -1, cota 15,90)"),
        ("cc_2501796_3d.jpg",
         "Modelo 3D do projeto — localização da célula A26; Edifício 1 assinalado"),
    ],
    "CC 200792": [
        ("cc_200792_alcado.jpg",
         "Alçado D-E — ancoragem DE (Piso -2, cota 12,45)"),
        ("planta_geral.jpg",
         "Planta de localização geral (recinto) — Edifício 1 assinalado"),
    ],
}

ALCADOS_24M = {"FG", "GH", "JK", "KL", "LM", "MNO", "OP"}   # relatorios de alvos (LM desde o R36)
ALCADOS_17M = {"AB", "CD", "DE", "BF", "PQ"}               # relatorios de alvos (DE desde o R36)

# -------------------------------------------------------------------------
# ALTURA DA CORTINA H POR ALCADO (m) — Tabela 3.7 da dissertacao (topo das
# estacas ao fundo de escavacao, 4,55 m; sem muros/taludes acima do
# coroamento). Alguns alcados tem intervalo: guarda-se (Hmin, Hmax) e usa-se
# o ponto medio na representacao. Fonte: projeto JETsj [52].
ALTURA_CORTINA = {
    "AB": (16.3, 16.3), "BF": (16.3, 16.7), "CD": (16.3, 16.7),
    "DE": (16.3, 16.7), "FG": (21.1, 22.0), "GH": (20.1, 20.1),
    "JK": (20.1, 23.6), "KL": (22.35, 22.35), "LM": (20.1, 22.35),
    "MNO": (17.05, 20.1), "OP": (17.05, 17.05), "PQ": (17.05, 17.05),
}
# Altura da cortina no alcado onde se situa cada inclinometro (Tabela 5.2).
H_INCLINOMETRO = {"I1": (16.3, 16.3), "I2": (20.1, 22.35), "I3": (16.3, 16.3)}

# Referenciais empiricos do deslocamento horizontal maximo, em % de H
# (Tabela 2.1): (rotulo, min%, max%, cor RGB). Clough & O'Rourke e um valor
# medio (~0,2 %). Fontes: Clough & O'Rourke [21], Long [24], Yoo [26].
REFERENCIAS_EMPIRICAS = [
    ("Long (2001) — solos rijos (0,05–0,25 %)", 0.05, 0.25, "31,119,180"),
    ("Yoo (2001) — terreno sobre rocha (0,05–0,15 %)", 0.05, 0.15, "44,160,44"),
    ("Clough & O'Rourke — média ≈ 0,2 %", 0.20, 0.20, "214,39,40"),
]


def _extrair_alcado(edif):
    """Devolve o código do alçado (ex. 'FG') ou None se não for contenção."""
    import re
    if isinstance(edif, str) and "Alçado" in edif:
        m = re.search(r"Alçado (\w+)", edif)
        return m.group(1) if m else None
    return None


def criterios_do_alvo(edif):
    """
    Devolve (critério, rotulo, a_confirmar) para uma linha de alvo, a partir
    do nome do edifício/elemento. 'critério' e o tuplo (Há,Hm,Va,Vm).
    """
    alc = _extrair_alcado(edif)
    if alc is None:
        return CRIT_VIZINHOS, "Edifício adjacente (15/25 · 10/20)", False
    if alc in ALCADOS_24M:
        return CRIT_CONT_24, f"Contenção 24 m — Alçado {alc} (30/40 · 10/15)", False
    if alc in ALCADOS_17M:
        return CRIT_CONT_17, f"Contenção 17 m — Alçado {alc} (20/40 · 10/15)", False
    # alcado sem classificacao nos relatorios (nao ocorre nos dados atuais)
    return CRIT_CONT_17, f"Alçado {alc} (17 m assumido — não consta dos relatórios)", True


def estado_calculado(h, v, criterio):
    """
    Estado a partir do deslocamento horizontal (h) e vertical (v) acumulados,
    dado um critério (Há,Hm,Va,Vm). O estado e o mais severo entre H e V.
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
      'Critério'          — rotulo legível do critério aplicado
      'Estado calculado'  — estado recalculado de ΔH/ΔV com os critérios oficiais
      'Confere'           — True se coincide com a coluna 'Estado' do Excel
      'Fachada'        — 'Frente escavação' | 'Lateral (mar)' | '' (só Edifício 1)
    Não altera a coluna 'Estado' original: serve de auditoria lado a lado.
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
        fachadas.append(fachada_edificio1(r.get(COLS["edificio"]), r.get(COLS["alvo"])))
    d["Critério"] = rotulos
    d["Estado calculado"] = estados
    d["Confere"] = confere
    d["Fachada"] = fachadas
    return d


# =========================================================================
# EDIFÍCIO 1 — DUAS FACHADAS E SUBSTITUICAO DE ALVOS
# -------------------------------------------------------------------------
# O Edifício 1 tem duas fachadas instrumentadas:
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


def fachada_edificio1(edif, alvo):
    """Devolve a fachada do Edifício 1 a que o alvo pertence, ou '' se não aplicar."""
    if not (isinstance(edif, str) and "Edifício 1" in edif):
        return ""
    a = str(alvo)
    if a in SC_FACHADA_1:
        return "Frente escavação"
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
    "SC6/Pz": [(0.0, 0.5, "Aterro"), (0.5, 21.05, "Grés (C1As)")],
    "SC7":    [(0.0, 0.5, "Aterro"), (0.5, 19.89, "Grés (C1As)")],
    "SC8/Pz": [(0.0, 0.5, "Aterro"), (0.5, 19.5, "Grés (C1As)"),
               (19.5, 21.0, "Calcário (C1A)")],
    "SC9/Pz": [(0.0, 0.5, "Aterro"), (0.5, 21.41, "Grés (C1As)")],
}

# cor de cada unidade litologica (para a coluna)
GEO_CORES_LITO = {
    "Aterro": "#c0641e",          # mesmo laranja do ZG6 (aterro)
    "Grés (C1As)": "#a9c47f",     # verde-base do gres (alinhado ao ZG4)
    "Calcário (C1A)": "#5c8a45",  # verde mais escuro (calcario de fundo)
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
    {"Zona":"ZG6","Descrição":"Aterro heterogéneo de origem não seletiva",
     "gama (kN/m3)":"17-18","c' (kPa)":"<5","fi' (graus)":"<26","E'":"<5 MPa"},
    {"Zona":"ZG5","Descrição":"Grés pouco consolidado, SPT 11-30, RQD 0%",
     "gama (kN/m3)":"19-20","c' (kPa)":"5-15","fi' (graus)":"28-33","E'":"8-30 MPa"},
    {"Zona":"ZG4","Descrição":"Grés pouco consolidado, SPT 31-56, RQD 0%",
     "gama (kN/m3)":"20-21","c' (kPa)":"5-30","fi' (graus)":"30-36","E'":"25-50 MPa"},
    {"Zona":"ZG3","Descrição":"Grés/calcário irreg. consolidado, SPT>=60, RQD 0-25%",
     "gama (kN/m3)":"22-24","c' (kPa)":"0.04-0.40 MPa","fi' (graus)":"29-32","E'":"0.07-0.22 GPa"},
    {"Zona":"ZG2","Descrição":"Grés irreg. consolidado a consolidado, SPT>=60, RQD 45-75%",
     "gama (kN/m3)":"24-25","c' (kPa)":"1.26-3.14 MPa","fi' (graus)":"32-36","E'":"0.80-2.81 GPa"},
    {"Zona":"ZG1","Descrição":"Grés consolidado, SPT>=60, RQD 76-100%",
     "gama (kN/m3)":"25-26","c' (kPa)":"2.70-8.69 MPa","fi' (graus)":"38-42","E'":"7.94-14.13 GPa"},
]

# =========================================================================
# COTAS DE PROJETO — ESCAVACAO E CONTENCAO (projeto JETsj, PRO/2023/368,
# EDN-JET-...-0001 rev.D). Cotas absolutas em metros, referidas aos toscos.
# Transcritas dos cortes (desenhos 0021-0024) e confirmadas por repeticao
# em varios cortes. As cotas dos pisos sao consistentes ao longo da obra;
# o coroamento da cortina varia por alcado (dominantes 20,85 e 24,65).
# =========================================================================
COTAS_PISOS = [
    ("Piso 2 (coroamento zona alta)", 24.65),
    ("Piso 1",  20.85),
    ("Piso -1", 15.90),
    ("Piso -2", 12.45),
    ("Piso -3",  9.00),
    ("Piso -4",  5.55),
]
COTA_FUNDO_ESCAVACAO = 4.55          # cota final de escavacao (dominante nos cortes)
COTA_COROAMENTO_PADRAO = 20.85       # coroamento da cortina (alcados correntes)
COTA_COROAMENTO_ALTA = 24.65         # coroamento na zona alta (piso 2)
COTA_MURO_EDIF1 = 22.50               # muro tradicional na fronteira com o Edifício 1

# Cores oficiais das bandas de laje por espessura, lidas da legenda das pecas
# desenhadas do projeto de contencao (JETsj, EDN-JET-...-DR-U-0021 e seg.).
# Usadas para dar as bandas de laje esquematicas as cores reais do projeto.
CORES_LAJE_PROJETO = {
    0.25: "#d9d9d9",   # cinza claro
    0.30: "#1a7a1a",   # verde escuro
    0.35: "#c8721a",   # laranja/castanho
    0.38: "#22dd22",   # verde vivo
    0.45: "#1f78d1",   # azul
    0.60: "#7a1fa0",   # roxo
    0.75: "#e020c0",   # magenta
}

# Nivel freatico de REPOUSO medido nos piezometros das sondagens
# (ENGGEO, Quadro III, leitura de 24/11/2022). Cota da agua, em metros.
NF_REPOUSO = [
    ("SC6/Pz", 16.1),
    ("SC8/Pz", 19.3),
    ("SC9/Pz", 16.3),
]

# =========================================================================
# INCLINOMETROS — metadados e associacao a sondagem
# -------------------------------------------------------------------------
# Profundidade e azimute do eixo A+ vem da folha Instrumentos do Excel.
# A sondagem "mais proxima" NAO consta dos dados (os inclinometros nao tem
# coordenadas no Excel); foi inferida por SOBREPOSICAO das duas plantas —
# a da prospecao (relatorio ENGGEO) e a dos inclinometros (relatorio de
# instrumentacao). E uma associacao SUGERIDA por proximidade, A CONFIRMAR
# com a equipa de instrumentacao. O nivel de confianca reflete a clareza
# da correspondencia visual entre as plantas.
INC_META = {
    "I1": {"sondagem": "SC6/Pz", "confianca": "média",
           "azimute": 330, "posicao": "canto SO (poente)"},
    # I1: a SC8/Pz fica no canto NE do Edifício 1, longe do canto SO; a
    # sondagem de 2022 mais proxima do I1 e a SC6/Pz (as sondagens da 2.a fase,
    # incluindo a SC3/Pz junto ao canto SO, nao estavam executadas em 2022).
    "I2": {"sondagem": "SC9/Pz", "confianca": "alta",
           "azimute": 225, "posicao": "topo N (bolbo curvo)"},
    "I3": {"sondagem": "SC6/Pz", "confianca": "média",
           "azimute": 335, "posicao": "SE/nascente"},
}


def _rumo_cardeal(az):
    """Converte azimute (graus) em rumo cardeal aproximado, para leitura."""
    dirs = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
            "S", "SSO", "SO", "OSO", "O", "ONO", "NO", "NNO"]
    return dirs[int((az % 360) / 22.5 + 0.5) % 16]

# =========================================================================
# CRONOGRAMA DA OBRA  (Plano de Trabalhos Alves Ribeiro/HCI, 05/05/2025)
# Datas PREVISTAS transcritas do PDF do plano. Sao o planeado, nao o real.
# =========================================================================
# (nome, inicio ISO, conclusao ISO)
# =========================================================================
# FASEAMENTO DA OBRA — datas do PLANO DE TRABALHOS IMPACTADO (22/04/2026; data
# de estado 21/08/2026). O plano impactado e o contratual com tres suspensoes
# de frentes inseridas durante as estacas; as duracoes sao as contratuais.
# Nao e um registo independente da execucao. Fonte: "Plano de Trabalhos -
# Impactado ECP" (o mesmo ficheiro MS Project do "planeamento detalhado").
# =========================================================================
FASES_OBRA = [
    ("Contenção periférica",                    "2025-05-13", "2026-04-15"),
    ("Estacas Poente e Norte",                  "2025-05-13", "2025-08-12"),
    ("Estacas Central e Nascente",              "2025-08-13", "2025-10-07"),
    ("Estacas junto ao Edifício 2",                           "2025-05-13", "2025-09-25"),
    ("Escavação + bandas de laje + ancoragens", "2025-07-08", "2026-04-15"),
    ("Viga coroamento + muros contenção",       "2025-07-08", "2025-11-13"),
    ("Fundação e laje de fundo",                "2026-03-10", "2026-05-04"),
    ("Estrutura pisos -3 a -1",                 "2026-04-14", "2026-06-29"),
    ("Estrutura pisos 1 e 2 (em curso)",        "2026-06-16", "2026-08-24"),
]

# ESCAVACAO POR COTA — marcos reais (cota atingida + datas), do planeamento
# impactado. Estes marcos permitem cruzar QUANDO a escavacao chegou a cada
# cota com o que a instrumentacao mediu nessas datas. As cotas estao no
# referencial de PROJETO (o mesmo das cotas dos pisos), nao no dos alvos.
# (rotulo, cota_final_m, data_inicio, data_fim)
ESCAVACAO_COTAS = [
    ("cota 26,50 → 23,75",              23.75, "2025-09-04", "2025-09-05"),
    ("cota 29 → 22,55 (fundo Anel P2)", 22.55, "2025-09-15", "2025-09-17"),
    ("cota 22,55 → 18,90",              18.90, "2025-09-17", "2025-09-22"),
    ("até cota 21,45",                  21.45, "2025-09-22", "2025-09-22"),
    ("até cota 19,80/19,00 (fundo VD Piso 1)", 19.00, "2025-10-22", "2025-10-24"),
    ("até cota VD Piso -1",             15.90, "2025-11-13", "2025-11-26"),
    ("de 4,95 m → cota 14,85",          14.85, "2025-11-12", "2025-11-25"),
    ("até cota VD Piso -2",             12.45, "2025-12-26", "2026-01-06"),
    ("até cota VD Piso -3",              9.00, "2026-02-10", "2026-02-25"),
]

# =========================================================================
# ESCAVACAO POR ZONA (frente de estacas) — do PLANEAMENTO DETALHADO (MS
# Project "Planeamento detalhado Eden", ramo "Escavação, bandas de laje,
# paredes de forro e ancoragens"). Cada zona e uma FRENTE da cortina,
# identificada no plano pelas estacas e mapeada aos edificios/alvos que lhe
# ficam em frente (confirmado com a foto aerea e a planta de estacas):
#   poente_sul  = cortina poente, estacas E8_35–E8_96 + anel E6_C  -> Edifício 1
#   norte       = cortina norte, alinhamentos IJ/JK/KL, E8_92–E6_11 -> Edifício 3
#   nascente_sul= cortina nascente/sul, E8_2–E8_34 + E6_16–E6_21    -> Edifício 2
# Datas REAIS (impactadas) transcritas do plano — nada inventado. Cada evento:
#   (data_iso, tipo, cota_atingida_m|None, rotulo)
#   tipo: "escavacao" (cota atingida) | "ancoragem" | "banda_laje"
# =========================================================================
ESCAVACAO_ZONAS = {
    "poente_sul": [
        ("2025-07-25", "escora",    None,  "Escoras metálicas F-G/G-H (viga de coroamento)"),
        ("2025-09-05", "escavacao", 23.75, "cota 26,50 -> 23,75"),
        ("2025-09-09", "ancoragem", None,  "Ancoragens viga de coroamento Piso 2 (E8_78-E8_90)"),
        ("2025-10-02", "escora",    None,  "Escoras metálicas de canto (alinhamentos B/C/D/E)"),
        ("2025-10-06", "ancoragem", None,  "Ancoragens VD Piso 1 (E8_78-E8_87)"),
        ("2025-11-11", "ancoragem", None,  "Ancoragens viga de coroamento Piso 1 (E8_51-E8_35)"),
        ("2025-11-25", "escavacao", 14.85, "cota 14,85 (-4,95 m) + VD da grua (2 ancoragens)"),
        ("2025-12-25", "ancoragem", None,  "Ancoragens VD Piso -1 (2 lados)"),
        ("2026-01-06", "escavacao", 12.45, "VD Piso -2"),
        ("2026-01-23", "escora",    None,  "Escoras metálicas de canto (Piso -2)"),
        ("2026-02-09", "ancoragem", None,  "Ancoragens VD Piso -2 (2 lados)"),
        ("2026-02-25", "escavacao", 9.00,  "VD Piso -3 + VD da grua (2 ancoragens)"),
        ("2026-03-23", "escavacao", 4.55,  "fundo de escavação"),
    ],
    "norte": [
        ("2025-07-28", "escora",    None,  "Escoras metálicas na VD H-I/I-J"),
        ("2025-09-17", "escavacao", 22.55, "cota 29 -> 22,55"),
        ("2025-10-14", "escora",    None,  "Escoras metálicas na banda de laje do Piso 2"),
        ("2025-09-22", "escavacao", 18.90, "cota 22,55 -> 18,90"),
        ("2025-10-24", "escavacao", 19.00, "fundo VD Piso 1 (19,80/19,00)"),
        ("2025-11-04", "banda_laje", None, "Banda Laje Piso 1 (Fase 2)"),
        ("2025-12-31", "escavacao", 12.45, "VD Piso -2 + fundo Banda Laje"),
        ("2026-02-12", "escora",    None,  "Travessa metálica (escoramento) na banda do Piso -2"),
        ("2026-02-13", "banda_laje", None, "Banda Laje Piso -2 (F2)"),
        ("2026-04-15", "escavacao", 4.55,  "fundo de escavação"),
    ],
    "nascente_sul": [
        ("2025-10-27", "banda_laje", None, "Banda Laje Piso 1 (Fase 1)"),
        ("2025-11-26", "escavacao", 15.90, "VD Piso -1"),
        ("2025-12-23", "escavacao", 12.45, "VD Piso -2 + fundo Banda Laje"),
        ("2026-02-04", "banda_laje", None, "Banda Laje Piso -2 (F1)"),
        ("2026-03-11", "escavacao", 9.00,  "VD Piso -3"),
        ("2026-04-14", "escavacao", 4.55,  "fundo de escavação"),
    ],
}
import copy as _copy
ESCAVACAO_ZONAS_PLANO = _copy.deepcopy(ESCAVACAO_ZONAS)   # base (plano impactado)

# Datas do PLANO para cada lote de ancoragens (execução e tensionamento), para a
# comparação com os registos de execução (folha «Ancoragens»).
ANC_PLANO = {
    "A37–A42": ("28/08–09/09/2025 (tarefa única)", "não discriminado"),
    "A19–A24": ("22–26/09/2025", "06/10/2025"),
    "A25–A30": ("09–16/12/2025 (VD, 2 lados)", "24–25/12/2025"),
    "A31–A36": ("22–29/01/2026 (VD, 2 lados)", "06–09/02/2026"),
    "A1–A6": ("30/10–03/11/2025 (coroamento E8_51–E8_35)?", "11/11/2025?"),
    "A7–A12": ("09–16/12/2025 (VD, 2 lados)?", "24–25/12/2025?"),
    "A13–A18": ("22–29/01/2026 (VD, 2 lados)?", "06–09/02/2026?"),
    "A7G, A8G": ("com a escavação de 12–25/11/2025", "—"),
    "A3G, A6G": ("com a escavação de 10–25/02/2026", "—"),
}

# Ocorrências e medidas indicadas pela obra (ponto 4). Datas None = a confirmar.
MEDIDAS_OBRA = [
    ("2025-06-02", "Ancoragens", "Injeção repetitiva (IR) em vez da IRS do projeto, com o acordo do projetista", "Confirmada"),
    ("2026-01-19", "Edifício 1", "Fissurómetros em cinco fissuras do edifício (zeragem)", "Confirmada"),
    ("2026-01-26", "Monitorização", "Leituras dos alvos duas vezes por semana", "Confirmada nos dados"),
    ("2026-02-13", "Alçado GH", "Água abundante numa ancoragem junto à célula da A26; recomendadas furação de drenagem entre estacas e tirante preventivo", "Confirmada"),
    ("2026-02-16", "Alçado GH / A26", "Avaliação da carga da A26 pelo projetista; estudado um escoramento «pé de galinha» entre AB e GH (não executado)", "Confirmada"),
    ("2026-02-16", "Monitorização", "Leituras das células de carga duas vezes por semana", "Confirmada nos dados"),
    (None, "Alçado GH", "Tirantes nos muros sobre as vigas de coroamento", "Data a confirmar"),
    (None, "Alçado GH", "Um nível de carotes Ø80 mm ao nível do piso 1, entre estacas, para aliviar a pressão da água", "Data a confirmar"),
    (None, "Alçado GH", "Dois perfis HEB300 contra a viga de distribuição, apoiados na sapata do Núcleo 1 (após o fundo)", "Data a confirmar"),
    (None, "Edifício 1", "Leituras por cordas no edifício", "Data a confirmar"),
    (None, "Bandas de laje", "Escavação para as bandas dos pisos 2 e 1 até 4,0 m (previstos 2,5 m)", "Data a confirmar"),
    (None, "Bandas de laje", "Banda do piso −2 contra o terreno só no alçado do Edifício 1; nas restantes, escavação e escoramento", "Alçado e data a confirmar"),
    (None, "Anel (LM)", "Supressão de um dos perfis de ligação do anel de compressão às fundações", "Data a confirmar"),
]


def _col_nivel(df):
    for c in ("Alçado e nível", "Nível provável (inferido)"):
        if c in df.columns:
            return c
    return "Lote"


def aplicar_registos_ancoragens(dados):
    """Substitui, na frente poente/sul, os marcos de ancoragens do plano pelos
    tensionamentos registados (folha «Ancoragens»), quando a folha existe.
    Dentro de cada lote, datas separadas por até 2 dias úteis formam um só
    marco, colocado no último dia (tensionamento concluído)."""
    base = _copy.deepcopy(ESCAVACAO_ZONAS_PLANO)
    anc = dados.get("ancoragens")
    col_t = "Tensionamento e ensaio (adotado)"
    if anc is not None and not anc.empty and col_t in anc.columns:
        a = anc.copy()
        a[col_t] = pd.to_datetime(a[col_t], errors="coerce")
        a = a.dropna(subset=[col_t])
        ev = []
        for lote, g in a.groupby("Lote", sort=False):
            nivel = str(g[_col_nivel(g)].iloc[0])
            datas = sorted(g[col_t].unique())
            grupos, cur = [], [datas[0]]
            for d in datas[1:]:
                gap = np.busday_count(pd.Timestamp(cur[-1]).date(), pd.Timestamp(d).date())
                if gap <= 2:
                    cur.append(d)
                else:
                    grupos.append(cur); cur = [d]
            grupos.append(cur)
            for gr in grupos:
                t0, t1 = pd.Timestamp(gr[0]), pd.Timestamp(gr[-1])
                nomes = g.loc[g[col_t].isin(gr), "Ancoragem"].tolist()
                nomes = lote if len(nomes) == len(g) else ", ".join(nomes)
                per = (f"{t1:%d/%m/%Y}" if t0 == t1 else f"{t0:%d/%m} a {t1:%d/%m/%Y}")
                ev.append((f"{t1:%Y-%m-%d}", "ancoragem", None,
                           f"Ancoragens {nomes} — {nivel}: tensionamento {per} (registo)"))
        zs = []
        for d, tp, cota, rot in base["poente_sul"]:
            if tp == "ancoragem":
                continue
            if tp == "escavacao":
                rot = rot.replace(" + VD da grua (2 ancoragens)", "")
            zs.append((d, tp, cota, rot))
        base["poente_sul"] = sorted(zs + ev, key=lambda e: e[0])
    ESCAVACAO_ZONAS.clear(); ESCAVACAO_ZONAS.update(base)


ZONA_ESCAV_NOME = {
    "poente_sul":   "Poente/Sul — cortina do Edifício 1",
    "norte":        "Norte — cortina do Edifício 3",
    "nascente_sul": "Nascente/Sul — cortina do Edifício 2",
}

# Associacao INSTRUMENTACAO -> frente de escavacao (a mesma zona).
#   Inclinometros: I1 canto SO (poente) / I2 topo N / I3 SE (nascente).
#   Celulas: ambas na cortina do Edifício 1 (poente).
INC_ZONA = {"I1": "poente_sul", "I2": "norte", "I3": "nascente_sul"}
CELULA_ZONA = {"CC 2501796": "poente_sul", "CC 200792": "poente_sul"}
# alcados de contencao -> frente, pela alinhamento das vigas de distribuicao
# do plano (EF/FG/GH na frente poente; IJ/JK/KL na frente norte; restantes
# pela continuidade do perimetro). Os que ficam duvidosos ficam None (global).
ALCADO_ZONA = {
    "AB": "nascente_sul", "CD": "poente_sul", "DE": "poente_sul",
    "EF": "poente_sul", "FG": "poente_sul", "GH": "poente_sul", "BF": "poente_sul",
    "HI": "norte", "IJ": "norte", "JK": "norte", "KL": "norte",
    "LM": "nascente_sul", "MN": "nascente_sul", "MNO": "nascente_sul",
    "OP": "nascente_sul", "PQ": "nascente_sul", "QR": "nascente_sul",
}

CORES_FASES = ["#8dd3c7", "#ffffb3", "#bebada", "#fb8072", "#80b1d3",
               "#fdb462", "#b3de69", "#fccde5", "#d9d9d9"]

# PLANO IMPACTADO vs CONTRATUAL. O plano impactado (22/04/2026) e o plano
# contratual com tres suspensoes de frentes (impactos) inseridas durante a
# execucao das estacas e recalculado: as 180 tarefas elementares de escavacao e
# travamento mantem a duracao contratual e foram deslocadas no tempo. As
# ligacoes de precedencia do ficheiro MS Project mostram por onde cada impacto
# se propagou. Datas transcritas do plano.
IMPACTOS_PLANO = [
    # (nome, inicio, fim, dias uteis, o que ficou suspenso, propagacao no plano)
    ("Impacto 1 — caneiro", "2025-05-13", "2025-07-01", 36,
     "Estacas E6_C_17 a E6_C_2 e E8_51 (alinhamento coincidente com o caneiro)",
     "Integral (+36 d.u.), pela sequência das estacas da frente poente e do corpo central/nascente"),
    ("Impacto 2 — cabos BT/MT", "2025-05-13", "2025-07-21", 50,
     "Estacas E6_13 a E6_15 e E8_113 a E8_116 (coincidentes ou encostadas aos cabos)",
     "Parcial (+32 d.u.): o plano permite iniciar as estacas seguintes 18 d.u. antes do fim da suspensão"),
    ("Impacto 3 — Edifício 2", "2025-05-13", "2025-08-21", 73,
     "Frente junto ao Edifício 2; demolição do muro antigo (adjudicação 28/07, desmonte 29/07 a 21/08)",
     "Nenhuma: no plano, a suspensão não precede outras tarefas"),
]
# (fase, ini_contratual, fim_contratual, ini_impactado, fim_impactado)
FASES_COMPARACAO = [
    ("Estacas Poente e Norte",        "2025-05-13", "2025-06-23", "2025-05-13", "2025-08-12"),
    ("Estacas Central e Nascente",    "2025-06-24", "2025-08-18", "2025-08-13", "2025-10-07"),
    ("Estacas periféricas (Edifício 2)",   "2025-06-26", "2025-08-18", "2025-08-07", "2025-09-25"),
    ("Viga coroamento + muros",       "2025-07-08", "2025-09-24", "2025-07-08", "2025-11-13"),
    ("Nível piso 2",                  "2025-07-24", "2025-09-12", "2025-07-24", "2025-10-21"),
    ("Nível piso 1",                  "2025-08-12", "2025-10-06", "2025-09-15", "2025-11-25"),
    ("Nível piso -1",                 "2025-09-30", "2025-11-17", "2025-11-13", "2026-01-06"),
    ("Nível piso -2",                 "2025-10-30", "2026-01-26", "2025-12-15", "2026-03-11"),
    ("Nível piso -3",                 "2026-01-06", "2026-02-26", "2026-02-25", "2026-04-13"),
    ("Fundo de escavação",            "2026-01-20", "2026-03-02", "2026-03-10", "2026-04-15"),
]
# Desfasamento por frente (dias uteis seg-sex; dias corridos) e impacto de origem,
# seguindo as ligacoes de precedencia do plano.
DESFASAMENTO_FRENTES = [
    ("Poente/Sul", "Edifício 1", "Impacto 1 — caneiro", "+36", "≈ +50",
     "20/01–02/02/2026", "10/03–23/03/2026"),
    ("Corpo central/nascente (estacas E8_2–E8_44, viga de coroamento)", "—",
     "Impacto 1 — caneiro", "+36", "≈ +50", "—", "—"),
    ("Norte", "Edifício 3", "Impacto 2 — cabos BT/MT", "+32", "≈ +44 a +46",
     "26/02–02/03/2026", "13/04–15/04/2026"),
    ("Nascente/Sul (escavação)", "Edifício 2", "Impacto 2 (via bandas de laje dos pisos 2 e 1)",
     "+32", "≈ +44 a +46", "16/02–27/02/2026", "01/04–14/04/2026"),
    ("Junto ao Edifício 2 (estacas periféricas, viga, muro M03, banda do piso 1)", "Edifício 2",
     "Impacto 2 (as estacas seguem as da frente norte)", "+28 a +30", "≈ +38 a +42", "—", "—"),
    ("Sem desfasamento: primeiras estacas poente e norte, viga de coroamento E8_77–E8_104 e "
     "ancoragens, escoras F-G/G-H e H-I/I-J, VD do piso 1 em EF/FG/GH e ancoragens", "Edifício 1 (pisos 2 e 1)",
     "—", "0", "0", "—", "—"),
]


def _atribuir_sublinhas(fases):
    """Interval partitioning: distribui as fases por sub-linhas de modo que
    dentro de cada sub-linha nenhuma se sobreponha no tempo. Recebe tuplos
    (t0, t1, nome, cor, n) ordenados por t0; devolve (lista com +li, n_linhas)."""
    linhas = []
    resultado = []
    for item in fases:
        t0, t1 = item[0], item[1]
        colocada = False
        for li, ocup in enumerate(linhas):
            if all(t1 <= o0 or t0 >= o1 for o0, o1 in ocup):
                ocup.append((t0, t1))
                resultado.append(item + (li,))
                colocada = True
                break
        if not colocada:
            linhas.append([(t0, t1)])
            resultado.append(item + (len(linhas) - 1,))
    return resultado, len(linhas)


def adicionar_fases_obra(fig, dt_min, dt_max, faixas=True, marcos=True,
                         barra_topo=True):
    """
    Sobrepõe as fases da obra a um gráfico com o tempo no eixo X.

    barra_topo=True (por omissão): BARRA DE FASEAMENTO (tipo Gantt) numa banda
    no topo do próprio gráfico — cada fase e um segmento de COR SÓLIDA (sem
    faixas de fundo translúcidas, que se misturavam onde as fases se
    sobrepunham), com o seu NÚMERO fixo. Fases sobrepostas são empilhadas em
    sub-linhas para não se misturarem. O eixo Y dos dados e encolhido (domain)
    para abrir espaço; a Gantt partilha o eixo X e alinha automaticamente.
    NÃO usar com eixos Y duplos (overlaying) — o domain só afeta um eixo.

    barra_topo=False: modo legado para gráficos de eixo Y duplo (ex. Correlações)
    — faixas de fundo translúcidas + número no início de cada fase, SEM tocar
    no domain do eixo Y.

    O nome completo esta na legenda a direita (trace fantasma) e no hover.
    Devolve a lista de fases visíveis (para a legenda).
    """
    dt_min = pd.to_datetime(dt_min)
    dt_max = pd.to_datetime(dt_max)
    margem = pd.Timedelta(days=20)

    visiveis = []
    for i, (nome, ini, fim) in enumerate(FASES_OBRA):
        t0, t1 = pd.to_datetime(ini), pd.to_datetime(fim)
        if t1 < dt_min - margem or t0 > dt_max + margem:
            continue
        # numero FIXO global (posicao em FASES_OBRA, base 1): a fase X tem
        # sempre o mesmo numero em toda a app, com saltos (1, 3, 5) se faltarem.
        num_fixo = i + 1
        visiveis.append((t0, t1, nome, CORES_FASES[i % len(CORES_FASES)], num_fixo))
    if not visiveis:
        return visiveis
    visiveis.sort(key=lambda v: v[0])

    # entradas de legenda (trace fantasma, cor SOLIDA = igual ao segmento Gantt)
    for t0, t1, nome, cor, n in visiveis:
        fig.add_trace(go.Scatter(
            x=[None], y=[None], mode="markers",
            name=f"{n}. {nome}", legendgroup="fases",
            legendgrouptitle_text="Fases da obra",
            marker=dict(size=13, symbol="square", color=cor,
                        line=dict(color="white", width=1)),
            hoverinfo="skip", showlegend=True))

    if not faixas:
        return visiveis

    lim_esq = dt_min - margem
    lim_dir = dt_max + margem

    # --- modo legado (eixo Y duplo): faixas translucidas + numero, sem domain
    if not barra_topo:
        # IMPORTANTE: com eixo Y secundario (overlaying="y"), o add_vrect ancora
        # as faixas ao dominio do eixo primario ("y domain"), o que entra em
        # conflito com o eixo sobreposto e FAZ AS SERIES DE DADOS DESAPARECEREM
        # no browser. Por isso desenhamos as faixas/linhas com yref="paper"
        # (independentes de qualquer eixo Y).
        for t0, t1, nome, cor, n in visiveis:
            vt0 = max(t0, lim_esq)
            vt1 = min(t1, lim_dir)
            fig.add_shape(
                type="rect", xref="x", yref="paper",
                x0=vt0, x1=vt1, y0=0, y1=1,
                fillcolor=cor, opacity=0.13, line_width=0, layer="below")
            if marcos and lim_esq <= t0 <= lim_dir:
                fig.add_shape(
                    type="line", xref="x", yref="paper",
                    x0=t0, x1=t0, y0=0, y1=1,
                    line=dict(color=cor, width=1, dash="dot"), layer="below")
            xc = vt0 + (vt1 - vt0) / 2
            fig.add_annotation(
                x=xc, y=0.99, yref="paper", xref="x", text=f"<b>{n}</b>",
                showarrow=False, xanchor="center", yanchor="top",
                font=dict(size=11, color="white"),
                bgcolor=cor, borderpad=3, opacity=0.95, hovertext=nome)
        return visiveis

    # --- modo barra no topo (Gantt) ---
    # segmentos clampados a janela visivel e empilhados em sub-linhas
    clamp = [(max(t0, lim_esq), min(t1, lim_dir), nome, cor, n)
             for t0, t1, nome, cor, n in visiveis]
    res, n_lin = _atribuir_sublinhas(clamp)

    # reservar banda no topo: dados passam a ocupar [0, ytop] do eixo Y
    frac_por_linha = 0.055
    banda = min(0.42, frac_por_linha * n_lin + 0.02)
    ytop = 1 - banda
    fig.update_yaxes(domain=[0, ytop])

    # desenhar cada segmento como retangulo SOLIDO na sua sub-linha (paper-y),
    # com o numero centrado. Linha 0 fica em cima; sublinhas descem.
    gap = 0.012
    alt = (banda - gap) / n_lin
    for t0, t1, nome, cor, n, li in res:
        y1 = 1 - li * alt
        y0 = y1 - alt * 0.82
        fig.add_shape(
            type="rect", xref="x", yref="paper",
            x0=t0, x1=t1, y0=y0, y1=y1,
            fillcolor=cor, opacity=1.0,
            line=dict(color="white", width=1), layer="above")
        xc = t0 + (t1 - t0) / 2
        yc = (y0 + y1) / 2
        fig.add_annotation(
            x=xc, y=yc, xref="x", yref="paper", text=f"<b>{n}</b>",
            showarrow=False, xanchor="center", yanchor="middle",
            font=dict(size=11, color="white"), hovertext=nome)
        # linha ponteada de inicio real da fase (se cair na janela)
        if marcos and lim_esq <= t0 <= lim_dir:
            fig.add_vline(x=t0, line=dict(color=cor, width=1, dash="dot"),
                          layer="below")
    return visiveis


def legenda_fases(visiveis):
    """Escreve, por baixo do gráfico, a legenda número -> nome das fases.
    Usa o número FIXO global que vêm no tuplo (não re-enumera)."""
    if not visiveis:
        return
    itens = "  ·  ".join(f"**{n}**. {nome}"
                         for _, _, nome, _, n in visiveis)
    st.caption("Fases da obra (datas do plano de trabalhos impactado): " + itens)


def barra_faseamento(dt_min, dt_max, altura=34):
    """
    Desenha uma mini-barra de faseamento (tipo Gantt) para o período visível:
    cada fase e uma barra horizontal na sua própria linha, com o nome legível,
    sem sobreposições. Devolve uma figura plotly compacta para colocar POR CIMA
    do gráfico principal — assim as etiquetas das fases saem de dentro do
    gráfico e deixam de colidir.
    """
    import plotly.graph_objects as go
    dt_min = pd.to_datetime(dt_min)
    dt_max = pd.to_datetime(dt_max)
    margem = pd.Timedelta(days=20)

    visiveis = []
    for i, (nome, ini, fim) in enumerate(FASES_OBRA):
        t0, t1 = pd.to_datetime(ini), pd.to_datetime(fim)
        if t1 < dt_min - margem or t0 > dt_max + margem:
            continue
        visiveis.append((nome, max(t0, dt_min), min(t1, dt_max),
                         CORES_FASES[i % len(CORES_FASES)]))
    if not visiveis:
        return None

    fig = go.Figure()
    for linha, (nome, t0, t1, cor) in enumerate(visiveis):
        y = len(visiveis) - linha          # uma linha por fase (topo->fundo)
        fig.add_trace(go.Scatter(
            x=[t0, t1], y=[y, y], mode="lines",
            line=dict(color=cor, width=14),
            hovertemplate=f"{nome}<br>%{{x|%d/%m/%Y}}<extra></extra>",
            showlegend=False))
        # nome da fase, alinhado a esquerda no inicio da barra
        fig.add_annotation(x=t0, y=y, text=" " + nome, xanchor="left",
                           yanchor="middle", showarrow=False,
                           font=dict(size=10, color="#333"))
    fig.update_yaxes(visible=False, range=[0.3, len(visiveis) + 0.7])
    fig.update_xaxes(range=[dt_min, dt_max], showticklabels=False,
                     showgrid=False)
    fig.update_layout(height=altura * len(visiveis) + 20,
                      margin=dict(l=0, r=0, t=4, b=0),
                      plot_bgcolor="white")
    return fig


def configurar_eixo_tempo(fig, granularidade="Automático"):
    """
    Define a granularidade das marcas do eixo temporal (X).
    'Mensal' -> 1 marca/mês; 'Quinzenal' -> de 15 em 15 dias;
    'Semanal' -> de 7 em 7 dias; 'Automático' -> deixa o plotly decidir.
    Marcas mais finas ajudam a ler o faseamento da obra ao nível a que as
    campanhas existem (~8 em 8 dias).
    """
    # as etiquetas em portugues sao geradas por _datas_pt (antes de desenhar)
    if granularidade == "Mensal":
        fig.update_xaxes(dtick="M1", tickangle=-30)
    elif granularidade == "Quinzenal":
        fig.update_xaxes(dtick=14 * 24 * 3600 * 1000, tickangle=-45)
    elif granularidade == "Semanal":
        fig.update_xaxes(dtick=7 * 24 * 3600 * 1000, tickangle=-45)
    # 'Automatico' -> nao mexe


# =========================================================================
# DATAS EM PORTUGUES NOS GRAFICOS
# -------------------------------------------------------------------------
# O plotly.js escreve os meses em ingles ("Dec 14") e o Streamlit nao carrega
# os ficheiros de localizacao. Antes de cada grafico ser desenhado, os eixos X
# com datas recebem marcas explicitas (tickvals/ticktext) com os meses em
# portugues: mensais ("out 2025") ou de 7/14 dias ("15 dez"), conforme a
# granularidade escolhida (dtick) ou a extensao do periodo.
# =========================================================================
_MESES_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set",
             "out", "nov", "dez"]


def _datas_do_eixo(fig, ref):
    datas = []
    for tr in fig.data:
        if (getattr(tr, "xaxis", None) or "x") != ref:
            continue
        for arr in (getattr(tr, "x", None), getattr(tr, "base", None)):
            if arr is None:
                continue
            try:
                v = pd.to_datetime(pd.Series(list(arr)), errors="coerce")
            except Exception:
                continue
            if len(v) and v.notna().mean() > 0.8 and not pd.api.types.is_numeric_dtype(pd.Series(list(arr))):
                datas.extend(v.dropna().tolist())
    return datas


def _datas_pt(fig):
    try:
        lay = fig.layout
        refs = {(getattr(tr, "xaxis", None) or "x") for tr in fig.data}
        for ref in refs:
            nome = "xaxis" if ref == "x" else "xaxis" + ref[1:]
            ax = lay[nome] if nome in lay else None
            if ax is None or ax.tickvals is not None:
                continue
            datas = _datas_do_eixo(fig, ref)
            if not datas:
                continue
            d0, d1 = min(datas), max(datas)
            if ax.range is not None:
                try:
                    d0, d1 = pd.to_datetime(ax.range[0]), pd.to_datetime(ax.range[1])
                except Exception:
                    pass
            dias = max((d1 - d0).days, 1)
            dt = ax.dtick
            if dt == "M1" or (dt is None and dias > 120):
                passo = 1 if dias <= 400 else (2 if dias <= 800 else 3)
                t = pd.Timestamp(d0.year, d0.month, 1)
                if t < d0:
                    t = t + pd.DateOffset(months=1)
                vals = []
                while t <= d1:
                    vals.append(t); t = t + pd.DateOffset(months=passo)
                txt = [f"{_meses()[v.month - 1]} {v.year}" for v in vals]
            else:
                step = 7 if (isinstance(dt, (int, float)) and dt <= 7 * 86400000 + 1) or (dt is None and dias <= 60) else 14
                vals = list(pd.date_range(pd.Timestamp(d0.date()), d1, freq=f"{step}D"))
                txt = [f"{v.day} {_meses()[v.month - 1]}" + (f"<br>{v.year}" if (k == 0 or v.month == 1 and v.day <= step) else "")
                       for k, v in enumerate(vals)]
            if len(vals) >= 2:
                ax.update(tickmode="array", tickvals=vals, ticktext=txt,
                          hoverformat="%d/%m/%Y")
    except Exception:
        pass
    return fig


_plotly_chart_original = (st.plotly_chart if hasattr(st.plotly_chart, "__self__")
                          else st._main.plotly_chart)


def _plotly_chart_pt(fig, *args, **kwargs):
    return _plotly_chart_original(_datas_pt(fig), *args, **kwargs)


st.plotly_chart = _plotly_chart_pt


# =========================================================================
# IDIOMA DA INTERFACE (Português / English)
# -------------------------------------------------------------------------
# O idioma escolhe-se na barra lateral («Idioma / Language»). Em inglês, os
# textos mostrados (títulos, legendas, rótulos dos controlos, separadores,
# tabelas e gráficos) são traduzidos à saída, através do dicionário
# _I18N_EN: primeiro procura-se a frase inteira e, se não existir, traduzem-
# -se os fragmentos conhecidos (termos e partes fixas das f-strings). Os
# VALORES devolvidos pelos controlos não mudam (só o que é mostrado), pelo
# que a lógica da aplicação é a mesma nos dois idiomas. Textos que não
# constem do dicionário ficam em português.
# =========================================================================
import re as _re_i18n
from streamlit.delta_generator import DeltaGenerator as _DG_i18n

_I18N_EN = {
    '#### 7. Elementos adjacentes — cada um com o seu critério': '#### 7. Adjacent elements — each with its own criterion',
    'Compara dois alvos de elementos diferentes, por exemplo um edifício vizinho e o alçado da contenção que lhe fica em frente. Cada série é lida com os limiares do critério do seu elemento, e a faixa inferior mostra o estado de cada alvo segundo esse critério. Opcionalmente, o primeiro alvo é também representado a partir da data de zeragem do segundo, para comparar movimentos com a mesma origem.': 'Compares two targets of different elements, for example a neighbouring building and the wall section facing it. Each series is read with the thresholds of the criterion of its element, and the lower band shows the state of each target under that criterion. Optionally, the first target is also plotted from the zeroing date of the second, to compare movements with the same origin.',
    'São precisos pelo menos dois alvos.': 'At least two targets are needed.',
    'Alvo 1': 'Target 1',
    'Alvo 2': 'Target 2',
    'Representar também o alvo 1 a partir da zeragem do alvo 2': 'Also plot target 1 from the zeroing of target 2',
    'Os dois alvos não têm leituras num período comum.': 'The two targets have no readings in a common period.',
    'Período': 'Period',
    'Deslocamento horizontal (mm)': 'Horizontal displacement (mm)',
    '#### C. Sinal generalizado por elemento': '#### C. Generalised signal by element',
    'Persistência no espaço: numa campanha, um elemento (edifício vizinho ou alçado) dá sinal generalizado quando pelo menos a fração escolhida dos seus alvos ou prismas lidos nessa campanha (mínimo de 3) dá sinal pela regra da barra lateral, com a velocidade calculada sobre pelo menos 7 dias. Reduz o número de sinais a analisar face aos sinais por alvo.': 'Persistence in space: in a campaign, an element (neighbouring building or wall section) gives a generalised signal when at least the chosen fraction of its targets or prisms read in that campaign (minimum of 3) gives a signal by the rule of the sidebar, with the velocity computed over at least 7 days. It reduces the number of signals to analyse compared with the signals per target.',
    'Fração mínima de alvos com sinal': 'Minimum fraction of targets with a signal',
    'Sem elementos com pelo menos 3 leituras na mesma campanha.': 'No elements with at least 3 readings in the same campaign.',
    'Fração com sinal': 'Fraction with signal',
    'Campanha (fração com sinal)': 'Campaign (fraction with signal)',
    'Sinal generalizado': 'Generalised signal',
    'Primeiro alerta': 'First alert',
    'Sinais generalizados': 'Generalised signals',
    'Primeiro sinal generalizado': 'First generalised signal',
    'Antecedência (dias)': 'Lead time (days)',
    'Elemento': 'Element',
    ' sinais generalizados** em ': ' generalised signals** in ',
    ' pares de elemento e campanha com pelo menos 3 leituras, em ': ' element–campaign pairs with at least 3 readings, in ',
    ' elementos. Em **': ' elements. In **',
    ' elementos** que atingiram o alerta, o primeiro sinal generalizado precedeu-o': ' elements** that reached the alert, the first generalised signal preceded it',
    ' Como quase todos os elementos atingiram o alerta, a seletividade deste indicador não pode ser avaliada neste caso.': ' As almost all elements reached the alert, the selectivity of this indicator cannot be assessed in this case.',
    ' (data da análise do bloco A), sinal generalizado em: ': ' (analysis date of block A), generalised signal at: ',
    ' (data da análise do bloco A), nenhum elemento tinha dado sinal generalizado.': ' (analysis date of block A), no element had given a generalised signal.',
    'Primeiro alerta: ': 'First alert: ',
    ': sinal generalizado': ': generalised signal',
    ' alvos com sinal': ' targets with a signal',
    " mais.": " more.",
    " vs. previsto ": " vs. predicted ",
    "(frente ": "(front ",
    "poente": "west",
    "nascente": "east",
    "direta": "direct",
    "ganhou carga": "gained load",
    "perdeu carga": "lost load",
    ", terminando a ": ", ending at ",
    "Escavação e Contenção Periférica": "Excavation and retaining wall",
    "viga de coroamento": "capping beam",
    "viga de distribuição do piso": "distribution beam of level",
    "troço poente": "western stretch",
    "inferido": "inferred",
    "  Regra da aplicação (v ≥ 0,5 ou acel. ≥ 1,8)": "  Rule of the application (v ≥ 0.5 or accel. ≥ 1.8)",
    "  Só aceleração, ≥ 1,8×": "  Acceleration only, ≥ 1.8×",
    "  Só velocidade, v ≥ 0,3": "  Velocity only, v ≥ 0.3",
    "  Só velocidade, v ≥ 0,3, dois sinais seguidos": "  Velocity only, v ≥ 0.3, two consecutive signals",
    "  Só velocidade, v ≥ 0,5": "  Velocity only, v ≥ 0.5",
    "  Só velocidade, v ≥ 1,0": "  Velocity only, v ≥ 1.0",
    " (17 m assumido — não consta dos relatórios)": " (17 m assumed — not in the reports)",
    " (cor do projeto)": " (design colour)",
    " (cota ": " (elev. ",
    " (célula ": " (cell ",
    " (no Edifício 1)": " (on Building 1)",
    " (peças do projeto)": " (design drawings)",
    " + VD da grua (2 ancoragens)": " + crane distribution beam (2 anchors)",
    " No Edifício 1, a fachada frontal (losangos, média ": " At Building 1, the front façade (diamonds, mean ",
    " O segundo é «": " The second is «",
    " alvos, **": " targets, **",
    " atingiram o alerta** e ": " reached the alert** and ",
    " campanhas. Como a cota DESCE quando a escavação aprofunda, um r negativo significa que o deslocamento sobe a medida que a frente local desce. r mede a associação, não a causa.": " campaigns. As the level FALLS when the excavation deepens, a negative r means that the displacement rises as the local front goes down. r measures the association, not the cause.",
    " campanhas. Uma correlação negativa significa que o deslocamento sobe a medida que a cota da água desce. r mede a associação estatística, não a causa.": " campaigns. A negative correlation means that the displacement rises as the water level falls. r measures the statistical association, not the cause.",
    " carga: ": " load: ",
    " começou ~": " started ~",
    " desceu de **": " fell from **",
    " dias depois da ": " days after ",
    " dias úteis)<br>": " working days)<br>",
    " dias**. A regra é sensível mas pouco seletiva: a maioria dos alvos que nunca alertaram também gerou sinais, e calcular a velocidade sobre ≥7 dias melhora a seletividade à custa de antecipar menos alertas.": " days**. The rule is sensitive but poorly selective: most targets that never reached the alert also gave signals, and computing the velocity over ≥7 days improves the selectivity at the cost of anticipating fewer alerts.",
    " e o alvo já teria entrado em alerta (": " and the target would already have entered alert (",
    " em várias campanhas (cor mais clara = mais antiga; mais escura = mais recente).": " in several campaigns (lighter colour = older; darker = more recent).",
    " evoluiu de ": " evolved from ",
    " fica em ": " is at ",
    " instrumento(s) com leitura em atraso face à data de estado.": " instrument(s) with readings delayed relative to the status date.",
    " layers**. Estão pre-selecionadas as estruturais (contorno, estacas, grelha). Ajusta se quiseres:": " layers**. The structural ones are preselected (outline, piles, grid). Adjust if you wish:",
    " leitura(s) com aceleração acima dos critérios.": " reading(s) with acceleration above the criteria.",
    " leituras (100%). As duas implementações dos critérios são coerentes; a comparação com o estado publicado nos relatórios 33GRADOS é feita à parte.": " readings (100%). The two implementations of the criteria are consistent; the comparison with the state published in the 33GRADOS reports is made separately.",
    " leituras entre o estado recalculado e a coluna 'Estado' do Excel. Ver tabela de divergências abaixo.": " readings between the recalculated state and the 'Estado' column of the Excel file. See the discrepancies table below.",
    " m (cor conforme legenda do projeto)": " m (colour according to the design legend)",
    " m (projeto JETsj)<extra></extra>": " m (JETsj design)<extra></extra>",
    " m abaixo da água atual e ~": " m below the current water and ~",
    " m abaixo do coroamento<extra></extra>": " m below the capping beam<extra></extra>",
    " m abaixo do nível freático de repouso de 2022 (~": " m below the 2022 rest water table (~",
    " m de altura escavada no ponto mais alto. Fonte: levantamento topográfico (DXF). Notas: o exagero vertical é apenas visual; o filtro de fase mostra a escavação até a cota do piso do projeto (a progressão real e por fases). A coloração dos alçados por movimento, quando ativa, e ILUSTRATIVA e qualitativa — associa o estado dos alvos ao lado do recinto por orientação, não por coordenada exata, porque terreno e alvos usam referenciais distintos.": " m of excavated height at the highest point. Source: topographic survey (DXF). Notes: the vertical exaggeration is only visual; the stage filter shows the excavation down to the design floor elevation (the real progress is by stages). The colouring of the wall sections by movement, when enabled, is ILLUSTRATIVE and qualitative — it associates the state of the targets with the side of the site by orientation, not by exact coordinate, because terrain and targets use different reference systems.",
    " m do coroamento)": " m from the capping beam)",
    " m — aproximadamente ": " m — approximately ",
    " m) no período monitorizado. O fundo de escavação (**": " m) over the monitored period. The excavation bottom (**",
    " m) — as duas séries movem-se em sentido oposto ao longo do tempo.": " m) — the two series move in opposite directions over time.",
    " m** enquanto o ": " m** while the ",
    " m**) fica ": " m**) is ",
    " m; fundo de escavação a ": " m; excavation bottom at ",
    " mm) está drapejado sobre o relevo real. Vermelho = alarme, laranja = alerta, verde = regular. IMPORTANTE: a posição dos alvos sobre o terreno é aproximada (alinhamento por caixa-envolvente entre referenciais diferentes) — para leitura tendencial do movimento sobre o relevo, não métrica. A medição precisa está no separador Alvos (2D). Se algum alvo cair no lado errado, usa o 'Ajuste fino do alinhamento'.": " mm) is draped over the real relief. Red = alarm, orange = alert, green = regular. IMPORTANT: the position of the targets on the terrain is approximate (bounding-box alignment between different reference systems) — for a trend reading of the movement over the relief, not a metric one. The precise measurement is in the Targets (2D) tab. If a target falls on the wrong side, use 'Fine alignment adjustment'.",
    " mm) move-se mais que a lateral (círculos, ": " mm) moves more than the side one (circles, ",
    " mm) — coerente com a exposição direta a escavação.": " mm) — consistent with the direct exposure to the excavation.",
    " mm) — o primeiro move-se ": " mm) — the first moves ",
    " mm), no elemento «": " mm), in element «",
    " mm** enquanto a cota da água desceu de ": " mm** while the water level fell from ",
    " mm, média ": " mm, mean ",
    " mm/dia; aceleração ≥ ": " mm/day; acceleration ≥ ",
    " nunca o atingiram (contagens «de 38» e «de 66» na dissertação). Com a regra da aplicação, a antecedência mediana entre o primeiro sinal e o alerta é de **": " never reached it (counts «of 38» and «of 66» in the dissertation). With the rule of the application, the median lead time between the first signal and the alert is **",
    " não tem frente de escavação mapeada. Escolhe um alvo de um edifício/alçado com zona associada (Edifícios 1, 2 e 3, ou um alçado de contenção).": " has no mapped excavation front. Choose a target of a building/wall section with an associated zone (Buildings 1, 2 and 3, or a retaining wall section).",
    " para **": " to **",
    " sobrepostos (SPT no eixo de cima). A leitura central do back-analysis: ve se o 'joelho' de maior deformação do perfil coincide com uma subida do SPT (grés a consolidar) ou com o nível freático. Onde o inclinómetro passa da base da sondagem, não há dado geológico.": " overlaid (SPT on the top axis). The central reading of the back-analysis: see whether the 'knee' of largest deformation of the profile coincides with a rise of the SPT (sandstone becoming consolidated) or with the water table. Where the inclinometer goes beyond the base of the borehole, there is no geological data.",
    " — cota ": " — elev. ",
    "#### 1. Deformação × Água × Escavação × Faseamento": "#### 1. Deformation × Water × Excavation × Construction phases",
    "#### 1. Quanto — magnitude das movimentações": "#### 1. How much — magnitude of the movements",
    "#### 2. Deformação × Carga na ancoragem": "#### 2. Deformation × Anchor load",
    "#### 2. Onde — distribuição espacial": "#### 2. Where — spatial distribution",
    "#### 3. Correlação direta — deslocamento vs. cota da água": "#### 3. Direct correlation — displacement vs. water level",
    "#### 3. Quando e a que ritmo — velocidade de movimentação": "#### 3. When and at what rate — movement velocity",
    "#### 4. A que profundidade — perfil inclinométrico": "#### 4. At what depth — inclinometer profile",
    "#### 4. Deformação × Escavação da frente local": "#### 4. Deformation × Excavation of the local front",
    "#### 5. Matriz de correlações entre instrumentos": "#### 5. Correlation matrix between instruments",
    "#### 5. Observado vs. Previsto — comparação com o projeto": "#### 5. Observed vs. Predicted — comparison with the design",
    "#### 6. Envelopes empíricos — δ_h,máx em função de H": "#### 6. Empirical envelopes — δ_h,max as a function of H",
    "#### A. Leitura em tempo real simulada": "#### A. Simulated real-time reading",
    "#### Ancoragens — plano impactado vs. registos de execução": "#### Anchors — impacted schedule vs. execution records",
    "#### B. Sensibilidade e seletividade da regra (todos os alvos)": "#### B. Sensitivity and selectivity of the rule (all targets)",
    "#### Cobertura das leituras — avisos de atraso": "#### Reading coverage — delay warnings",
    "#### Cronologia de execução das ancoragens": "#### Anchor execution chronology",
    "#### Direção e proporção do movimento (última campanha)": "#### Direction and proportion of the movement (last campaign)",
    "#### Fissurómetros do Edifício 1 (F1–F5)": "#### Building 1 crack gauges (F1–F5)",
    "#### Impactos e desfasamento por frente": "#### Impacts and delay by front",
    "#### Instrumentação monitorizada": "#### Monitored instrumentation",
    "#### O que podes explorar": "#### What you can explore",
    "#### Obra": "#### Works",
    "#### Ocorrências e medidas indicadas pela obra": "#### Events and measures reported by the site",
    "#### Pressupostos e inferências da ferramenta": "#### Assumptions and inferences of the tool",
    "#### Questões de qualidade dos dados (do relatório)": "#### Data quality issues (from the report)",
    "#### Tratamento de dados (critérios aplicados)": "#### Data processing (criteria applied)",
    "#### Zonamento geotécnico e parâmetros de projeto": "#### Geotechnical zoning and design parameters",
    "##### Efeito térmico": "##### Thermal effect",
    "% acima da estimativa de projeto.": "% above the design estimate.",
    "% da blocagem (": "% of the lock-off load (",
    "%{y} · furação %{x|%d/%m/%Y}<extra></extra>": "%{y} · drilling %{x|%d/%m/%Y}<extra></extra>",
    "%{y} · tensionamento %{x|%d/%m/%Y}<br>%{customdata:.0f} dias após injeção<extra></extra>": "%{y} · stressing %{x|%d/%m/%Y}<br>%{customdata:.0f} days after grouting<extra></extra>",
    "%{y} · tensionamento %{x|%d/%m/%Y}<extra></extra>": "%{y} · stressing %{x|%d/%m/%Y}<extra></extra>",
    "' está vazia ou não foi encontrada.": "' is empty or was not found.",
    "'. Poe o Excel na pasta do script ou usa 'Carregar manualmente'.": "'. Put the Excel file in the script folder or use 'Load manually'.",
    ") está na **": ") is at the **",
    ") não está confirmada. O cruzamento só tem significado físico se o alvo pertencer ao mesmo alçado da ancoragem.": ") is not confirmed. The crossing only has physical meaning if the target belongs to the same wall section as the anchor.",
    ")** — recalculado dos deslocamentos ΔH/ΔV com os critérios oficiais:": ")** — recalculated from the ΔH/ΔV displacements with the official criteria:",
    "). Eixo A+ orientado a ": "). A+ axis oriented to ",
    "). Os maiores intervalos concentram-se no alçado GH (viga de distribuição do piso −2) e no troço poente do AB (piso −1), onde o tensionamento foi diferido.": "). The largest intervals are concentrated in section GH (level −2 distribution beam) and in the western stretch of AB (level −1), where stressing was deferred.",
    "** (confiança ": "** (confidence ",
    "** ancoragens ultrapassam a janela de 7–8 dias entre injeção e tensionamento e **": "** anchors exceed the 7–8 day window between grouting and stressing and **",
    "** cota ": "** elev. ",
    "** para **": "** to **",
    "** ultrapassam 20 dias (até ": "** exceed 20 days (up to ",
    "**, diferente da sugerida.": "**, different from the suggested one.",
    "**, o primeiro sinal já teria sido lido a **": "**, the first signal would already have been read on **",
    "**. As linhas tracejadas nos gráficos marcam os limiares de alerta e alarme.": "**. The dashed lines in the charts mark the alert and alarm thresholds.",
    "**. O seletor de alvos abaixo está restrito a essa zona — o cruzamento é fisicamente válido.": "**. The target selector below is restricted to that zone — the crossing is physically valid.",
    "**. Os marcos abaixo são as cotas atingidas nessa frente.": "**. The milestones below are the levels reached at that front.",
    "**. Um instrumento é assinalado quando o intervalo desde a sua última leitura excede o dobro da sua cadência habitual.": "**. An instrument is flagged when the interval since its last reading exceeds twice its usual interval.",
    "**Comparação entre células (ancoragens)**": "**Comparison between cells (anchors)**",
    "**Correlação direta — deslocamento vs. cota escavada local**": "**Direct correlation — displacement vs. local excavated level**",
    "**Correlação:** r = **": "**Correlation:** r = **",
    "**Estado na última campanha (": "**State in the last campaign (",
    "**Hotel Eden, Estoril**  \nReformulação — escavação e contenção periférica  \nMonte Estoril, Cascais  \n_Back-analysis de instrumentação_": "**Hotel Eden, Estoril**  \nRedevelopment — excavation and peripheral retaining structure  \nMonte Estoril, Cascais  \n_Instrumentation back-analysis_",
    "**Leitura:** a água no ": "**Reading:** the water at ",
    "**Leitura:** na frente da ": "**Reading:** at the front of ",
    "**Leitura:** nenhuma atividade tem, no plano impactado, duração diferente da contratual — as 180 tarefas elementares de escavação e travamento mantêm a duração e foram deslocadas. Todo o atraso vem das suspensões de frentes durante as estacas: o impacto 1 (caneiro) desloca a frente poente, sob o Edifício 1, em 36 dias úteis; o impacto 2 (cabos BT/MT) desloca a frente norte e a escavação da frente nascente/sul em 32. A escavação termina a **15/04/2026** em vez de 02/03/2026 e a conclusão da 1.ª fase passa de 11/01/2027 para 01/03/2027.": "**Reading:** no activity has, in the impacted schedule, a duration different from the contractual one — the 180 elementary excavation and bracing tasks keep their duration and were shifted. All the delay comes from the suspensions of fronts during piling: impact 1 (culvert) shifts the west front, below Building 1, by 36 working days; impact 2 (LV/MV cables) shifts the north front and the excavation of the east/south front by 32. The excavation ends on **15/04/2026** instead of 02/03/2026 and the completion of phase 1 moves from 11/01/2027 to 01/03/2027.",
    "**Leitura:** no período, a deformação evoluiu de ": "**Reading:** over the period, the deformation evolved from ",
    "**Ponto de referência 1**": "**Reference point 1**",
    "**Ponto de referência 2** (escolhe um bem afastado do 1)": "**Reference point 2** (choose one well away from point 1)",
    "**Ranking dos 10 alvos mais deslocados** (última campanha, ": "**Ranking of the 10 most displaced targets** (last campaign, ",
    "**Razão assentamento/horizontal (|ΔZ|/δh)**": "**Settlement/horizontal ratio (|ΔZ|/δh)**",
    "**entre leituras consecutivas**. Ativa a opção ≥ 7 dias na barra lateral para reduzir o ruído.": "**between consecutive readings**. Enable the ≥ 7 days option in the sidebar to reduce the noise.",
    "**≥ 7 dias** (regra melhorada).": "**≥ 7 days** (improved rule).",
    "+15% alerta": "+15% alert",
    "+25% alarme": "+25% alarm",
    ", a escavação desceu de cota **": ", the excavation went down from elev. **",
    ", antes de qualquer alerta.": ", before any alert.",
    ", por isso perdeu a fase inicial de carregamento. A comparação das variações (%) entre períodos diferentes deve ser lida com esta ressalva.": ", so it missed the initial loading phase. The comparison of the variations (%) between different periods must be read with this caveat.",
    ". Fonte: relatório de instrumentação (33GRADOS).": ". Source: instrumentation report (33GRADOS).",
    ". Imagens das peças desenhadas do projeto de contenção (JETsj).": ". Images of the drawings of the retaining structure design (JETsj).",
    ". Setas e contornos: vermelho = alarme, laranja = alerta, verde = regular (critérios oficiais recalculados). A caixa mostra a profundidade real de escavação (16,3 m, do projeto) como distância abaixo da cortina — os alvos e o projeto usam referenciais de cota diferentes, por isso e profundidade, não cota absoluta.": ". Arrows and outlines: red = alarm, orange = alert, green = regular (recalculated official criteria). The box shows the real excavation depth (16.3 m, from the design) as a distance below the wall — the targets and the design use different elevation references, so it is depth, not absolute elevation.",
    ". Verifica o bloco CONFIGURAÇÃO (COLS) ou o Excel.": ". Check the CONFIGURATION block (COLS) or the Excel file.",
    "09–16/12/2025 (VD, 2 lados)": "09–16/12/2025 (distribution beam, 2 sides)",
    "09–16/12/2025 (VD, 2 lados)?": "09–16/12/2025 (distribution beam, 2 sides)?",
    "1.º sinal": "1st signal",
    "22–29/01/2026 (VD, 2 lados)": "22–29/01/2026 (distribution beam, 2 sides)",
    "22–29/01/2026 (VD, 2 lados)?": "22–29/01/2026 (distribution beam, 2 sides)?",
    "28/08–09/09/2025 (tarefa única)": "28/08–09/09/2025 (single task)",
    "30/10–03/11/2025 (coroamento E8_51–E8_35)?": "30/10–03/11/2025 (capping beam E8_51–E8_35)?",
    "9–20 dias": "9–20 days",
    "; comparação ": "; comparison ",
    "<br>Estado: ": "<br>State: ",
    "<br>Propagação: ": "<br>Propagation: ",
    "<div style='border-radius:12px; height:190px; margin-bottom:18px; background:linear-gradient(120deg, #0f2037 0%, #1f3a5f 55%, #2e5c8a 100%); display:flex; align-items:center; padding-left:36px;'><div style='color:white;'><div style='font-size:3.2em; font-weight:800; letter-spacing:3px; line-height:1;'>IMS</div><div style='font-size:1.2em; opacity:0.92; margin-top:8px;'>Instrumentation Monitoring System</div><div style='font-size:0.95em; opacity:0.8; margin-top:12px; max-width:560px;'>Análise e visualização de instrumentação geotécnica — deslocamentos, velocidades, sinais precursores, geologia e sequência de obra.</div></div></div>": "<div style='border-radius:12px; height:190px; margin-bottom:18px; background:linear-gradient(120deg, #0f2037 0%, #1f3a5f 55%, #2e5c8a 100%); display:flex; align-items:center; padding-left:36px;'><div style='color:white;'><div style='font-size:3.2em; font-weight:800; letter-spacing:3px; line-height:1;'>IMS</div><div style='font-size:1.2em; opacity:0.92; margin-top:8px;'>Instrumentation Monitoring System</div><div style='font-size:0.95em; opacity:0.8; margin-top:12px; max-width:560px;'>Analysis and visualisation of geotechnical instrumentation — displacements, velocities, precursor signals, geology and construction sequence.</div></div></div>",
    "O Edifício 1 tem duas fachadas instrumentadas: **Frente a escavação** (A1–A4) e **Lateral, virada ao mar** (A5–A8). Os alvos A5–A8 foram tapados por um painel publicitário e substituídos por **A5b–A8b**, RE-ZERADOS em 20/10/2025 — por isso os acumulados dos 'b' não são comparáveis diretamente com A1–A4 (arrancam de zero mais tarde).": "Building 1 has two instrumented façades: **Facing the excavation** (A1–A4) and **Side, facing the sea** (A5–A8). Targets A5–A8 were covered by an advertising panel and replaced by **A5b–A8b**, RE-ZEROED on 20/10/2025 — so the cumulative values of the 'b' targets are not directly comparable with A1–A4 (they start from zero later).",
    "A ancoragem ": "Anchor ",
    "A caixa de escavação usa a PROFUNDIDADE real (16,3 m), não a cota absoluta — os sistemas de cota dos alvos (Z 42-63) e do projeto (4,5-24,6) não estão relacionados.": "The excavation box uses the real DEPTH (16.3 m), not the absolute elevation — the elevation systems of the targets (Z 42-63) and of the design (4.5-24.6) are not related.",
    "A carregar o Excel...": "Loading the Excel file...",
    "A coluna **Temperatura (°C)** existe no modelo mas está vazia, por isso a app não consegue separar o efeito térmico do estrutural. É a dúvida da Secção 5.2.8: as duas células subiram em simultâneo entre as leituras de final de junho e a de 06/07/2026 — a A26 +27,6 kN e a da ancoragem DE +24,2 kN. Com a deriva do sensor da A26 (0,75 kN/°C, certificado de calibração), seria preciso ≈ 37 °C de variação para explicar o salto só por temperatura. Este gráfico (carga × temperatura) aparece automaticamente quando carregares as temperaturas das células.": "The **Temperature (°C)** column exists in the model but is empty, so the app cannot separate the thermal effect from the structural one. This is the question of Section 5.2.8: both cells rose simultaneously between the readings of late June and that of 06/07/2026 — A26 by +27.6 kN and the cell of anchor DE by +24.2 kN. With the drift of the A26 sensor (0.75 kN/°C, calibration certificate), a variation of ≈ 37 °C would be needed to explain the jump by temperature alone. This chart (load × temperature) appears automatically when the cell temperatures are loaded.",
    "A coluna colorida à esquerda de cada sondagem é a litologia; os pontos e a linha são o SPT (N pancadas), no mesmo eixo de profundidade. A cor do ponto SPT indica a zona geotécnica provável pelo valor de N (ver tabela em baixo). A linha azul tracejada é o nível freático.": "The coloured column to the left of each borehole is the lithology; the points and the line are the SPT (N blows), on the same depth axis. The colour of the SPT point indicates the probable geotechnical zone from the N value (see table below). The dashed blue line is the water table.",
    "A confirmar (falta 1 par de cotas nos 2 sistemas)": "To be confirmed (1 pair of elevations in the 2 systems missing)",
    "A confirmar com a equipa de instrumentação": "To be confirmed with the instrumentation team",
    "A cor da zona (ZG) é a provável pelo valor de N; o relatório não define fronteiras de zona por profundidade.": "The zone colour (ZG) is the probable one from the N value; the report does not define zone boundaries by depth.",
    "A escala calculada e invulgar. Confirma as coordenadas dos pontos no desenho — pode haver troca de X/Y ou de ponto.": "Unusual computed scale. Check the coordinates of the points in the drawing — X/Y or the points may have been swapped.",
    "A extrair geometria do DXF...": "Extracting DXF geometry...",
    "A folha de '": "The sheet '",
    "A ler o desenho DXF...": "Reading the DXF drawing...",
    "A regra e as suas variantes aplicadas ao deslocamento horizontal de todos os alvos e prismas. Contam-se os sinais anteriores ao primeiro alerta de cada alvo (todos, nos alvos sem alerta). Um alerta diz-se antecipado se teve um sinal nas 4 semanas anteriores; um sinal é acerto se o alerta do mesmo alvo se seguiu nesse prazo. Reproduz a Tabela 5.5.": "The rule and its variants applied to the horizontal displacement of all targets and prisms. Signals before the first alert of each target are counted (all signals, in targets without alert). An alert is anticipated if it had a signal in the 4 preceding weeks; a signal is a hit if the alert of the same target followed within that period. Reproduces Table 5.5.",
    "Aceleração": "Acceleration",
    "Acima do limite superior de Long (0,25 % de H): ": "Above the upper limit of Long (0.25% of H): ",
    "Acumulados": "Cumulative",
    "Ajuste fino do alinhamento": "Fine alignment adjustment",
    "Alarme": "Alarm",
    "Alarme ": "Alarm ",
    "Alarme -": "Alarm -",
    "Alarme 25%": "Alarm 25%",
    "Alarme contenção (40)": "Wall alarm (40)",
    "Alerta": "Alert",
    "Alerta ": "Alert ",
    "Alerta -": "Alert -",
    "Alerta 15%": "Alert 15%",
    "Alertas c/ sinal antes": "Alerts w/ earlier signal",
    "Alertas c/ sinal ≤4 sem.": "Alerts w/ signal ≤4 wk",
    "Alinhado": "Aligned",
    "Alinhar a planta com os alvos (sobreposição)": "Align the plan with the targets (overlay)",
    "Altura da cortina H (m)": "Wall height H (m)",
    "Altura relativa (Z)": "Relative height (Z)",
    "Alvo": "Target",
    "Alvo %{text}<br>Desl. h: %{customdata:.1f} mm<extra></extra>": "Target %{text}<br>Hor. disp.: %{customdata:.1f} mm<extra></extra>",
    "Alvo %{text}<br>Desl. h: %{customdata[0]:.1f} mm<br>Estado: %{customdata[1]}<br>%{customdata[2]}<extra>": "Target %{text}<br>Hor. disp.: %{customdata[0]:.1f} mm<br>State: %{customdata[1]}<br>%{customdata[2]}<extra>",
    "Alvo 1": "Target 1",
    "Alvo 2": "Target 2",
    "Alvo a analisar (velocidade)": "Target to analyse (velocity)",
    "Alvo a sobrepor (mesma zona da célula)": "Target to overlay (same zone as the cell)",
    "Alvos": "Targets",
    "Alvos (2D)": "Targets (2D)",
    "Alvos A5-A8 -> A5b-A8b": "Targets A5-A8 -> A5b-A8b",
    "Alvos a incluir na matriz": "Targets to include in the matrix",
    "Alvos c/ sinal → alerta (%)": "Targets w/ signal → alert (%)",
    "Alvos em alerta ou alarme na última campanha:": "Targets in alert or alarm in the last campaign:",
    "Alvos monitorizados": "Monitored targets",
    "Alvos na campanha": "Targets in the campaign",
    "Alvos no espaço com a geometria da obra: contorno do recinto, edifícios vizinhos e vetores de deslocamento amplificados.": "Targets in space with the geometry of the works: site boundary, neighbouring buildings and amplified displacement vectors.",
    "Alvos s/ alerta c/ sinal": "Targets w/o alert w/ signal",
    "Alvos topogr.": "Survey targets",
    "Alvos topográficos": "Survey targets",
    "Alvos topográficos — evolução temporal": "Survey targets — evolution in time",
    "Alçado": "Wall section",
    "Alçado ": "Section ",
    "Alçado %{text}<br>H = %{customdata[0]:.1f}–%{customdata[1]:.1f} m<br>δ_h = %{y:.1f} mm<br>δ_h/H = %{customdata[2]:.2f} %<extra></extra>": "Section %{text}<br>H = %{customdata[0]:.1f}–%{customdata[1]:.1f} m<br>δ_h = %{y:.1f} mm<br>δ_h/H = %{customdata[2]:.2f} %<extra></extra>",
    "Alçado D-E — ancoragem DE (Piso -2, cota 12,45)": "Section D-E — anchor DE (Level -2, elev. 12.45)",
    "Alçado GH": "Section GH",
    "Alçado GH (A19–A42): segundo o alçado com a numeração das ancoragens (A26 no piso −1, cota 15,90). Troço poente do AB (A1–A18): nível inferido da cronologia e do comprimento livre.": "Section GH (A19–A42): according to the elevation drawing with the anchor numbering (A26 at level −1, elev. 15.90). Western stretch of AB (A1–A18): level inferred from the chronology and the free length.",
    "Alçado GH / A26": "Section GH / A26",
    "Alçado e data a confirmar": "Section and date to be confirmed",
    "Alçado e nível": "Section and level",
    "Alçado e nível de cada ancoragem": "Section and level of each anchor",
    "Alçados (prismas)": "Wall sections (prisms)",
    "Ambas": "Both",
    "Amplia a escala vertical para realçar o relevo. 1.0 = escala real.": "Increases the vertical scale to highlight the relief. 1.0 = true scale.",
    "Amplia só a escala vertical do relevo.": "Increases only the vertical scale of the relief.",
    "Amplificação do deslocamento": "Displacement amplification",
    "Ancoragem": "Anchor",
    "Ancoragem DE": "Anchor DE",
    "Ancoragens": "Anchors",
    "Ancoragens ": "Anchors ",
    "Ancoragens DE": "DE anchors",
    "Ancoragens VD Piso -1 (2 lados)": "Anchors level -1 distribution beam (2 sides)",
    "Ancoragens VD Piso -2 (2 lados)": "Anchors level -2 distribution beam (2 sides)",
    "Ancoragens VD Piso 1 (E8_78-E8_87)": "Anchors level 1 distribution beam (E8_78-E8_87)",
    "Ancoragens viga de coroamento Piso 1 (E8_51-E8_35)": "Capping beam anchors Level 1 (E8_51-E8_35)",
    "Ancoragens viga de coroamento Piso 2 (E8_78-E8_90)": "Capping beam anchors Level 2 (E8_78-E8_90)",
    "Anel (LM)": "Ring (LM)",
    "Antecedência": "Lead time",
    "Análise": "Analysis",
    "Análise · Correlações": "Analysis · Correlations",
    "Análise · Pressupostos": "Analysis · Assumptions",
    "Análise · Resultados": "Analysis · Results",
    "Análise · Terreno + Alvos 3D": "Analysis · Terrain + 3D targets",
    "Análise — análise de resultados": "Analysis — analysis of results",
    "As células de carga terminaram a 06/07/2026 (fim das leituras indicado pela obra, a confirmar o motivo); os restantes casos são instrumentos cuja série termina antes da data de estado.": "The load cells ended on 06/07/2026 (end of readings reported by the site, reason to be confirmed); the other cases are instruments whose series ends before the status date.",
    "As medidas com data são assinaladas nos gráficos dos fissurómetros do Edifício 1 (separador «Alvos (2D)»).": "The dated measures are marked on the charts of the Building 1 crack gauges («Targets (2D)» tab).",
    "As quatro (comparar)": "All four (compare)",
    "Associação inclinómetro-sondagem": "Inclinometer-borehole association",
    "Associação inclinómetro-sondagem inferida da sobreposição das plantas — a confirmar com a instrumentação.": "Inclinometer-borehole association inferred from overlaying the plans — to be confirmed with the instrumentation.",
    "Aterro": "Fill",
    "Aterro heterogéneo de origem não seletiva": "Heterogeneous fill of non-selective origin",
    "Até ": "Until ",
    "Até **": "Until **",
    "Auditoria: divergência em ": "Audit: discrepancy in ",
    "Auditoria: o estado recalculado pela app coincide com a coluna 'Estado' do Excel (calculada por fórmula) em ": "Audit: the state recalculated by the app coincides with the 'Estado' column of the Excel file (computed by formula) in ",
    "Automático": "Automatic",
    "Avaliação da carga da A26 pelo projetista; estudado um escoramento «pé de galinha» entre AB e GH (não executado)": "Assessment of the A26 load by the designer; a raking strut between AB and GH was studied (not built)",
    "Banda Laje Piso -2 (F1)": "Slab strip Level -2 (F1)",
    "Banda Laje Piso -2 (F2)": "Slab strip Level -2 (F2)",
    "Banda Laje Piso 1 (Fase 1)": "Slab strip Level 1 (Phase 1)",
    "Banda Laje Piso 1 (Fase 2)": "Slab strip Level 1 (Phase 2)",
    "Banda de laje ": "Slab strip ",
    "Banda do piso −2 contra o terreno só no alçado do Edifício 1; nas restantes, escavação e escoramento": "Level −2 slab strip against the ground only at the Building 1 section; elsewhere, excavation and bracing",
    "Bandas de laje": "Slab strips",
    "Blocagem": "Lock-off load",
    "Blocagem (kN)": "Lock-off load (kN)",
    "Blocagem indicada nos relatórios": "Lock-off load stated in the reports",
    "CONTRATUAL: ": "CONTRACTUAL: ",
    "Cada ponto é uma campanha: no eixo X a cota da água (interpolada à data de cada leitura do alvo), no eixo Y o deslocamento horizontal do alvo. Se os pontos alinham numa tendência, há correlação. O coeficiente r quantifica a forca da associação estatística — não prova a causa.": "Each point is a campaign: on the X axis the water level (interpolated to the date of each target reading), on the Y axis the horizontal displacement of the target. If the points line up in a trend, there is correlation. The coefficient r quantifies the strength of the statistical association — it does not prove the cause.",
    "Cadência (d)": "Interval (d)",
    "Caixa de escavação": "Excavation box",
    "Calcário": "Limestone",
    "Calcário (C1A)": "Limestone (C1A)",
    "Campanha": "Campaign",
    "Campanhas": "Campaigns",
    "Carga": "Load",
    "Carga ": "Load ",
    "Carga (kN)": "Load (kN)",
    "Carga (kN) — bloc. ": "Load (kN) — lock-off ",
    "Carga atual": "Current load",
    "Carga atual (kN)": "Current load (kN)",
    "Carga da célula em função da temperatura; a deriva ajustada compara-se com a do certificado de calibração.": "Cell load as a function of temperature; the fitted drift is compared with that of the calibration certificate.",
    "Carga fim (kN)": "Final load (kN)",
    "Carga nas ancoragens vs. blocagem, com limiares de alerta (15%) e alarme (25%).": "Anchor loads vs. lock-off load, with alert (15%) and alarm (25%) thresholds.",
    "Carrega o Excel na barra lateral para começar.": "Load the Excel file in the sidebar to start.",
    "Carrega um DXF para ver a planta. Sugestão: a planta de escavação do piso -1, -2 ou -3 mostra bem o contorno do recinto de contenção.": "Load a DXF file to see the plan. Suggestion: the excavation plan of level -1, -2 or -3 shows the outline of the retained area well.",
    "Carrega uma planta em DXF (por exemplo a planta de escavação e contenção de um piso). Podes ve-la sozinha, ou ALINHA-LA com os alvos topográficos indicando 2 pontos de referência — a app calcula a transformação e sobrepõe os alvos coloridos pelo deslocamento.": "Load a plan in DXF (for example the excavation and retaining structure plan of a level). You can view it alone, or ALIGN it with the survey targets by indicating 2 reference points — the app computes the transformation and overlays the targets coloured by displacement.",
    "Carregar manualmente": "Load manually",
    "Casas dos edifícios": "Building blocks",
    "Casas dos edifícios (ilustrativo)": "Building blocks (illustrative)",
    "Edifício 3": "Building 3",
    "Clough & O'Rourke — média ≈ 0,2 %": "Clough & O'Rourke — average ≈ 0.2%",
    "Edifício 3": "Building 3",
    "Coeficiente de Pearson r entre as séries temporais da instrumentação, calculado sobre o período comum a cada par (interpolação linear nas datas da primeira série). À esquerda, os valores ACUMULADOS; à direita, os INCREMENTOS entre campanhas (variação período-a-período), que medem se os instrumentos se MOVEM ao mesmo tempo, não apenas se cresceram ambos. Reproduz a Tabela 5.3 da dissertação.": "Pearson coefficient r between the instrumentation time series, computed over the period common to each pair (linear interpolation at the dates of the first series). On the left, the CUMULATIVE values; on the right, the INCREMENTS between campaigns (period-to-period change), which measure whether the instruments MOVE at the same time, not just whether both grew. Reproduces Table 5.3 of the dissertation.",
    "Colore os lados do recinto conforme o nível de movimento dos alvos próximos, por orientação. É QUALITATIVO e ilustrativo: o terreno e os alvos estão em referenciais diferentes, pelo que a associação e por lado, não por ponto exato.": "Colours the sides of the site according to the movement level of the nearby targets, by orientation. It is QUALITATIVE and illustrative: the terrain and the targets are in different reference systems, so the association is by side, not by exact point.",
    "Colorir alçados por movimento (ilustrativo)": "Colour wall sections by movement (illustrative)",
    "Colunas litológicas das sondagens, ensaios SPT em profundidade e zonamento geotécnico (ZG1-ZG6).": "Lithological columns of the boreholes, SPT tests with depth and geotechnical zoning (ZG1-ZG6).",
    "Como ler: para cada sondagem, a barra colorida à esquerda é a litologia; a curva à direita é o SPT (escala 0–60, com a linha pontilhada na nega). Onde o SPT sobe, o grés está mais consolidado — e e ai que a rigidez do maciço aumenta. Profundidades de cada zona ZG NÃO estão definidas no relatório por ponto; a cor do SPT é a zona PROVÁVEL pelo valor de N.": "How to read: for each borehole, the coloured bar on the left is the lithology; the curve on the right is the SPT (scale 0–60, with the dotted line at refusal). Where the SPT rises, the sandstone is more consolidated — and that is where the stiffness of the rock mass increases. The depths of each ZG zone are NOT defined in the report for each point; the SPT colour is the PROBABLE zone from the N value.",
    "Como obter o X,Y no desenho: abre o DXF no AutoCAD ou num visualizador, aponta o cursor ao sítio onde o alvo esta, e le as coordenadas. Quanto mais afastados os 2 alvos, melhor.": "How to get X,Y in the drawing: open the DXF in AutoCAD or a viewer, point the cursor at the place where the target is, and read the coordinates. The further apart the 2 targets, the better.",
    "Comparação": "Comparison",
    "Comparação impactado vs. contratual": "Impacted vs. contractual comparison",
    "Componente": "Component",
    "Confere": "Matches",
    "Confirmada": "Confirmed",
    "Confirmada nos dados": "Confirmed in the data",
    "Contenção": "Retaining wall",
    "Contenção 17 m — Alçado ": "Wall 17 m — Section ",
    "Contenção 24 m — Alçado ": "Wall 24 m — Section ",
    "Contenção periférica": "Peripheral retaining structure",
    "Contenção — Alçado ": "Wall — Section ",
    "Contorno do recinto": "Site boundary",
    "Contorno do recinto (envolvente)": "Site boundary (envelope)",
    "Contratual": "Contractual",
    "Corpo central/nascente (estacas E8_2–E8_44, viga de coroamento)": "Central/east body (piles E8_2–E8_44, capping beam)",
    "Correlação — ": "Correlation — ",
    "Correlações": "Correlations",
    "Correlações — interação entre instrumentos": "Correlations — interaction between instruments",
    "Cortina sob o Edifício 1 — Piso -1 (cota 15,90), junto a A26": "Wall below Building 1 — Level -1 (elev. 15.90), next to A26",
    "Cortina sob o Edifício 1 — Piso -2 (cota 12,45), perto da A26": "Wall below Building 1 — Level -2 (elev. 12.45), near A26",
    "Cota": "Elevation",
    "Cota (m)": "Elevation (m)",
    "Cota da boca (m)": "Collar elevation (m)",
    "Cota da água ": "Water level ",
    "Cota da água (PZ)": "Water level (PZ)",
    "Cota da água (m)": "Water level (m)",
    "Cota da água calculada (m)": "Computed water level (m)",
    "Cota escavada na frente (m)": "Excavated level at the front (m)",
    "Cota: %{intensity:.1f} m<extra></extra>": "Elevation: %{intensity:.1f} m<extra></extra>",
    "Cotas de projeto: escavação e contenção periférica (JETsj, PRO/2023/368). NF de repouso: piezómetros das sondagens (ENGGEO, Quadro III, 24/11/2022).": "Design elevations: excavation and peripheral retaining structure (JETsj, PRO/2023/368). Rest water table: borehole piezometers (ENGGEO, Table III, 24/11/2022).",
    "Cotas dos pisos e fundo de escavação": "Floor levels and excavation bottom",
    "Critério": "Criterion",
    "Critério aplicado: **": "Criterion applied: **",
    "Critério de alerta e alarme de cada alçado": "Alert and alarm criterion of each wall section",
    "Cronograma a mostrar": "Schedule to display",
    "Cronograma da obra — plano contratual e plano impactado": "Works schedule — contractual and impacted schedules",
    "Cruza a CARGA medida numa célula de carga (ancoragem) com o MOVIMENTO de um alvo, no mesmo eixo de tempo. Só faz sentido cruzar célula e alvo do MESMO alçado/zona. Critério das células: +15% (alerta), +25% (alarme) sobre a blocagem.": "Crosses the LOAD measured in a load cell (anchor) with the MOVEMENT of a target, on the same time axis. It only makes sense to cross a cell and a target of the SAME wall section/zone. Cell criterion: +15% (alert), +25% (alarm) over the lock-off load.",
    "Cruza o movimento de um alvo com a COTA a que a escavação chegou NA FRENTE desse alvo (plano de trabalhos impactado, por zona de estacas). A cota (castanho) desce em degraus a medida que a frente aprofunda; ve-se a deformação (vermelho) crescer com o aprofundamento local — e não com a escavação do lado oposto.": "Crosses the movement of a target with the LEVEL reached by the excavation AT THE FRONT of that target (impacted schedule, by pile zone). The level (brown) goes down in steps as the front deepens; the deformation (red) grows with the local deepening — and not with the excavation on the opposite side.",
    "Curvas de nível": "Contour lines",
    "Curvas de nível do levantamento, que dao a leitura do relevo (alçados).": "Contour lines of the survey, which give the reading of the relief (wall sections).",
    "Célula": "Cell",
    "Célula de carga": "Load cell",
    "Células carga": "Load cells",
    "Células de carga": "Load cells",
    "Dados": "Data",
    "Dados · Alvos (2D)": "Data · Targets (2D)",
    "Dados · Células de carga": "Data · Load cells",
    "Dados · Geologia": "Data · Geology",
    "Dados · Inclinómetros": "Data · Inclinometers",
    "Dados · Obra": "Data · Works",
    "Dados · Piezómetros": "Data · Piezometers",
    "Dados · Planta (DXF)": "Data · Plan (DXF)",
    "Dados — instrumentação, geologia e planeamento": "Data — instrumentation, geology and planning",
    "Dados: ilustração dos resultados (cada instrumento, a geologia e o planeamento), sem análise. Análise: leitura dos resultados e correlação entre os dados dos vários instrumentos.": "Data: illustration of the results (each instrument, the geology and the planning), without analysis. Analysis: reading of the results and correlation between the data of the various instruments.",
    "Data": "Date",
    "Data a confirmar": "Date to be confirmed",
    "Data da análise (replay)": "Analysis date (replay)",
    "Data de estado (última leitura global): **": "Status date (last overall reading): **",
    "Datas do plano de trabalhos impactado (o plano contratual com três suspensões de frentes, mantendo as durações contratuais); não há registo diário de obra. As ancoragens usam as datas dos registos de execução (folha «Ancoragens»), que diferem do plano em semanas.": "Dates of the impacted schedule (the contractual schedule with three suspensions of fronts, keeping the contractual durations); there is no daily site record. The anchors use the dates of the execution records («Ancoragens» sheet), which differ from the schedule by weeks.",
    "Deformada do ": "Deformed shape of ",
    "Deriva ≈ ": "Drift ≈ ",
    "Descrição": "Description",
    "Desenha cada edifício vizinho como uma casa simples (bloco + telhado). Posição real dos alvos; forma esquemática — não representa a geometria real do edifício.": "Draws each neighbouring building as a simple house (block + roof). Real position of the targets; schematic shape — it does not represent the real geometry of the building.",
    "Desenha o volume escavado abaixo do contorno, com a PROFUNDIDADE real de escavação do projeto (coroamento-fundo = 16,3 m). É uma distância, não uma cota absoluta — os alvos e o projeto usam referenciais de cota diferentes.": "Draws the excavated volume below the boundary, with the real excavation DEPTH of the design (capping beam to bottom = 16.3 m). It is a distance, not an absolute elevation — the targets and the design use different elevation references.",
    "Desenha, de forma ESQUEMÁTICA e ILUSTRATIVA, a viga de coroamento, as bandas de laje (as cotas reais dos pisos) e os escoramentos de canto. A geometria de projeto destes elementos não consta dos dados — a forma é assumida, só as cotas das lajes/viga são reais. Os escoramentos são totalmente ilustrativos.": "Draws, in a SCHEMATIC and ILLUSTRATIVE way, the capping beam, the slab strips (the real floor elevations) and the corner struts. The design geometry of these elements is not in the data — the shape is assumed, only the slab/beam elevations are real. The struts are entirely illustrative.",
    "Desl.": "Disp.",
    "Desl. H ": "Hor. disp. ",
    "Desl. H (mm)": "Hor. disp. (mm)",
    "Desl. H máximo (mm)": "Max. hor. disp. (mm)",
    "Desl. V ΔZ (mm)": "Vert. disp. ΔZ (mm)",
    "Desl. acumulado total (mm)": "Total cumulative disp. (mm)",
    "Desl. h (mm)": "Hor. disp. (mm)",
    "Desl. horizontal ": "Horizontal disp. ",
    "Desl. horizontal (mm)": "Horizontal disp. (mm)",
    "Desl. horizontal máx. (mm)": "Max. horizontal disp. (mm)",
    "Desl. máx (mm)": "Max. disp. (mm)",
    "Desl. máx. ": "Max. disp. ",
    "Desloc. horizontal acumulado (mm)": "Cumulative horizontal disp. (mm)",
    "Deslocamento": "Displacement",
    "Deslocamento (mm)": "Displacement (mm)",
    "Deslocamento acumulado (mm)": "Cumulative displacement (mm)",
    "Deslocamento acumulado ao longo da profundidade. Base fixa.": "Cumulative displacement along depth. Fixed base.",
    "Deslocamento horizontal OBSERVADO vs. ESTIMATIVA DE PROJETO (memória descritiva JETsj: ~20 mm na cortina poente/norte, ~10 mm na nascente/sul — valor de CÁLCULO para a última fase de escavação), atribuida pela frente de escavação adjacente. Para os edifícios vizinhos a comparação e INDIRETA: a estimativa refere-se a cortina, não ao edifício.": "OBSERVED horizontal displacement vs. DESIGN ESTIMATE (JETsj design report: ~20 mm on the west/north wall, ~10 mm on the east/south wall — CALCULATION value for the last excavation stage), assigned by the adjacent excavation front. For the neighbouring buildings the comparison is INDIRECT: the estimate refers to the wall, not to the building.",
    "Deslocamento horizontal máximo de cada alçado (máximo dos prismas em todas as campanhas) e de cada inclinómetro (máximo do perfil), em função da altura da cortina H (Tabela 3.7), face aos referenciais empíricos e às estimativas do projeto. As linhas a tracejado são δ_h = p%·H; a sombra marca as faixas de Long e de Yoo. Reproduz a Figura 5.3 da dissertação.": "Maximum horizontal displacement of each wall section (maximum of the prisms over all campaigns) and of each inclinometer (maximum of the profile), as a function of the wall height H (Table 3.7), against the empirical references and the design estimates. The dashed lines are δ_h = p%·H; the shading marks the ranges of Long and of Yoo. Reproduces Figure 5.3 of the dissertation.",
    "Deslocamentos milimétricos sobre coordenadas em metros; amplia-se para se verem.": "Millimetric displacements on coordinates in metres; amplified to be visible.",
    "Destacar alarmes/alertas": "Highlight alarms/alerts",
    "Detalhe do eixo temporal": "Time axis detail",
    "Deteção de precursores": "Precursor detection",
    "Dias corridos": "Calendar days",
    "Dias desde a última": "Days since the last",
    "Dias entre o fim da injeção e o tensionamento": "Days between the end of grouting and stressing",
    "Dias fim da injeção → tensionamento": "Days end of grouting → stressing",
    "Dias úteis": "Working days",
    "Dias úteis contados de segunda a sexta. O impacto 3 (Edifício 2) não está ligado a tarefas seguintes no plano, e as estacas E6_21–E6_27 aparecem executadas a partir de 07/08/2025, antes do fim dessa suspensão (21/08/2025) — incoerência a confirmar com o empreiteiro. Por manter as durações contratuais, o plano não mede o ritmo efetivo de cada atividade.": "Working days counted from Monday to Friday. Impact 3 (Building 2) is not linked to subsequent tasks in the schedule, and piles E6_21–E6_27 appear executed from 07/08/2025, before the end of that suspension (21/08/2025) — inconsistency to be confirmed with the contractor. As it keeps the contractual durations, the schedule does not measure the actual pace of each activity.",
    "Diferença (mm)": "Difference (mm)",
    "Distingue a fachada frontal (A1-A4, exposta a escavação) da lateral (A5-A8, ao mar).": "Distinguishes the front façade (A1-A4, exposed to the excavation) from the side one (A5-A8, facing the sea).",
    "Distinguidas frontal (A1-A4) e lateral (A5-A8) pela posição em planta.": "Front (A1-A4) and side (A5-A8) distinguished by position in plan.",
    "Documentado nos relatórios; motivo da classificação de MNO e OP a confirmar": "Documented in the reports; reason for the classification of MNO and OP to be confirmed",
    "Dois perfis HEB300 contra a viga de distribuição, apoiados na sapata do Núcleo 1 (após o fundo)": "Two HEB300 sections against the distribution beam, bearing on the footing of Core 1 (after the bottom)",
    "Dos ": "Of the ",
    "Edifício": "Building",
    "Edifício / elemento": "Building / element",
    "Edifício 1": "Building 1",
    "Edifício 1": "Building 1",
    "Edifício adjacente (15/25 · 10/20)": "Adjacent building (15/25 · 10/20)",
    "Edifício em frente": "Building opposite",
    "Elemento com maior movimentação: «": "Element with the largest movement: «",
    "Elementos de contenção (esquemático)": "Retaining elements (schematic)",
    "Em ALARME": "In ALARM",
    "Em ALERTA": "In ALERT",
    "Em alarme": "In alarm",
    "Em alerta": "In alert",
    "Em cada fase: barra azul = plano impactado; barra cinzenta = plano contratual. Barras vermelhas = suspensões (impactos).": "In each phase: blue bar = impacted schedule; grey bar = contractual schedule. Red bars = suspensions (impacts).",
    "Entrada em alerta": "Entry into alert",
    "Escala desenho->alvos": "Scale drawing->targets",
    "Escavação": "Excavation",
    "Escavação + bandas de laje + ancoragens": "Excavation + slab strips + anchors",
    "Escavação até cota ": "Excavation to elev. ",
    "Escavação completa": "Complete excavation",
    "Escavação completa (fundo ": "Complete excavation (bottom ",
    "Escavação para as bandas dos pisos 2 e 1 até 4,0 m (previstos 2,5 m)": "Excavation for the level 2 and 1 strips to 4.0 m (2.5 m planned)",
    "Escolhe dois alvos diferentes.": "Choose two different targets.",
    "Escolhe pelo menos uma layer para desenhar.": "Choose at least one layer to draw.",
    "Escora metálica provisória (horizontal)<br>posição esquemática<extra></extra>": "Temporary steel strut (horizontal)<br>schematic position<extra></extra>",
    "Escoras metálicas F-G/G-H (viga de coroamento)": "Steel struts F-G/G-H (capping beam)",
    "Escoras metálicas de canto (Piso -2)": "Corner steel struts (Level -2)",
    "Escoras metálicas de canto (alinhamentos B/C/D/E)": "Corner steel struts (lines B/C/D/E)",
    "Escoras metálicas na VD H-I/I-J": "Steel struts on distribution beam H-I/I-J",
    "Escoras metálicas na banda de laje do Piso 2": "Steel struts on the Level 2 slab strip",
    "Escoras metálicas provisórias (esquemático)": "Temporary steel struts (schematic)",
    "Espelhar horizontal": "Mirror horizontally",
    "Espelhar vertical": "Mirror vertically",
    "Esquema a partir das coordenadas reais dos alvos (posições relativas fieis; escala esquemática). O alvo selecionado aparece em estrela. Para uma foto real, coloca a imagem em 'fotos_alvos/": "Sketch from the real coordinates of the targets (relative positions faithful; schematic scale). The selected target appears as a star. For a real photo, put the image in 'fotos_alvos/",
    "Esta aplicação apoia a **análise inversa** (*back-analysis*) da contenção periférica da reformulação do **Hotel Eden**, no Monte Estoril. A obra compreende uma escavação profunda — fundo à cota 4,55 m, cerca de 16 m abaixo do piso 1 (cota 20,95 m) e mais nas zonas em que o terreno natural é mais alto — executada ao abrigo de uma cortina de estacas ancorada e travada por bandas de laje, num contexto exigente: confina com **edifícios sensíveis**, em particular três edifícios vizinhos instrumentados (Edifícios 1, 2 e 3), e desenvolve-se **abaixo do nível freático** de repouso. A instrumentação instalada — inclinómetros, alvos topográficos, células de carga e piezómetros — permite acompanhar, ao longo do tempo, os deslocamentos induzidos pela escavação, tanto na própria contenção como nos edifícios vizinhos.": "This application supports the **back-analysis** of the peripheral retaining structure of the redevelopment of **Hotel Eden**, in Monte Estoril. The works comprise a deep excavation — bottom at elevation 4.55 m, about 16 m below level 1 (elevation 20.95 m) and more where the natural ground is higher — carried out within a pile wall anchored and braced by slab strips, in a demanding setting: it adjoins **sensitive buildings**, in particular three instrumented neighbouring buildings (Buildings 1, 2 and 3), and extends **below the rest water table**. The installed instrumentation — inclinometers, survey targets, load cells and piezometers — makes it possible to follow, over time, the displacements induced by the excavation, both in the retaining structure itself and in the neighbouring buildings.",
    "Estacas Central e Nascente": "Central and east piles",
    "Estacas junto ao Edifício 2": "Piles next to Building 2",
    "Estacas E6_C_17 a E6_C_2 e E8_51 (alinhamento coincidente com o caneiro)": "Piles E6_C_17 to E6_C_2 and E8_51 (alignment coinciding with the culvert)",
    "Estacas Poente e Norte": "West and north piles",
    "Estacas periféricas (Edifício 2)": "Perimeter piles (Building 2)",
    "Estado": "State",
    "Estado calculado": "Computed state",
    "Estado de confirmação": "Confirmation status",
    "Estado dos alvos": "Target states",
    "Estas a ver **": "You are viewing **",
    "Este alçado não consta da lista de critérios dos relatórios de alvos; o critério de 17 m é assumido.": "This section is not in the list of criteria of the target reports; the 17 m criterion is assumed.",
    "Estimativa de projeto": "Design estimate",
    "Estrutura pisos -3 a -1": "Structure levels -3 to -1",
    "Estrutura pisos 1 e 2 (em curso)": "Structure levels 1 and 2 (in progress)",
    "Etiqueta com o nome de cada edifício, flutuando sobre os seus alvos.": "Label with the name of each building, floating over its targets.",
    "Evolução da cota da água subterrânea.": "Evolution of the groundwater level.",
    "Evolução do deslocamento": "Displacement evolution",
    "Exagero vertical": "Vertical exaggeration",
    "Exagero vertical do terreno": "Vertical exaggeration of the terrain",
    "Excel de instrumentação": "Instrumentation Excel file",
    "Excesso (%)": "Excess (%)",
    "Fachada": "Façade",
    "Fachada": "Façade",
    "Fachadas do Edifício 1": "Building 1 façades",
    "Falta a biblioteca 'ezdxf'. Acrescenta 'ezdxf' ao requirements.txt (e 'pip install ezdxf' se correres localmente) para ativar a leitura de desenhos.": "The 'ezdxf' library is missing. Add 'ezdxf' to requirements.txt (and 'pip install ezdxf' if running locally) to enable reading drawings.",
    "Fase de escavação a representar": "Excavation stage to display",
    "Faseamento da obra": "Construction phases",
    "Fases da obra": "Construction phases",
    "Fases da obra (datas do plano de trabalhos impactado): ": "Construction phases (dates of the impacted schedule): ",
    "Fases de escavação (pisos)": "Excavation stages (levels)",
    "Fator de aceleração (x)": "Acceleration factor (x)",
    "Ficheiro DXF": "DXF file",
    "Ficheiro de dados": "Data file",
    "Ficheiro fonte": "Source file",
    "Fim da injeção": "End of grouting",
    "Fissurometros": "Crack gauges",
    "Fissurómetro": "Crack gauge",
    "Fissurómetros": "Crack gauges",
    "Fissurómetros do Edifício 1": "Building 1 crack gauges",
    "Fissurómetros em cinco fissuras do edifício (zeragem)": "Crack gauges on five cracks of the building (zeroing)",
    "Fonte": "Source",
    "Foto real do relatório, se disponível; caso contrário um alçado esquemático a partir das coordenadas.": "Real photo from the report, if available; otherwise a schematic elevation from the coordinates.",
    "Frente": "Front",
    "Frente escavação": "Excavation front",
    "Frente junto ao Edifício 2; demolição do muro antigo (adjudicação 28/07, desmonte 29/07 a 21/08)": "Front next to Building 2; demolition of the old wall (award 28/07, dismantling 29/07 to 21/08)",
    "Fundação e laje de fundo": "Foundation and bottom slab",
    "Funde o modelo real do terreno (levantamento topográfico, com curvas de nível e volume de escavação) com os alvos e o seu movimento. Os alvos estão num referencial diferente do terreno, por isso são POUSADOS sobre a superfície por alinhamento aproximado (caixa-envolvente) — a posição é ILUSTRATIVA, para ler a tendência do movimento sobre o relevo. A medição precisa está no separador Alvos (2D).": "Merges the real terrain model (topographic survey, with contour lines and excavation volume) with the targets and their movement. The targets are in a reference system different from the terrain, so they are PLACED on the surface by approximate alignment (bounding box) — the position is ILLUSTRATIVE, to read the trend of the movement over the relief. The precise measurement is in the Targets (2D) tab.",
    "Fundo": "Bottom",
    "Fundo de escavação": "Excavation bottom",
    "Fundo de escavação (": "Excavation bottom (",
    "Fundo de escavação (contratual)": "Excavation bottom (contractual)",
    "Fundo de escavação (impactado)": "Excavation bottom (impacted)",
    "Fundo de escavação (−": "Excavation bottom (−",
    "Furação e selagem (adotada)": "Drilling and grouting (adopted)",
    "Furação/selagem": "Drilling/grouting",
    "GH confirmado; AB a confirmar": "GH confirmed; AB to be confirmed",
    "Geologia": "Geology",
    "Geologia do terreno (Relatório ENGGEO, proc. 220216)": "Ground geology (ENGGEO report, proc. 220216)",
    "Grandeza medida e convenção de sinal a confirmar": "Measured quantity and sign convention to be confirmed",
    "Granularidade do eixo temporal": "Time axis granularity",
    "Grupo": "Group",
    "Gráficos de interação entre instrumentos: deformação vs. água vs. escavação vs. faseamento; deformação vs. carga na ancoragem; e a correlação direta (XY) com coeficiente r.": "Charts of interaction between instruments: deformation vs. water vs. excavation vs. construction phases; deformation vs. anchor load; and the direct correlation (XY) with coefficient r.",
    "Gráficos que cruzam várias séries da instrumentação para mostrar como os movimentos acentuados acompanham os outros dados (água, escavação, faseamento, carga nas ancoragens). Demonstram a correlação entre os dados; a interpretação das causas e do autor.": "Charts that cross several instrumentation series to show how sharp movements follow the other data (water, excavation, construction phases, anchor loads). They show the correlation between the data; the interpretation of the causes is the author’s.",
    "Grés": "Sandstone",
    "Grés (C1As)": "Sandstone (C1As)",
    "Grés consolidado, SPT>=60, RQD 76-100%": "Consolidated sandstone, SPT>=60, RQD 76-100%",
    "Grés irreg. consolidado a consolidado, SPT>=60, RQD 45-75%": "Irregularly consolidated to consolidated sandstone, SPT>=60, RQD 45-75%",
    "Grés pouco consolidado, SPT 11-30, RQD 0%": "Poorly consolidated sandstone, SPT 11-30, RQD 0%",
    "Grés pouco consolidado, SPT 31-56, RQD 0%": "Poorly consolidated sandstone, SPT 31-56, RQD 0%",
    "Grés/calcário irreg. consolidado, SPT>=60, RQD 0-25%": "Irregularly consolidated sandstone/limestone, SPT>=60, RQD 0-25%",
    "Há duas componentes, selecionáveis na barra lateral: **Dados** (ilustração dos resultados de cada instrumento, geologia e planeamento, sem análise) e **Análise** (leitura dos resultados e correlação entre instrumentos). Usa os separadores no topo para navegar dentro de cada componente.": "There are two components, selectable in the sidebar: **Data** (illustration of the results of each instrument, geology and planning, without analysis) and **Analysis** (reading of the results and correlation between instruments). Use the tabs at the top to navigate within each component.",
    "I1<->SC6, I2<->SC9, I3<->SC6, inferido por sobreposição das plantas de prospeção e de instrumentação. A sobreposição gráfica é feita por PROFUNDIDADE, não por cota (cota da boca dos tubos desconhecida).": "I1<->SC6, I2<->SC9, I3<->SC6, inferred by overlaying the investigation and instrumentation plans. The graphic overlay is made by DEPTH, not by elevation (collar elevation of the tubes unknown).",
    "IMPACTADO: ": "IMPACTED: ",
    "Identificar edifícios": "Identify buildings",
    "Imagens das peças desenhadas do projeto de contenção periférica (JETsj, pranchas EDN-JET-ZZ-ZZ-DR-U-0002 a 0005). São a representação AUTORITATIVA e rigorosa da solução — cortina de estacas, bandas de laje por cota, ancoragens, escoras metálicas e faseamento. O 3D interativo acima é uma leitura do movimento dos alvos; estas vistas são a geometria de projeto.": "Images of the drawings of the peripheral retaining structure design (JETsj, sheets EDN-JET-ZZ-ZZ-DR-U-0002 to 0005). They are the AUTHORITATIVE and rigorous representation of the solution — pile wall, slab strips by elevation, anchors, steel struts and phasing. The interactive 3D above is a reading of the movement of the targets; these views are the design geometry.",
    "Imagens das vistas 3D do projeto não encontradas (pasta 'projeto_vistas/'). Verifica que foi publicada junto com a app.": "3D views of the design not found ('projeto_vistas/' folder). Check that it was published with the app.",
    "Imagens de localização não encontradas (pasta 'fotos_celulas/'). Verifica que foram publicadas com a app.": "Location images not found ('fotos_celulas/' folder). Check that they were published with the app.",
    "Impactado": "Impacted",
    "Impacto": "Impact",
    "Impacto 1 — caneiro": "Impact 1 — culvert",
    "Impacto 2 (as estacas seguem as da frente norte)": "Impact 2 (the piles follow those of the north front)",
    "Impacto 2 (via bandas de laje dos pisos 2 e 1)": "Impact 2 (via the level 2 and 1 slab strips)",
    "Impacto 2 — cabos BT/MT": "Impact 2 — LV/MV cables",
    "Impacto 3 — Edifício 2": "Impact 3 — Building 2",
    "Inclinómetro": "Inclinometer",
    "Inclinómetro (deformada no tempo)": "Inclinometer (deformed shape over time)",
    "Inclinómetros": "Inclinometers",
    "Inclinómetros / perfis": "Inclinometers / profiles",
    "Inclinómetros / resumo": "Inclinometers / summary",
    "Incrementos": "Increments",
    "Injeção": "Grouting",
    "Injeção repetitiva (IR) em vez da IRS do projeto, com o acordo do projetista": "Repeated grouting (IR) instead of the IRS of the design, with the agreement of the designer",
    "Instala 'ezdxf' para ativar a leitura de plantas DXF.": "Install 'ezdxf' to enable reading DXF plans.",
    "Instala 'scipy' para ativar a superfície 3D interpolada.": "Install 'scipy' to enable the interpolated 3D surface.",
    "Instrumentação": "Instrumentation",
    "Instrumento": "Instrument",
    "Instrumentos": "Instruments",
    "Integral (+36 d.u.), pela sequência das estacas da frente poente e do corpo central/nascente": "Full (+36 w.d.), through the sequence of the piles of the west front and of the central/east body",
    "Início": "Home",
    "Início da injeção": "Start of grouting",
    "Janela de velocidade: ": "Velocity window: ",
    "Janela temporal": "Time window",
    "Junto ao Edifício 2 (estacas periféricas, viga, muro M03, banda do piso 1)": "Next to Building 2 (perimeter piles, beam, wall M03, level 1 strip)",
    "Lado E": "Side E",
    "Lado N": "Side N",
    "Lado O": "Side W",
    "Lado S": "Side S",
    "Lateral (mar)": "Side (sea)",
    "Layers a desenhar": "Layers to draw",
    "Legenda": "Legend",
    "Leitura / geologia": "Reading / geology",
    "Leitura acumulada (mm)": "Cumulative reading (mm)",
    "Leitura de plantas de escavação em DXF, com opção de alinhamento aos alvos por pontos de referência.": "Reading of excavation plans in DXF, with the option of aligning them with the targets by reference points.",
    "Leitura integrada por sondagem: a coluna litológica, os ensaios SPT e o zonamento geotécnico partilham o eixo de profundidade, para se lerem em conjunto. Sob ~0,5 m de aterro, o terreno e essencialmente grés dos 'Grés Superiores' (C1As), com calcário (C1A) apenas no fundo do SC8. A resistência não é uniforme: há troços de menor N(SPT) intercalados com nega (por exemplo, N 19-47 no SC8 entre as cotas ~24 e ~18 e N 22-38 no SC6 entre as cotas ~14 e ~6,6) e o RQD é nulo em grande parte das sondagens. Compare sondagens por COTA, não por profundidade (as bocas estão a cotas diferentes).": "Integrated reading by borehole: the lithological column, the SPT tests and the geotechnical zoning share the depth axis, to be read together. Under ~0.5 m of fill, the ground is essentially sandstone of the 'Upper Sandstones' (C1As), with limestone (C1A) only at the bottom of SC8. The strength is not uniform: there are stretches of lower N(SPT) interbedded with refusal (for example, N 19-47 in SC8 between elevations ~24 and ~18 and N 22-38 in SC6 between elevations ~14 and ~6.6) and the RQD is zero in much of the boreholes. Compare boreholes by ELEVATION, not by depth (the collars are at different elevations).",
    "Leitura qualitativa (não fronteiras reais)": "Qualitative reading (not real boundaries)",
    "Leituras": "Readings",
    "Leituras acumuladas em mm desde 19/01/2026, tal como recebidas.": "Cumulative readings in mm since 19/01/2026, as received.",
    "Leituras das células de carga duas vezes por semana": "Load cell readings twice a week",
    "Leituras dos alvos duas vezes por semana": "Target readings twice a week",
    "Leituras nulas de A5-A8 após 20/10/2025": "Null readings of A5-A8 after 20/10/2025",
    "Leituras por cordas no edifício": "Readings by chords on the building",
    "Limiar": "Threshold",
    "Limiar de velocidade (mm/dia)": "Velocity threshold (mm/day)",
    "Linhas verticais nas datas do plano de trabalhos impactado em que a escavação atingiu cada cota NA FRENTE do instrumento escolhido (plano de trabalhos impactado, por zona de estacas).": "Vertical lines at the dates of the impacted schedule on which the excavation reached each level AT THE FRONT of the chosen instrument (impacted schedule, by pile zone).",
    "Litologia, NF e SPT da sondagem ": "Lithology, water table and SPT of borehole ",
    "Litologia: ": "Lithology: ",
    "Localização": "Location",
    "Localização dos alvos — ": "Location of the targets — ",
    "Localização: ": "Location: ",
    "Long (2001) — solos rijos (0,05–0,25 %)": "Long (2001) — stiff soils (0.05–0.25%)",
    "Lote": "Lot",
    "Magnitude, distribuição espacial, ritmo e profundidade das movimentações, e observado vs. previsto. Ilustração factual dos resultados, para a escrita da tese.": "Magnitude, spatial distribution, rate and depth of the movements, and observed vs. predicted. Factual illustration of the results, for writing the dissertation.",
    "Maior deslocamento horizontal: alvo ": "Largest horizontal displacement: target ",
    "Maior excedência: «": "Largest exceedance: «",
    "Marca a vermelho os alvos e setas em alarme e a laranja os em alerta, segundo os critérios oficiais recalculados.": "Marks in red the targets and arrows in alarm and in orange those in alert, according to the recalculated official criteria.",
    "Marca dentro da caixa os níveis dos pisos (-1 a -4) como planos, a partir das cotas de projeto. São distâncias abaixo do coroamento, invariantes ao referencial.": "Marks inside the box the floor levels (-1 to -4) as planes, from the design elevations. They are distances below the capping beam, independent of the reference system.",
    "Marcas mais finas ajudam a ler as campanhas (~8 em 8 dias).": "Finer ticks help to read the campaigns (~every 8 days).",
    "Marcos de escavação da zona (datas do planeamento)": "Excavation milestones of the zone (planning dates)",
    "Medida": "Measure",
    "Mensal": "Monthly",
    "Modelo 3D do projeto — localização da célula A26; Edifício 1 assinalado": "3D design model — location of cell A26; Building 1 marked",
    "Modelo de terreno não disponível (falta terreno_mdt.json).": "Terrain model not available (terreno_mdt.json missing).",
    "Monitorização": "Monitoring",
    "Mostra a escavação até a cota do piso correspondente. As cotas são as do projeto; a escavação real progride por fases.": "Shows the excavation down to the elevation of the corresponding level. The elevations are those of the design; the real excavation progresses by stages.",
    "Mostra a litologia, o nível freático e o perfil SPT de uma sondagem no mesmo eixo de profundidade, para relacionar a deformação com a resistência do terreno.": "Shows the lithology, the water table and the SPT profile of a borehole on the same depth axis, to relate the deformation to the strength of the ground.",
    "Mostrar localização dos alvos no edifício": "Show the location of the targets on the building",
    "Mostrar volume de escavação": "Show excavation volume",
    "Movimento dos alvos no espaço, com a geometria da obra": "Movement of the targets in space, with the geometry of the works",
    "Máx vs bloc. (%)": "Max vs lock-off (%)",
    "Máx.": "Max.",
    "Máx. desloc. acumulado total (mm)": "Max. total cumulative disp. (mm)",
    "Máximo": "Maximum",
    "Máximo global": "Overall maximum",
    "Máximo global vs. profundidade fixa.": "Overall maximum vs. fixed depth.",
    "Média": "Mean",
    "Mínimo": "Minimum",
    "NF de repouso (2022)": "Rest water table (2022)",
    "No Edifício 1, a tese obtém ≈ 0,2 em A1/A2 e 0,4–0,5 nos restantes alvos.": "At Building 1, the dissertation obtains ≈ 0.2 at A1/A2 and 0.4–0.5 at the other targets.",
    "Nascente": "East",
    "Nascente/Sul (escavação)": "East/South (excavation)",
    "Nascente/Sul — cortina do Edifício 2": "East/South — Building 2 wall",
    "Navegação": "Navigation",
    "Nega (N=60)": "Refusal (N=60)",
    "Nenhuma aceleração acima dos critérios.": "No acceleration above the criteria.",
    "Nenhuma: no plano, a suspensão não precede outras tarefas": "None: in the schedule, the suspension does not precede other tasks",
    "Norte": "North",
    "Norte — cortina do Edifício 3": "North — Building 3 wall",
    "Nos registos mostrados, **": "In the records shown, **",
    "Nota metodológica: a integração dos perfis inclinométricos assume a base fixa. As datas da sequência de obra provem do plano de trabalhos impactado (o contratual com três suspensões de frentes), não de um registo diário de obra. A posição dos alvos no 3D é aproximada — apoio visual, não modelo do maciço.": "Methodological note: the integration of the inclinometer profiles assumes a fixed base. The dates of the construction sequence come from the impacted schedule (the contractual one with three suspensions of fronts), not from a daily site record. The position of the targets in 3D is approximate — a visual aid, not a model of the ground.",
    "Nota metodológica: estas limitações não invalidam a análise — delimitam o seu alcance. A app apresenta medições e correlações; a interpretação das causas (escavação, furação das ancoragens, rebaixamento da água, rigidez do maciço) é feita pelo engenheiro e uma correlação temporal não demonstra, por si só, causalidade.": "Methodological note: these limitations do not invalidate the analysis — they define its scope. The app presents measurements and correlations; the interpretation of the causes (excavation, anchor drilling, drawdown of the water, stiffness of the ground) is made by the engineer, and a temporal correlation does not, on its own, demonstrate causation.",
    "Notas": "Notes",
    "Num só eixo de tempo: a deformação medida (esquerda), a cota da água subterrânea (direita), os marcos de escavação (plano impactado) DA ZONA do alvo/inclinómetro escolhido (E1-E..., no fundo) e as fases da obra (faixas de cor). Cada instrumento e cruzado com a escavação da frente que lhe fica em frente — não com a do lado oposto.": "On a single time axis: the measured deformation (left), the groundwater level (right), the excavation milestones (impacted schedule) OF THE ZONE of the chosen target/inclinometer (E1-E..., at the bottom) and the construction phases (colour bands). Each instrument is crossed with the excavation of the front facing it — not with that of the opposite side.",
    "Não": "No",
    "Não consegui calcular o alinhamento: ": "Could not compute the alignment: ",
    "Não consegui extrair a geometria. Detalhe: ": "Could not extract the geometry. Detail: ",
    "Não consegui ler as layers do DXF. Detalhe: ": "Could not read the DXF layers. Detail: ",
    "Não consegui ler o Excel. Detalhe: ": "Could not read the Excel file. Detail: ",
    "Não encontrei '": "Could not find '",
    "Não encontrei geometria nas layers escolhidas. Tenta outras.": "No geometry found in the chosen layers. Try others.",
    "Não foi possível alinhar alvos e terreno.": "Could not align targets and terrain.",
    "Nível freático": "Water table",
    "Nível freático de repouso (2022)": "Rest water table (2022)",
    "Nível piso -1": "Level -1",
    "Nível piso -2": "Level -2",
    "Nível piso -3": "Level -3",
    "Nível piso 1": "Level 1",
    "Nível piso 2": "Level 2",
    "Nível provável (inferido)": "Probable level (inferred)",
    "O alvo ": "Target ",
    "O alvo mais afetado (": "The most affected target (",
    "O contorno castanho é a envolvente dos alvos da contenção (convex hull), que aproxima o limite do recinto de escavação. Os alvos dos edifícios vizinhos aparecem agrupados e identificados por cor. As setas mostram a direção e magnitude do deslocamento acumulado (amplificado). Todos os alvos são de periferia — na cortina de contenção ou nas fachadas vizinhas; não há instrumentos dentro da escavação. Tudo no sistema de coordenadas dos alvos, sem necessidade de DXF.": "The brown outline is the envelope of the wall targets (convex hull), which approximates the boundary of the excavation. The targets of the neighbouring buildings appear grouped and identified by colour. The arrows show the direction and magnitude of the cumulative displacement (amplified). All targets are on the perimeter — on the retaining wall or on the neighbouring façades; there are no instruments inside the excavation. Everything in the coordinate system of the targets, with no need for a DXF.",
    "O desenho tem **": "The drawing has **",
    "O eixo vertical é a COTA (m), partilhada com as cotas de projeto da escavação/contenção e com o nível freático de repouso. Assim ve-se a que profundidade anda a água face a cada piso e ao fundo de escavação.": "The vertical axis is the ELEVATION (m), shared with the design elevations of the excavation/retaining structure and with the rest water table. This shows at what depth the water is relative to each level and to the excavation bottom.",
    "O plano IMPACTADO (22/04/2026; data de estado 21/08/2026) é o plano CONTRATUAL com três suspensões de frentes inseridas durante a execução das estacas, recalculado com as durações contratuais. As datas são as do plano — não um registo independente da execução. A janela de instrumentação está assinalada.": "The IMPACTED schedule (22/04/2026; status date 21/08/2026) is the CONTRACTUAL schedule with three suspensions of fronts inserted during the execution of the piles, recalculated with the contractual durations. The dates are those of the schedule — not an independent record of the execution. The instrumentation window is marked.",
    "O propósito da ferramenta não se esgota na visualização das leituras: procura sobretudo **relacionar a deformação medida com os acontecimentos da obra** — o avanço da escavação, o rebaixamento do nível freático, a execução das ancoragens e a natureza do maciço. O separador **Correlações** cruza estas séries num eixo temporal comum; a interpretação das causas cabe ao engenheiro, e uma correlação temporal não demonstra, por si só, causalidade.": "The purpose of the tool is not limited to visualising the readings: above all it seeks to **relate the measured deformation to the events of the works** — the progress of the excavation, the drawdown of the water table, the execution of the anchors and the nature of the ground. The **Correlations** tab crosses these series on a common time axis; the interpretation of the causes belongs to the engineer, and a temporal correlation does not, on its own, demonstrate causation.",
    "O que ficou suspenso": "What was suspended",
    "O que se assume": "What is assumed",
    "O que é medido vs. o que é assumido: inferências, limitações e qualidade dos dados.": "What is measured vs. what is assumed: inferences, limitations and data quality.",
    "Obra": "Works",
    "Observado (máx)": "Observed (max)",
    "Observado máx (mm)": "Observed max (mm)",
    "Observações": "Remarks",
    "Ocorrência": "Event",
    "Ocorrência ou medida": "Event or measure",
    "Origem": "Origin",
    "Origem do desfasamento": "Origin of the delay",
    "Os alvos tapados tem leitura 0,0000 no relatório; a app trata-as como ausência de leitura (a série não cai a zero).": "The covered targets have a reading of 0.0000 in the report; the app treats them as missing readings (the series does not drop to zero).",
    "Os alçados mais deslocados situam-se em torno de 0,2 % de H (Clough & O'Rourke), acima das estimativas de projeto; os troços curtos travados por escoras de canto (FG, BF) e o anel (LM) ficam abaixo de 0,1 %. Valores subestimados: excluem o movimento anterior à zeragem de cada instrumento.": "The most displaced wall sections are around 0.2% of H (Clough & O'Rourke), above the design estimates; the short stretches braced by corner struts (FG, BF) and the ring (LM) stay below 0.1%. Underestimated values: they exclude the movement before the zeroing of each instrument.",
    "Os deslocamentos são milimétricos e as coordenadas em metros; amplifica-se para ver.": "The displacements are millimetric and the coordinates in metres; they are amplified to be visible.",
    "Os dois pontos do desenho coincidem.": "The two drawing points coincide.",
    "Os relatórios calculam a variação sobre a «blocagem». Na A26, esse valor (833,5 kN) é a leitura da célula no patamar de 1,0P do ensaio, e não a carga de cravação: a primeira leitura da monitorização é 957,4 kN.": "The reports compute the variation over the «lock-off load». At A26, that value (833.5 kN) is the cell reading at the 1.0P stage of the test, and not the lock-off load: the first monitoring reading is 957.4 kN.",
    "Os relatórios calculam a variação sobre a «blocagem»; na A26 esse valor (833,5 kN) é a leitura no patamar de ensaio. A app permite usar a primeira leitura (957,4 kN).": "The reports compute the variation over the «lock-off load»; at A26 that value (833.5 kN) is the reading at the test stage. The app allows the first reading (957.4 kN) to be used.",
    "Outros": "Others",
    "Para cada ancoragem: furação/selagem (○), injeção (barra roxa, do início ao fim) e tensionamento/ensaio (◆). A cor do tensionamento indica o intervalo entre o FIM da injeção e o tensionamento — o procedimento prevê 7 a 8 dias. Datas adotadas (folha Qualidade_Dados corrige registos incoerentes).": "For each anchor: drilling/grouting (○), grouting (purple bar, from start to end) and stressing/test (◆). The colour of the stressing indicates the interval between the END of grouting and stressing — the procedure provides for 7 to 8 days. Adopted dates (the Qualidade_Dados sheet corrects inconsistent records).",
    "Para um alvo, percorre-se o tempo usando apenas as leituras disponíveis até cada data. Marca-se o primeiro sinal, a entrada em alerta/alarme (critérios dos relatórios, ΔH ou ΔV) e a antecedência entre ambos.": "For one target, time is traversed using only the readings available up to each date. The first signal, the entry into alert/alarm (criteria of the reports, ΔH or ΔV) and the lead time between both are marked.",
    "Parcial (+32 d.u.): o plano permite iniciar as estacas seguintes 18 d.u. antes do fim da suspensão": "Partial (+32 w.d.): the schedule allows the following piles to start 18 w.d. before the end of the suspension",
    "Paredes de escavação": "Excavation walls",
    "Parâmetro": "Parameter",
    "Parâmetros propostos (Quadros V-VII do relatório) que alimentam a modelação numérica da contenção. A coluna 'SPT típico' é a ponte para o perfil acima: é por ela que se lê em que zona está cada troço de terreno.": "Proposed parameters (Tables V-VII of the report) feeding the numerical modelling of the retaining structure. The 'Typical SPT' column is the bridge to the profile above: it shows in which zone each stretch of ground lies.",
    "Perfil deformado": "Deformed profile",
    "Perfil deformado em profundidade, evolução no tempo, velocidade e sinais precursores. Sobreposição da geologia.": "Deformed profile with depth, evolution over time, velocity and precursor signals. Geology overlay.",
    "Período": "Period",
    "Piezometros": "Piezometers",
    "Piezómetro": "Piezometer",
    "Piezómetros": "Piezometers",
    "Piezómetros — cota da água vs. escavação": "Piezometers — water level vs. excavation",
    "Piso": "Level",
    "Piso -1": "Level -1",
    "Piso -2": "Level -2",
    "Piso -3": "Level -3",
    "Piso -4": "Level -4",
    "Piso 1": "Level 1",
    "Piso 2 (coroamento zona alta)": "Level 2 (capping beam, high zone)",
    "Plano de trabalhos contratual e impactado, as suspensões de frentes (impactos) e o desfasamento de cada frente.": "Contractual and impacted schedules, the suspensions of fronts (impacts) and the delay of each front.",
    "Plano impactado": "Impacted schedule",
    "Plano: execução": "Schedule: execution",
    "Plano: tensionamento": "Schedule: stressing",
    "Planta (DXF)": "Plan (DXF)",
    "Planta alinhada com os alvos. Verifica visualmente se o contorno bate certo com a nuvem de alvos; se não, ajusta os pontos de referência.": "Plan aligned with the targets. Check visually whether the outline matches the cloud of targets; if not, adjust the reference points.",
    "Planta de localização das células — A26 (Piso -1, cota 15,90)": "Location plan of the cells — A26 (Level -1, elev. 15.90)",
    "Planta de localização geral (recinto) — Edifício 1 assinalado": "General location plan (site) — Building 1 marked",
    "Planta do projeto (DXF)": "Design plan (DXF)",
    "Planta em coordenadas locais do desenho. Ativa o alinhamento acima para a sobrepor aos alvos.": "Plan in local drawing coordinates. Enable the alignment above to overlay it on the targets.",
    "Poente": "West",
    "Poente/Sul": "West/South",
    "Poente/Sul — cortina do Edifício 1": "West/South — Building 1 wall",
    "Pontos insuficientes na janela da escavação para o XY.": "Not enough points in the excavation window for the XY plot.",
    "Posição ao longo da fachada (m, relativo)": "Position along the façade (m, relative)",
    "Poucos pontos na janela comum ao piezómetro para calcular a correlação.": "Too few points in the window common to the piezometer to compute the correlation.",
    "Precisas de indicar, para 2 alvos, onde eles estão no desenho. A coordenada no sistema dos alvos já vêm do Excel.": "You need to indicate, for 2 targets, where they are in the drawing. The coordinate in the target system already comes from the Excel file.",
    "Precursores": "Precursors",
    "Preenche as coordenadas dos 2 pontos no desenho para calcular o alinhamento.": "Fill in the coordinates of the 2 points in the drawing to compute the alignment.",
    "Pressuposto": "Assumption",
    "Pressuposto declarado": "Stated assumption",
    "Pressupostos": "Assumptions",
    "Pressupostos, inferências e qualidade dos dados": "Assumptions, inferences and data quality",
    "Previsto (mm)": "Predicted (mm)",
    "Previsto (projeto)": "Predicted (design)",
    "Primeira": "First",
    "Primeira leitura": "First reading",
    "Primeira leitura (957,4 kN)": "First reading (957.4 kN)",
    "Primeira leitura da monitorização": "First monitoring reading",
    "Profundidade": "Depth",
    "Profundidade (m)": "Depth (m)",
    "Profundidade abaixo da boca (m)": "Depth below collar (m)",
    "Profundidade do deslocamento máximo de cada inclinómetro e evolução da deformada ao longo do tempo.": "Depth of the maximum displacement of each inclinometer and evolution of the deformed shape over time.",
    "Profundidade do máx (m)": "Depth of max (m)",
    "Profundidade do máximo (m)": "Depth of the maximum (m)",
    "Profundidade fixa (": "Fixed depth (",
    "Profundidade fixa (m)": "Fixed depth (m)",
    "Propagação no plano": "Propagation in the schedule",
    "Página": "Page",
    "Quantificação factual do que a instrumentação mediu: quanto, onde, a que ritmo e a que profundidade, e como se compara com a estimativa de projeto. Ilustração de resultados, sem interpretação de causas.": "Factual quantification of what the instrumentation measured: how much, where, at what rate and at what depth, and how it compares with the design estimate. Illustration of results, without interpretation of causes.",
    "Quinzenal": "Fortnightly",
    "Re-zeragem A5b–A8b": "Re-zeroing A5b–A8b",
    "Realçar Edifício 1": "Highlight Building 1",
    "Realçar fachadas do Edifício 1": "Highlight the Building 1 façades",
    "Recalculado de ΔH/ΔV com os critérios do relatório; comparado com a coluna Estado do Excel (fórmula) — coerência entre implementações, não confirmação do estado publicado nos relatórios.": "Recalculated from ΔH/ΔV with the criteria of the report; compared with the State column of the Excel file (formula) — consistency between implementations, not confirmation of the state published in the reports.",
    "Referencial de cota (3D)": "Elevation reference (3D)",
    "Referência da variação de carga": "Reference of the load variation",
    "Referência das células de carga": "Load cell reference",
    "Registo: furação": "Record: drilling",
    "Registo: tensionamento": "Record: stressing",
    "Registos do empreiteiro (furação e selagem, multi-injeção e ensaios). Alçado GH segundo o alçado com a numeração das ancoragens; no troço poente do AB, o nível é inferido e a correspondência com o plano (?) está por confirmar. O plano trata as ancoragens dos dois lados de cada nível numa só tarefa. O procedimento de execução prevê o tensionamento 7 a 8 dias após o fim das injeções. Datas incoerentes nos registos foram estimadas (folha Qualidade_Dados).": "Contractor records (drilling and grouting, multi-grouting and tests). Section GH according to the elevation drawing with the anchor numbering; on the western stretch of AB, the level is inferred and the correspondence with the schedule (?) is to be confirmed. The schedule treats the anchors on both sides of each level as a single task. The execution procedure provides for stressing 7 to 8 days after the end of grouting. Inconsistent dates in the records were estimated (Qualidade_Dados sheet).",
    "Regra": "Rule",
    "Regra melhorada (Secção 5.3.5): compara cada leitura com a última feita há pelo menos 7 dias, em vez da leitura imediatamente anterior. Aplica-se à réplica em tempo real dos precursores.": "Improved rule (Section 5.3.5): compares each reading with the last one taken at least 7 days before, instead of the immediately preceding reading. Applies to the real-time replay of the precursors.",
    "Regular": "Regular",
    "Relatórios": "Reports",
    "Relatórios (833,5 kN)": "Reports (833.5 kN)",
    "Ressalvas identificadas no próprio modelo de dados, transcritas da folha Qualidade_Dados.": "Caveats identified in the data model itself, transcribed from the Qualidade_Dados sheet.",
    "Edifício 2": "Building 2",
    "Resultados": "Results",
    "Resultados — movimentações observadas": "Results — observed movements",
    "Rotação": "Rotation",
    "SE/nascente": "SE/east",
    "SPT típico (N)": "Typical SPT (N)",
    "Edifício 1 (pisos 2 e 1)": "Building 1 (levels 2 and 1)",
    "Edifício 1": "Building 1",
    "Sem ancoragens para o alçado escolhido.": "No anchors for the chosen section.",
    "Sem componentes ΔM/ΔP para desenhar a roseta destes alvos.": "No ΔM/ΔP components to draw the rose of these targets.",
    "Sem coordenadas suficientes para esquematizar este edifício.": "Not enough coordinates to sketch this building.",
    "Sem dados de alvos.": "No target data.",
    "Sem dados de células de carga.": "No load cell data.",
    "Sem dados de deformação para a série escolhida.": "No deformation data for the chosen series.",
    "Sem dados de inclinómetros para o perfil de profundidade.": "No inclinometer data for the depth profile.",
    "Sem dados suficientes para os envelopes empíricos.": "Not enough data for the empirical envelopes.",
    "Sem desfasamento: primeiras estacas poente e norte, viga de coroamento E8_77–E8_104 e ancoragens, escoras F-G/G-H e H-I/I-J, VD do piso 1 em EF/FG/GH e ancoragens": "No delay: first west and north piles, capping beam E8_77–E8_104 and anchors, struts F-G/G-H and H-I/I-J, level 1 distribution beam at EF/FG/GH and anchors",
    "Sem instrumentos em atraso face à data de estado.": "No instruments delayed relative to the status date.",
    "Sem leitura": "No reading",
    "Sem leituras com data válida.": "No readings with a valid date.",
    "Semanal": "Weekly",
    "Sim": "Yes",
    "Sinais": "Signals",
    "Sinais precursores — antecedência e seletividade": "Precursor signals — lead time and selectivity",
    "Sinais seguidos de alerta (%)": "Signals followed by alert (%)",
    "Sinal": "Signal",
    "Sinal precursor": "Precursor signal",
    "Sistema": "System",
    "Sobrepor geologia + SPT da sondagem": "Overlay geology + SPT of the borehole",
    "Sondagem": "Borehole",
    "Sondagem de referência": "Reference borehole",
    "Sugerida por proximidade: **": "Suggested by proximity: **",
    "Sul": "South",
    "Superfície do terreno entre as cotas ": "Terrain surface between elevations ",
    "Superfície real do terreno (modelo digital triangulado) do levantamento topográfico, no referencial da obra e com COTAS REAIS. Mostra a topografia original e o volume a escavar até a cota de fundo (4,55 m). Complementa o 3D dos alvos — que usa um referencial local e cotas relativas — com geometria absoluta.": "Real terrain surface (triangulated digital model) from the topographic survey, in the site reference system and with REAL ELEVATIONS. Shows the original topography and the volume to be excavated down to the bottom elevation (4.55 m). Complements the 3D of the targets — which uses a local reference system and relative elevations — with absolute geometry.",
    "Supressão de um dos perfis de ligação do anel de compressão às fundações": "Removal of one of the sections connecting the compression ring to the foundations",
    "Suspensão": "Suspension",
    "São precisas pelo menos duas séries para construir a matriz.": "At least two series are needed to build the matrix.",
    "Série": "Series",
    "Série curta de mais para a réplica.": "Series too short for the replay.",
    "Série de base": "Base series",
    "Série de deformação": "Deformation series",
    "Série insuficiente (alvo ou piezómetro) para a correlação.": "Insufficient series (target or piezometer) for the correlation.",
    "Séries temporais de deslocamento horizontal e assentamento vertical, por edifício e por alvo.": "Time series of horizontal displacement and vertical settlement, by building and by target.",
    "Temperatura (°C)": "Temperature (°C)",
    "Tendência": "Trend",
    "Tensionamento (": "Stressing (",
    "Tensionamento (sem registo do intervalo)": "Stressing (no record of the interval)",
    "Tensionamento e ensaio (adotado)": "Stressing and test (adopted)",
    "Terreno": "Terrain",
    "Terreno + Alvos 3D": "Terrain + 3D targets",
    "Terreno + Alvos 3D — movimento sobre o relevo real": "Terrain + 3D targets — movement over the real relief",
    "Terreno 3D — modelo do levantamento topográfico": "3D terrain — topographic survey model",
    "Terreno original": "Original terrain",
    "Terreno original (sem escavar)": "Original terrain (not excavated)",
    "The Eden - Instrumentação": "The Eden - Instrumentation",
    "Tirantes nos muros sobre as vigas de coroamento": "Tie-backs in the walls above the capping beams",
    "Todos": "All",
    "Total": "Total",
    "Transparência sobre o que é MEDIDO e o que é ASSUMIDO. Cada pressuposto tem a sua origem e o estado de confirmação. Esta secção reúne, num só sítio, as ressalvas assinaladas ao longo da aplicação.": "Transparency about what is MEASURED and what is ASSUMED. Each assumption has its origin and confirmation status. This section brings together, in one place, the caveats flagged throughout the application.",
    "Tratados como séries separadas; os 'b' foram re-zerados em 20/10/2025 e os acumulados não são comparáveis diretamente com A1-A4.": "Treated as separate series; the 'b' targets were re-zeroed on 20/10/2025 and their cumulative values are not directly comparable with A1-A4.",
    "Tratamento": "Processing",
    "Travessa metálica (escoramento) na banda do Piso -2": "Steel cross-member (bracing) on the Level -2 strip",
    "Um nível de carotes Ø80 mm ao nível do piso 1, entre estacas, para aliviar a pressão da água": "One row of Ø80 mm cored holes at level 1, between piles, to relieve the water pressure",
    "Um sinal precursor ocorre quando a velocidade do deslocamento horizontal entre leituras atinge o limiar (velocidade) ou acelera para o fator definido (aceleração). Os limiares da barra lateral (v ≥ ": "A precursor signal occurs when the velocity of the horizontal displacement between readings reaches the threshold (velocity) or accelerates by the defined factor (acceleration). The sidebar thresholds (v ≥ ",
    "Uma sondagem (detalhe)": "One borehole (detail)",
    "Unidade": "Unit",
    "Usar o ficheiro ao lado do script": "Use the file next to the script",
    "VD Piso -1": "Level -1 distribution beam",
    "VD Piso -2": "Level -2 distribution beam",
    "VD Piso -2 + fundo Banda Laje": "Level -2 distribution beam + bottom of slab strip",
    "VD Piso -3": "Level -3 distribution beam",
    "VD Piso -3 + VD da grua (2 ancoragens)": "Level -3 distribution beam + crane distribution beam (2 anchors)",
    "Valor": "Value",
    "Valores da memória descritiva (17 m: 20/40 · 10/15 mm; 24 m: 30/40 · 10/15 mm), atribuídos a cada alçado segundo os relatórios de alvos: 17 m em AB, CD, DE, BF e PQ; 24 m em FG, GH, JK, KL, LM, MNO e OP (DE e LM a partir do R36, de 16/03/2026). MNO e OP têm 17 a 20 m de cortina, mas são tratados como 24 m.": "Values of the design report (17 m: 20/40 · 10/15 mm; 24 m: 30/40 · 10/15 mm), assigned to each wall section according to the target reports: 17 m at AB, CD, DE, BF and PQ; 24 m at FG, GH, JK, KL, LM, MNO and OP (DE and LM from R36, of 16/03/2026). MNO and OP have 17 to 20 m of wall, but are treated as 24 m.",
    "Variação (%)": "Variation (%)",
    "Variação calculada": "Computed variation",
    "Variação calculada em relação à primeira leitura (": "Variation computed relative to the first reading (",
    "Velocidade": "Velocity",
    "Velocidade (mm/dia entre campanhas) do deslocamento horizontal de um alvo. Mostra quando o movimento acelerou.": "Velocity (mm/day between campaigns) of the horizontal displacement of a target. Shows when the movement accelerated.",
    "Velocidade (mm/dia)": "Velocity (mm/day)",
    "Velocidade e sinais precursores": "Velocity and precursor signals",
    "Velocidade entre leituras consecutivas": "Velocity between consecutive readings",
    "Velocidade máx. (mm/dia)": "Max. velocity (mm/day)",
    "Velocidade máxima do alvo ": "Maximum velocity of target ",
    "Velocidade sobre intervalo ≥ 7 dias": "Velocity over an interval ≥ 7 days",
    "Velocidade sobre ≥ 7 dias (reduz ruído)": "Velocity over ≥ 7 days (reduces noise)",
    "Ver divergências estado recalculado vs. relatório": "See discrepancies recalculated state vs. report",
    "Vetores do deslocamento horizontal (componentes ΔM/ΔP) na última campanha. 0° = Norte (+P), 90° = Este (+M). O desvio do rumo face à direção do alçado é leitura do autor (a tese obtém ≤ 26°).": "Horizontal displacement vectors (ΔM/ΔP components) in the last campaign. 0° = North (+P), 90° = East (+M). The deviation of the bearing from the direction of the wall section is the author’s reading (the dissertation obtains ≤ 26°).",
    "Viga coroamento + muros": "Capping beam + walls",
    "Viga coroamento + muros contenção": "Capping beam + retaining walls",
    "Viga da grua": "Crane beam",
    "Viga de coroamento/distribuição (projeto: cota real)": "Capping/distribution beam (design: real elevation)",
    "Viga de coroamento/distribuição<br>cota ": "Capping/distribution beam<br>elev. ",
    "Vista": "View",
    "Vista 3D (4/4)": "3D view (4/4)",
    "Vistas 3D (1/4) — vista global": "3D views (1/4) — global view",
    "Vistas 3D (2/4)": "3D views (2/4)",
    "Vistas 3D (3/4)": "3D views (3/4)",
    "Volume de escavação": "Excavation volume",
    "X no desenho": "X in the drawing",
    "X, Y local do desenho": "Local X, Y of the drawing",
    "Y no desenho": "Y in the drawing",
    "Yoo (2001) — terreno sobre rocha (0,05–0,15 %)": "Yoo (2001) — soil overlying rock (0.05–0.15%)",
    "Zeragem a 19/01/2026. A grandeza medida e a convenção de sinal não constam do quadro de leituras (a confirmar); variações de 0,5 mm estão no limite da resolução. Linhas roxas: medidas com data indicadas pela obra (separador «Obra»).": "Zeroed on 19/01/2026. The measured quantity and the sign convention are not given in the readings table (to be confirmed); variations of 0.5 mm are at the limit of the resolution. Purple lines: dated measures reported by the site («Works» tab).",
    "Zona": "Zone",
    "Zona geotécnica no perfil SPT": "Geotechnical zone in the SPT profile",
    "aceleração": "acceleration",
    "alarme": "alarm",
    "alerta": "alert",
    "alvo": "target",
    "ancoragem": "anchor",
    "ancoragens": "anchors",
    "até cota 19,80/19,00 (fundo VD Piso 1)": "to elev. 19.80/19.00 (bottom of level 1 beam)",
    "até cota 21,45": "to elev. 21.45",
    "até cota VD Piso -1": "to elev. of level -1 beam",
    "até cota VD Piso -2": "to elev. of level -2 beam",
    "até cota VD Piso -3": "to elev. of level -3 beam",
    "blocagem": "lock-off load",
    "campanha": "campaign",
    "campanhas": "campaigns",
    "canto SO (poente)": "SW corner (west)",
    "clínica": "clinic",
    "com a escavação de 10–25/02/2026": "with the excavation of 10–25/02/2026",
    "com a escavação de 12–25/11/2025": "with the excavation of 12–25/11/2025",
    "contenção": "wall",
    "contratual": "contractual",
    "cota": "elev.",
    "cota 14,85 (-4,95 m) + VD da grua (2 ancoragens)": "elev. 14.85 (-4.95 m) + crane beam (2 anchors)",
    "cota 22,55 -> 18,90": "elev. 22.55 -> 18.90",
    "cota 22,55 → 18,90": "elev. 22.55 → 18.90",
    "cota 26,50 -> 23,75": "elev. 26.50 -> 23.75",
    "cota 26,50 → 23,75": "elev. 26.50 → 23.75",
    "cota 29 -> 22,55": "elev. 29 -> 22.55",
    "cota 29 → 22,55 (fundo Anel P2)": "elev. 29 → 22.55 (bottom of ring P2)",
    "data": "date",
    "de 4,95 m → cota 14,85": "from 4.95 m → elev. 14.85",
    "deslocamento": "displacement",
    "dias": "days",
    "escavação": "excavation",
    "frontal": "front",
    "fundo": "bottom",
    "fundo VD Piso 1 (19,80/19,00)": "bottom of level 1 beam (19.80/19.00)",
    "fundo de escavação": "excavation bottom",
    "gama (kN/m3)": "gamma (kN/m3)",
    "gama: peso volúmico | c': coesão | fi': ângulo de atrito | E': módulo de deformabilidade. Zonas ZG3-ZG1 (rocha) com c' e E' em MPa/GPa; ZG6-ZG4 (solo/grés brando) em kPa/MPa. Nota: as três zonas de nega (ZG3-ZG1) distinguem-se pelo RQD, que o SPT sozinho não mede — por isso o perfil agrupa-as como 'nega'.": "gamma: unit weight | c': cohesion | phi': friction angle | E': deformation modulus. Zones ZG3-ZG1 (rock) with c' and E' in MPa/GPa; ZG6-ZG4 (soil/soft sandstone) in kPa/MPa. Note: the three refusal zones (ZG3-ZG1) are distinguished by the RQD, which the SPT alone does not measure — so the profile groups them as 'refusal'.",
    "impactado": "impacted",
    "indireta (edifício vizinho)": "indirect (neighbouring building)",
    "lateral": "side",
    "leitura": "reading",
    "leituras": "readings",
    "máx.": "max.",
    "média": "mean",
    "mín.": "min.",
    "não discriminado": "not specified",
    "piso": "level",
    "r > 0: as séries sobem juntas; r < 0: uma sobe enquanto a outra desce (típico da água, cuja cota desce quando o rebaixamento acompanha a escavação). A matriz de incrementos é a mais exigente: só fica alta se os movimentos coincidirem no tempo. Correlação não é causa — a leitura das causas é do autor no texto.": "r > 0: the series rise together; r < 0: one rises while the other falls (typical of the water, whose level falls when the drawdown follows the excavation). The increments matrix is the most demanding: it is only high if the movements coincide in time. Correlation is not causation — the reading of the causes is the author’s, in the text.",
    "relatório": "report",
    "relatórios": "reports",
    "sem": "none",
    "semanas": "weeks",
    "sinal": "signal",
    "sinal após alerta": "signal after alert",
    "sistema dos alvos (M, P)": "target system (M, P)",
    "topo N (bolbo curvo)": "N top (curved bulb)",
    "velocidade": "velocity",
    "» (máx ": "» (max ",
    "» — observado ": "» — observed ",
    "Água": "Water",
    "Água ": "Water ",
    "Água abundante numa ancoragem junto à célula da A26; recomendadas furação de drenagem entre estacas e tirante preventivo": "Abundant water at an anchor next to the A26 cell; cored drainage holes between piles and a preventive tie-back recommended",
    "×) aplicam-se à réplica em tempo real; a tabela de sensibilidade usa os valores da dissertação.": "×) apply to the real-time replay; the sensitivity table uses the values of the dissertation.",
    "Últ. desloc. máx. (mm)": "Last max. disp. (mm)",
    "Última leitura": "Last reading",
    "água": "water",
    "ΔM acumulado (mm)": "Cumulative ΔM (mm)",
    "ΔP acumulado (mm)": "Cumulative ΔP (mm)",
    "ΔZ acumulado (mm)": "Cumulative ΔZ (mm)",
    "δ_h,máx (mm)": "δ_h,max (mm)",
    "δh acumulado": "cumulative δh",
    "δh acumulado (mm)": "cumulative δh (mm)",
    "↔ Deslocamento horizontal (H)": "↔ Horizontal displacement (H)",
    "↕ Assentamento vertical (ΔZ)": "↕ Vertical settlement (ΔZ)",
    "↳ Frente de escavação associada a **": "↳ Excavation front associated with **",
    "↳ Frente de escavação: **": "↳ Excavation front: **",
    "↳ Sem zona de escavação associada a esta série — mostram-se os marcos globais de escavação.": "↳ No excavation zone associated with this series — the global excavation milestones are shown.",
    "≤ 8 dias": "≤ 8 days",
    "⚓ = ancoragens · ⟷ = escoras metálicas · ▭ = banda de laje (betonagem), nas datas do plano impactado para a frente. A cota desce em degraus (cada degrau = uma cota atingida).": "⚓ = anchors · ⟷ = steel struts · ▭ = slab strip (casting), at the dates of the impacted schedule for the front. The level goes down in steps (each step = one level reached).",
    "⚠ A localização da célula ": "⚠ The location of cell ",
    "⚠ As células NÃO começaram a medir na mesma data: a ": "⚠ The cells did NOT start measuring on the same date: ",
    "⚠ Os elementos de contenção (viga de coroamento/distribuição, bandas de laje, escoras metálicas) são ESQUEMÁTICOS mas INFORMADOS PELO PROJETO: as cotas dos pisos e as cores das lajes seguem as peças desenhadas da JETsj (EDN-JET-ZZ-ZZ-DR-U-0021..0027), e as escoras são horizontais entre lados (como nos cortes tipo). A FORMA em planta (contorno curvo, geometria exata de cada banda) e aproximada — não foi extraída do projeto. Para a representação rigorosa, ver as Vistas 3D do projeto no fim deste separador.": "⚠ The retaining elements (capping/distribution beam, slab strips, steel struts) are SCHEMATIC but INFORMED BY THE DESIGN: the floor elevations and the colours of the slabs follow the JETsj drawings (EDN-JET-ZZ-ZZ-DR-U-0021..0027), and the struts are horizontal between sides (as in the typical sections). The SHAPE in plan (curved outline, exact geometry of each strip) is approximate — it was not extracted from the design. For the rigorous representation, see the 3D views of the design at the end of this tab.",
    "⛏ Escavação na frente (cota atingida, data do planeamento): ": "⛏ Excavation at the front (level reached, planning date): ",
    "✓ Localização confirmada: a célula ": "✓ Location confirmed: cell ",
    "📈 Evolução do deslocamento": "📈 Displacement evolution",
    "📍 Ver localização da célula ": "📍 See location of cell ",
    "📍 Ver localização do alvo ": "📍 See location of target ",
    "📐 Perfil deformado": "📐 Deformed profile",
    "📐 Vistas 3D do projeto de contenção (JETsj) — referência": "📐 3D views of the retaining structure design (JETsj) — reference",
    # --- Alvos (2D): sobreposicoes (V79) ---
    "Faseamento no gráfico": "Phasing on the chart",
    "Nenhum": "None",
    "Períodos de escavação": "Excavation periods",
    "Fases da obra (plano)": "Construction phases (schedule)",
    "Assinalar o máximo": "Mark the maximum",
    "Datas de referência": "Reference dates",
    "Data de referência": "Reference date",
    "Comparar com (intervalo mín.–máx.)": "Compare with (min.–max. range)",
    "Períodos de escavação: seis períodos delimitados pelo início da escavação de cada nível e pelo fim da escavação, segundo o plano de trabalhos impactado (faixas alternadas). Fases da obra: barra de faseamento do plano no topo do gráfico. São uma referência comum: as frentes avançaram desfasadas.": "Excavation periods: six periods bounded by the start of the excavation of each level and by the end of the excavation, according to the impacted works schedule (alternating bands). Construction phases: phasing bar of the schedule at the top of the chart. They are a common reference: the fronts advanced out of step.",
    "Assinala o maior deslocamento horizontal e o maior |ΔZ| dos alvos escolhidos, no período visível.": "Marks the largest horizontal displacement and the largest |ΔZ| of the selected targets, in the visible period.",
    "Marca as campanhas escolhidas (linha vertical e valor de cada alvo) e lista, por baixo de cada gráfico, os valores nessas datas e os incrementos entre elas.": "Marks the selected campaigns (vertical line and value of each target) and lists, below each chart, the values on those dates and the increments between them.",
    "Sobrepõe, em faixa, o intervalo entre o mínimo e o máximo dos alvos de outro grupo em cada campanha. Só entram os alvos com a zeragem inicial do grupo; os de substituição, zerados mais tarde, ficam de fora.": "Overlays, as a band, the range between the minimum and the maximum of the targets of another group in each campaign. Only the targets with the initial zeroing of the group are included; replacement targets, zeroed later, are left out.",
    "Os pares de alvos à mesma posição em planta (a menos de 1 m) e a cotas diferentes partilham a cor: tom escuro no superior (sup.) e claro no inferior (inf.).": "Pairs of targets at the same plan position (less than 1 m apart) and at different elevations share the colour: dark shade for the upper one (upper) and light shade for the lower one (lower).",
    "Faixas: intervalo entre o mínimo e o máximo, em cada campanha, dos alvos de cada grupo comparado (só os alvos com a zeragem inicial do grupo).": "Bands: range between the minimum and the maximum, in each campaign, of the targets of each compared group (only the targets with the initial zeroing of the group).",
    "Valores nas datas de referência (mm; última leitura até cada data, no máximo 10 dias antes) e incrementos entre datas consecutivas:": "Values on the reference dates (mm; last reading up to each date, at most 10 days before) and increments between consecutive dates:",
    "Períodos de escavação, segundo o plano de trabalhos impactado (referência comum; as frentes avançaram desfasadas): ": "Excavation periods, according to the impacted works schedule (common reference; the fronts advanced out of step): ",
    "Antes da<br>escavação": "Before<br>excavation",
    "Pisos 2 e 1": "Levels 2 and 1",
    "Piso -3<br>e fundo": "Level -3<br>and bottom",
    "Laje de fundo<br>e estrutura": "Bottom slab<br>and structure",
    "antes da escavação geral": "before the general excavation",
    "pisos 2 e 1": "levels 2 and 1",
    "piso -1, com a escavação de 4,95 m da frente poente": "level -1, with the 4.95 m excavation of the west front",
    "piso -2 e escavação até à viga do piso -3": "level -2 and excavation down to the level -3 beam",
    "piso -3 e fundo": "level -3 and bottom",
    "laje de fundo e estrutura": "bottom slab and structure",
    "última leitura": "last reading",
    "Máx.": "Max.",
    "Outros grupos": "Other groups",
    "Escolher campanhas": "Choose campaigns",
    "Escolher grupos": "Choose groups",
    "(mín.–máx.)": "(min.–max.)",
    "mín.": "min.",
    "(sup.)": "(upper)",
    "(inf.)": "(lower)",
    "Re-zeragem": "Re-zeroing",
    # --- V81: estimativa do projeto ---
    "Estimativa do projeto": "Design estimate",
    "Estimativa do projeto: ": "Design estimate: ",
    "ultrapassada a ": "exceeded on ",
    "Estimativa do projeto ultrapassada a ": "Design estimate exceeded on ",
    "Estimativa do projeto (memória descritiva)": "Design estimate (design report)",
    "Desenha no gráfico horizontal a estimativa da memória descritiva para a última fase de escavação na zona deste alçado (20 mm a poente e norte; 10 mm a nascente e sul; atribuição inferida) e a data em que o máximo dos alvos escolhidos a ultrapassou.": "Draws on the horizontal chart the design-report estimate for the last excavation stage in the zone of this wall section (20 mm west and north; 10 mm east and south; inferred assignment) and the date on which the maximum of the selected targets exceeded it.",
    "Deslocamento horizontal máximo OBSERVADO (todas as campanhas) vs. ESTIMATIVA DO PROJETO para a última fase de escavação (memória descritiva: ~20 mm na cortina poente/norte e ~10 mm na nascente/sul). Nos alçados, a estimativa é a da zona do projeto em que cada um se situa (atribuição inferida da posição; menos segura nos alçados BF, CD e DE, no canto sudoeste). Nos edifícios vizinhos, a comparação é INDIRETA: a estimativa refere-se à cortina adjacente, não ao edifício. A coluna «Ultrapassada a» indica a primeira campanha em que o máximo do grupo atingiu a estimativa.": "Maximum OBSERVED horizontal displacement (all campaigns) vs. DESIGN ESTIMATE for the last excavation stage (design report: ~20 mm on the west/north wall and ~10 mm on the east/south wall). For wall sections, the estimate is that of the design zone where each one lies (assignment inferred from the position; less certain for sections BF, CD and DE, in the south-west corner). For neighbouring buildings, the comparison is INDIRECT: the estimate refers to the adjacent wall, not to the building. The «Exceeded on» column gives the first campaign in which the maximum of the group reached the estimate.",
    "Zona": "Zone",
    "Observado/previsto": "Observed/predicted",
    "Ultrapassada a": "Exceeded on",
    "Maior excedência na contenção: ": "Largest exceedance on the retaining wall: ",
    " vezes a estimativa de ": " times the estimate of ",
    "poente/norte": "west/north",
    "Deslocamento horizontal máximo de cada alçado (máximo dos prismas em todas as campanhas) e de cada inclinómetro (máximo do perfil), em função da profundidade escavada H (do topo das estacas ao fundo de escavação), face aos referenciais empíricos e às estimativas do projeto (memória descritiva, por zona do projeto). As linhas a tracejado são δ_h = p%·H; a sombra marca as faixas de Long e de Yoo.": "Maximum horizontal displacement of each wall section (maximum of the prisms over all campaigns) and of each inclinometer (maximum of the profile), as a function of the excavated depth H (from the top of the piles to the excavation bottom), against the empirical references and the design estimates (design report, by design zone). Dashed lines are δ_h = p%·H; shading marks the Long and Yoo bands.",
    "Os alçados mais deslocados situam-se em torno de 0,2 % de H (Clough & O'Rourke), acima das estimativas do projeto; os troços curtos travados por escoras de canto (FG, BF) e os alçados DE e LM, com prismas zerados só em 22/12/2025 e 05/02/2026, ficam abaixo de 0,1 %. Valores subestimados: excluem o movimento anterior à zeragem de cada instrumento.": "The most displaced wall sections lie around 0.2 % of H (Clough & O'Rourke), above the design estimates; the short sections braced by corner struts (FG, BF) and sections DE and LM, with prisms zeroed only on 22/12/2025 and 05/02/2026, stay below 0.1 %. Underestimated values: they exclude the movement before each instrument was zeroed.",
    "Profundidade escavada, H (m)": "Excavated depth, H (m)",
    "nascente/sul": "east/south",
    "indireta": "indirect",
    "norte": "north",
}

_MESES_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep",
             "Oct", "Nov", "Dec"]


def _i18n_padrao(k):
    a = r"(?<!\w)" if k[:1].isalnum() else ""
    b = r"(?!\w)" if k[-1:].isalnum() else ""
    return a + _re_i18n.escape(k) + b


_I18N_RX = _re_i18n.compile("|".join(
    _i18n_padrao(k) for k in sorted(_I18N_EN, key=len, reverse=True)
    if len(k.strip()) >= 3))
_I18N_DEC = _re_i18n.compile(r"(?<=\d),(?=\d)")


def _lang_en():
    try:
        return st.session_state.get("ims_lang", "Português") == "English"
    except Exception:
        return False


def _meses():
    return _MESES_EN if _lang_en() else _MESES_PT


def _tr(s):
    """Traduz um texto mostrado ao utilizador (só em inglês)."""
    if not isinstance(s, str) or not s or not _lang_en():
        return s
    t = _I18N_EN.get(s)
    if t is None:
        nu = s.strip()
        t = s.replace(nu, _I18N_EN[nu]) if nu and nu in _I18N_EN else \
            _I18N_RX.sub(lambda m: _I18N_EN.get(m.group(0), m.group(0)), s)
    t = _re_i18n.sub(r"(\d\**) de (\d)", r"\1 of \2", t)
    # separador decimal: 15,90 -> 15.90 (o HTML/CSS não tem números com vírgula)
    if "<div" not in t:
        t = _I18N_DEC.sub(".", t)
    return t


def _tr_exato(v):
    if isinstance(v, str) and _lang_en():
        return _I18N_EN.get(v, _I18N_EN.get(v.strip(), v))
    return v


def _tr_fig(fig):
    """Traduz títulos, eixos, legendas, anotações e nomes das séries."""
    if not _lang_en():
        return fig
    try:
        lay = fig.layout
        if lay.title and lay.title.text:
            lay.title.text = _tr(lay.title.text)
        for nome in list(lay):
            if nome.startswith(("xaxis", "yaxis")):
                ax = lay[nome]
                if ax.title and ax.title.text:
                    ax.title.text = _tr(ax.title.text)
                if ax.ticktext is not None:
                    ax.ticktext = [_tr(x) for x in ax.ticktext]
            if nome.startswith("scene") and lay[nome] is not None:
                for e in ("xaxis", "yaxis", "zaxis"):
                    ax = lay[nome][e]
                    if ax.title and ax.title.text:
                        ax.title.text = _tr(ax.title.text)
        if lay.legend and lay.legend.title and lay.legend.title.text:
            lay.legend.title.text = _tr(lay.legend.title.text)
        for an in (lay.annotations or ()):
            if an.text:
                an.text = _tr(an.text)
        for tr in fig.data:
            if getattr(tr, "name", None):
                tr.name = _tr(tr.name)
            lgt = getattr(tr, "legendgrouptitle", None)
            if lgt is not None and lgt.text:
                lgt.text = _tr(lgt.text)
            ht = getattr(tr, "hovertemplate", None)
            if isinstance(ht, str):
                tr.hovertemplate = _tr(ht)
            tx = getattr(tr, "text", None)
            if isinstance(tx, str):
                tr.text = _tr(tx)
            elif tx is not None and not isinstance(tx, (int, float)):
                try:
                    if all(isinstance(x, str) for x in tx):
                        tr.text = [_tr(x) for x in tx]
                except TypeError:
                    pass
            for eixo in ("x", "y"):
                v = getattr(tr, eixo, None)
                if v is not None and tr.type in ("bar", "box", "heatmap"):
                    try:
                        if len(v) and all(isinstance(x, str) for x in v):
                            setattr(tr, eixo, [_tr(x) for x in v])
                    except TypeError:
                        pass
            cb = getattr(getattr(tr, "marker", None), "colorbar", None) or getattr(tr, "colorbar", None)
            if cb is not None and cb.title and cb.title.text:
                cb.title.text = _tr(cb.title.text)
    except Exception:
        pass
    return fig


def _col_texto(serie):
    return (pd.api.types.is_object_dtype(serie) or pd.api.types.is_string_dtype(serie))


def _tr_celula(v):
    """Células: frase com espaços traduz-se como texto; códigos só por igualdade."""
    if isinstance(v, str) and " " in v.strip():
        return _tr(v)
    return _tr_exato(v)


def _tr_df(df):
    if not _lang_en():
        return df
    try:
        if isinstance(df, pd.io.formats.style.Styler):
            dados_ = df.data
            cols = [c for c in dados_.columns if _col_texto(dados_[c])]
            df = df.format_index(lambda c: _tr(str(c)), axis=1)
            df = df.format_index(lambda c: _tr_celula(c) if isinstance(c, str) else c, axis=0)
            if cols:
                df = df.format(_tr_celula, subset=cols)
            return df
        if isinstance(df, pd.DataFrame):
            out = df.copy()
            for c in out.columns:
                if _col_texto(out[c]):
                    out[c] = out[c].map(_tr_celula)
            out.columns = [_tr(str(c)) if isinstance(c, str) else c for c in out.columns]
            if _col_texto(out.index.to_series()):
                out.index = out.index.map(_tr_celula)
            if isinstance(out.index.name, str):
                out.index.name = _tr(out.index.name)
            return out
    except Exception:
        return df
    return df


_I18N_KW = ("help", "placeholder", "caption", "label", "value_label")


def _i18n_texto(orig):
    def w(self, *a, **k):
        if _lang_en():
            a = tuple(_tr(x) if (i == 0 or isinstance(x, str)) and isinstance(x, str) else x
                      for i, x in enumerate(a))
            for kk in _I18N_KW:
                if isinstance(k.get(kk), str):
                    k[kk] = _tr(k[kk])
        return orig(self, *a, **k)
    return w


def _i18n_metric(orig):
    def w(self, label, value=None, delta=None, *a, **k):
        if _lang_en():
            label = _tr(label)
            value = _tr(value) if isinstance(value, str) else value
            delta = _tr(delta) if isinstance(delta, str) else delta
            if isinstance(k.get("help"), str):
                k["help"] = _tr(k["help"])
        return orig(self, label, value, delta, *a, **k)
    return w


def _i18n_opcoes(orig):
    def w(self, label, *a, **k):
        if _lang_en():
            label = _tr(label)
            if isinstance(k.get("help"), str):
                k["help"] = _tr(k["help"])
            ff = k.get("format_func") or str
            k["format_func"] = lambda x, _ff=ff: _tr(_ff(x))
        return orig(self, label, *a, **k)
    return w


def _i18n_tabs(orig):
    def w(self, tabs, *a, **k):
        if _lang_en():
            tabs = [_tr(t) for t in tabs]
        return orig(self, tabs, *a, **k)
    return w


def _i18n_tabela(orig):
    def w(self, data=None, *a, **k):
        return orig(self, _tr_df(data), *a, **k)
    return w


def _i18n_plotly(orig):
    def w(self, fig, *a, **k):
        return orig(self, _tr_fig(_datas_pt(fig)), *a, **k)
    return w


# Os métodos originais guardam-se UMA vez: o Streamlit volta a executar o
# script em cada interação e, sem esta guarda, os invólucros acumulavam-se.
if not hasattr(_DG_i18n, "_i18n_orig"):
    _DG_i18n._i18n_orig = {_n: getattr(_DG_i18n, _n) for _n in (
        "markdown", "caption", "subheader", "header", "title", "info",
        "warning", "success", "error", "text", "expander", "button",
        "checkbox", "toggle", "slider", "number_input", "text_input",
        "file_uploader", "download_button", "popover", "radio", "selectbox",
        "multiselect", "select_slider", "metric", "tabs", "dataframe",
        "table", "plotly_chart", "write", "image") if hasattr(_DG_i18n, _n)}
_O = _DG_i18n._i18n_orig
for _n in ("markdown", "caption", "subheader", "header", "title", "info",
           "warning", "success", "error", "text", "expander", "button",
           "checkbox", "toggle", "slider", "number_input", "text_input",
           "file_uploader", "download_button", "popover"):
    if _n in _O:
        setattr(_DG_i18n, _n, _i18n_texto(_O[_n]))
for _n in ("radio", "selectbox", "multiselect", "select_slider"):
    setattr(_DG_i18n, _n, _i18n_opcoes(_O[_n]))
_DG_i18n.metric = _i18n_metric(_O["metric"])
_DG_i18n.tabs = _i18n_tabs(_O["tabs"])
_DG_i18n.dataframe = _i18n_tabela(_O["dataframe"])
_DG_i18n.table = _i18n_tabela(_O["table"])
_DG_i18n.plotly_chart = _i18n_plotly(_O["plotly_chart"])
_DG_i18n.write = _i18n_texto(_O["write"])


def _i18n_imagem(orig):
    def w(self, image, *a, **k):
        if _lang_en() and isinstance(k.get("caption"), str):
            k["caption"] = _tr(k["caption"])
        return orig(self, image, *a, **k)
    return w


_DG_i18n.image = _i18n_imagem(_O["image"])

_st_main_dg = getattr(_plotly_chart_original, "__self__", None) or st._main
for _n in ("markdown", "caption", "subheader", "header", "title", "info",
           "warning", "success", "error", "text", "expander", "button",
           "checkbox", "toggle", "slider", "number_input", "text_input",
           "file_uploader", "download_button", "image", "popover", "radio",
           "selectbox", "multiselect", "select_slider", "metric", "tabs",
           "dataframe", "table", "plotly_chart", "write"):
    if hasattr(_st_main_dg, _n):
        setattr(st, _n, getattr(_st_main_dg, _n))


st.set_page_config(page_title="IMS — Instrumentation Monitoring System",
                   layout="wide")


# =========================================================================
# CARREGAMENTO
# =========================================================================
@st.cache_data(show_spinner="A carregar o Excel...")
def _carregar_dados_cached(fonte, _versao):
    # _versao (mtime+tamanho do ficheiro) entra na chave de cache: quando o
    # Excel e substituido no repositorio, _versao muda e o Streamlit REle os
    # dados em vez de devolver a versao antiga em cache. O prefixo _ impede o
    # Streamlit de tentar fazer hash do proprio valor (basta que participe na
    # chave). Sem isto, trocar o Excel no GitHub nao atualizava a app.
    excel = pd.ExcelFile(fonte)
    dados = {}
    for chave, folha in FOLHAS.items():
        try:
            df = pd.read_excel(excel, folha)
        except Exception:
            df = pd.DataFrame()
        if COLS["data"] in df.columns:
            df[COLS["data"]] = pd.to_datetime(df[COLS["data"]], errors="coerce")
        if chave == "alvos" and not df.empty and COLS["alvo"] in df.columns:
            # A5-A8 foram tapados e substituidos por A5b-A8b (re-zerados a
            # 20/10/2025). As leituras 0,0000 dos originais a partir dessa data
            # nao sao medicoes: tratam-se como ausencia de leitura.
            msk = (df[COLS["alvo"]].astype(str).isin(["A5", "A6", "A7", "A8"])
                   & (df[COLS["data"]] >= pd.Timestamp("2025-10-20")))
            cols_num = [c for c in df.columns if c not in (
                COLS["data"], "Campanha", COLS["alvo"], COLS["edificio"],
                COLS["M0"], COLS["P0"], COLS["Z0"], "Ficheiro fonte", "Página")]
            for c in cols_num:
                if c == COLS["estado"]:
                    df.loc[msk, c] = "Sem leitura"
                elif pd.api.types.is_numeric_dtype(df[c]):
                    df.loc[msk, c] = np.nan
        dados[chave] = df
    return dados


def carregar_dados(fonte):
    # calcular uma "versao" do ficheiro para a chave de cache. Para um caminho
    # no disco (caso do deploy), usa mtime+tamanho — muda quando o Excel muda.
    # Para um upload manual, usa nome+tamanho do objeto carregado.
    versao = None
    try:
        import os
        if hasattr(fonte, "size") and hasattr(fonte, "name"):
            versao = f"upload:{fonte.name}:{fonte.size}"
        else:
            stt = os.stat(str(fonte))
            versao = f"path:{stt.st_mtime_ns}:{stt.st_size}"
    except Exception:
        versao = "sem-versao"
    return _carregar_dados_cached(fonte, versao)


def validar_colunas(df, nomes, contexto):
    if df.empty:
        st.warning(f"A folha de '{contexto}' está vazia ou não foi encontrada.")
        return False
    faltam = [n for n in nomes if n not in df.columns]
    if faltam:
        st.error(f"Em '{contexto}' faltam colunas: {faltam}. "
                 f"Verifica o bloco CONFIGURAÇÃO (COLS) ou o Excel.")
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
    (ARC e CIRCLE são essenciais: as estacas da cortina de contenção aparecem
    desenhadas como pequenos arcos/círculos.)
    Se 'layers_incluir' for dado, só devolve entidades dessas layers.
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
    """Devolve ('contenção', alçado) ou ('edifício', nome) para dar cor/forma."""
    import re
    if isinstance(nome_edificio, str) and "Alçado" in nome_edificio:
        m = re.search(r"Alçado (\w+)", nome_edificio)
        return ("contencao", m.group(1) if m else "?")
    return ("edificio", nome_edificio)


def contorno_recinto(campanha):
    """
    Constroi o contorno do recinto como a ENVOLVENTE CONVEXA (convex hull)
    dos alvos da contenção periférica. Ao contrário de ligar centroides, o
    convex hull envolve sempre os pontos por fora — nunca passa por dentro
    da nuvem, evitando a falsa impressao de alvos 'interiores'.
    Devolve (Ms, Ps, Zs) já fechado, ou None se não houver pontos/scipy.
    Tudo no sistema dos alvos — não precisa de DXF nem de alinhamento.
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
    "Edifício 1": "#d62728",
    "Edifício 2": "#2ca02c",
    "Edifício 3": "#1f77b4",
}


def _casa_edificio(fig, grp, cor, nome):
    """
    Desenha uma casa simples e MERAMENTE ILUSTRATIVA na posição de um edifício
    vizinho: um bloco (paredes) rematado por um telhado de duas águas. A posição
    e a extensão assentam nas coordenadas reais dos alvos do edifício; a forma e
    esquemática e não representa a geometria real do edifício.
    """
    import numpy as np
    x0, x1 = grp[COLS["M0"]].min(), grp[COLS["M0"]].max()
    y0, y1 = grp[COLS["P0"]].min(), grp[COLS["P0"]].max()
    zbase = grp[COLS["Z0"]].min()
    mx = max((x1 - x0) * 0.2, 2.0); my = max((y1 - y0) * 0.2, 2.0)
    x0 -= mx; x1 += mx; y0 -= my; y1 += my
    h_parede = 5.0
    h_telhado = 3.0
    zt = zbase + h_parede
    zc = zt + h_telhado

    # paredes (caixa)
    xs = [x0, x1, x1, x0, x0, x1, x1, x0]
    ys = [y0, y0, y1, y1, y0, y0, y1, y1]
    zs = [zbase, zbase, zbase, zbase, zt, zt, zt, zt]
    fig.add_trace(go.Mesh3d(
        x=xs, y=ys, z=zs,
        i=[0, 0, 0, 4, 1, 1, 2, 3, 0, 3],
        j=[1, 2, 4, 5, 2, 5, 3, 7, 3, 7],
        k=[2, 3, 5, 7, 5, 6, 7, 4, 7, 4],
        color=cor, opacity=0.30, name=nome, hoverinfo="name",
        showlegend=False, flatshading=True))

    # telhado de duas aguas (cumeeira ao meio em Y)
    ym = (y0 + y1) / 2.0
    xr = [x0, x1, x1, x0, x0, x1]
    yr = [y0, y0, y1, y1, ym, ym]
    zr = [zt, zt, zt, zt, zc, zc]
    fig.add_trace(go.Mesh3d(
        x=xr, y=yr, z=zr,
        i=[0, 1, 3, 2, 0, 1],
        j=[1, 4, 2, 5, 4, 4],
        k=[4, 5, 5, 4, 3, 0],
        color=cor, opacity=0.45, hoverinfo="skip",
        showlegend=False, flatshading=True))
    fig.add_trace(go.Scatter3d(
        x=[x0, x1], y=[ym, ym], z=[zc, zc], mode="lines",
        line=dict(color=cor, width=4), showlegend=False, hoverinfo="skip"))


def separador_3d(dados):
    alvos = dados["alvos"]
    ok = validar_colunas(
        alvos,
        [COLS["data"], COLS["alvo"], COLS["edificio"], COLS["M0"], COLS["P0"],
         COLS["Z0"], COLS["dM"], COLS["dP"], COLS["dZ"], COLS["desl_h"]],
        "Alvos topográficos",
    )
    if not ok:
        return

    st.subheader("Movimento dos alvos no espaço, com a geometria da obra")
    st.caption("O contorno castanho é a envolvente dos alvos da contenção "
               "(convex hull), que aproxima o limite do recinto de escavação. "
               "Os alvos dos edifícios vizinhos aparecem agrupados e "
               "identificados por cor. As setas mostram a direção e magnitude "
               "do deslocamento acumulado (amplificado). Todos os alvos são de "
               "periferia — na cortina de contenção ou nas fachadas vizinhas; "
               "não há instrumentos dentro da escavação. Tudo no sistema de "
               "coordenadas dos alvos, sem necessidade de DXF.")

    datas = sorted(alvos[COLS["data"]].dropna().unique())
    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        data_sel = st.select_slider(
            "Campanha", options=datas, value=datas[-1],
            format_func=lambda d: pd.to_datetime(d).strftime("%d/%m/%Y"))
    with col_b:
        fator = st.slider("Amplificação do deslocamento", 50, 2000, 500, 50,
                          help="Os deslocamentos são milimétricos e as "
                               "coordenadas em metros; amplifica-se para ver.")
    with col_c:
        mostrar_contorno = st.checkbox("Contorno do recinto", value=True)
        mostrar_caixa = st.checkbox(
            "Caixa de escavação", value=True,
            help="Desenha o volume escavado abaixo do contorno, com a "
                 "PROFUNDIDADE real de escavação do projeto (coroamento-fundo "
                 "= 16,3 m). É uma distância, não uma cota absoluta — os alvos "
                 "e o projeto usam referenciais de cota diferentes.")
        mostrar_fases = st.checkbox(
            "Fases de escavação (pisos)", value=False,
            help="Marca dentro da caixa os níveis dos pisos (-1 a -4) como "
                 "planos, a partir das cotas de projeto. São distâncias abaixo "
                 "do coroamento, invariantes ao referencial.")
        destacar_alarmes = st.checkbox(
            "Destacar alarmes/alertas", value=True,
            help="Marca a vermelho os alvos e setas em alarme e a laranja os "
                 "em alerta, segundo os critérios oficiais recalculados.")
        destacar_sc = st.checkbox(
            "Realçar fachadas do Edifício 1", value=True,
            help="Distingue a fachada frontal (A1-A4, exposta a escavação) da "
                 "lateral (A5-A8, ao mar).")
        identificar_edif = st.checkbox(
            "Identificar edifícios", value=True,
            help="Etiqueta com o nome de cada edifício, flutuando sobre os "
                 "seus alvos.")
        mostrar_casas = st.checkbox(
            "Casas dos edifícios (ilustrativo)", value=True,
            help="Desenha cada edifício vizinho como uma casa simples "
                 "(bloco + telhado). Posição real dos alvos; forma esquemática "
                 "— não representa a geometria real do edifício.")

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
        # caixa de escavacao: desce da envolvente uma PROFUNDIDADE real.
        # A profundidade (coroamento - fundo) e uma DISTANCIA, invariante ao
        # referencial; a cota absoluta nao (alvos e projeto usam sistemas de
        # cota diferentes). Por isso usamos a profundidade, nao a cota do fundo.
        # Topo ancorado ao alvo mais ALTO da contencao (o mais proximo do
        # coroamento), nao ao mais baixo — assim a caixa representa melhor a
        # altura escavada a partir do coroamento.
        if mostrar_caixa:
            prof = COTA_COROAMENTO_PADRAO - COTA_FUNDO_ESCAVACAO   # 16,3 m
            topo_z = max(Zs)               # alvo de contencao mais alto ~ coroamento
            base_z = topo_z - prof
            # paredes verticais (quads) ao longo do contorno
            for i in range(len(Ms) - 1):
                fig.add_trace(go.Scatter3d(
                    x=[Ms[i], Ms[i+1], Ms[i+1], Ms[i], Ms[i]],
                    y=[Ps[i], Ps[i+1], Ps[i+1], Ps[i], Ps[i]],
                    z=[topo_z, topo_z, base_z, base_z, topo_z],
                    mode="lines", line=dict(color="peru", width=1),
                    surfaceaxis=2, surfacecolor="rgba(210,180,140,0.18)",
                    showlegend=False, hoverinfo="skip",
                ))
            # fundo da escavacao — plano preenchido (da volume ao fundo)
            fig.add_trace(go.Scatter3d(
                x=list(Ms), y=list(Ps), z=[base_z] * len(Ms),
                mode="lines", line=dict(color="peru", width=3),
                surfaceaxis=2, surfacecolor="rgba(180,150,110,0.30)",
                name=f"Fundo de escavação (−{prof:.1f} m do coroamento)",
                hoverinfo="skip",
            ))
            # planos das FASES de escavacao (cotas dos pisos, como distancias
            # abaixo do coroamento — invariante ao referencial)
            if mostrar_fases:
                for nome, cota in COTAS_PISOS:
                    d_piso = COTA_COROAMENTO_PADRAO - cota   # prof. abaixo coroamento
                    if 0 < d_piso < prof:                    # so os que estao dentro
                        z_piso = topo_z - d_piso
                        fig.add_trace(go.Scatter3d(
                            x=list(Ms), y=list(Ps), z=[z_piso] * len(Ms),
                            mode="lines",
                            line=dict(color="rgba(90,90,90,0.55)", width=1),
                            name=f"{nome} (−{d_piso:.1f} m)",
                            hovertemplate=f"{nome}<br>{d_piso:.1f} m abaixo do "
                                          f"coroamento<extra></extra>",
                        ))

    # ---- alvos por grupo (cor por edificio; contencao a laranja) ---------
    # acumuladores para desenhar TODAS as setas em poucos traces (leve)
    seg_x, seg_y, seg_z, seg_cor = [], [], [], []
    cone_x, cone_y, cone_z, cone_u, cone_v, cone_w, cone_cor = ([] for _ in range(7))
    COR_ESTADO = {"Alarme": "#c0140f", "Alerta": "#e67e00", "Regular": "#1f9e55"}

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
        fachadas = grp["Fachada"].to_numpy()

        e_edificio1 = isinstance(chave, str) and "Edifício 1" in chave
        if tipo == "edificio":
            cor = CORES_EDIFICIO.get(chave, "#7f7f7f")
            nome_leg = chave
        else:
            cor = "#ff7f0e"
            nome_leg = f"Contenção — Alçado {etiqueta}"

        # simbolo por fachada do Edifício 1 (frontal vs lateral)
        if destacar_sc and e_edificio1:
            simbolos = ["diamond" if f == "Frente escavação" else "circle"
                        for f in fachadas]
        else:
            simbolos = "circle"

        # contorno do marcador por estado (cor por ponto; largura escalar)
        if destacar_alarmes:
            cor_borda = [COR_ESTADO.get(e, "rgba(0,0,0,0.2)") for e in estados]
            larg_borda = 4 if any(e in ("Alarme", "Alerta") for e in estados) else 1
        else:
            cor_borda = "rgba(0,0,0,0.2)"
            larg_borda = 1

        marker = dict(size=6, color=cor, symbol=simbolos,
                      line=dict(color=cor_borda, width=larg_borda))
        cd = np.column_stack([dh, estados, fachadas])
        fig.add_trace(go.Scatter3d(
            x=x0 + dx, y=y0 + dy, z=z0 + dz, mode="markers+text",
            marker=marker,
            text=nomes, textposition="top center", textfont=dict(size=8),
            name=nome_leg, customdata=cd,
            hovertemplate="Alvo %{text}<br>Desl. h: %{customdata[0]:.1f} mm"
                          "<br>Estado: %{customdata[1]}"
                          "<br>%{customdata[2]}"
                          "<extra>" + nome_leg + "</extra>",
        ))

        # acumular setas (segmento + cone na ponta), cor por estado
        for i in range(len(x0)):
            c = COR_ESTADO.get(estados[i], "#888888") if destacar_alarmes else "crimson"
            seg_x += [x0[i], x0[i] + dx[i], None]
            seg_y += [y0[i], y0[i] + dy[i], None]
            seg_z += [z0[i], z0[i] + dz[i], None]
            seg_cor.append(c)
            cone_x.append(x0[i] + dx[i]); cone_y.append(y0[i] + dy[i])
            cone_z.append(z0[i] + dz[i])
            cone_u.append(dx[i]); cone_v.append(dy[i]); cone_w.append(dz[i])
            cone_cor.append(c)

        # etiqueta identificadora do edificio, sobre o centro dos seus alvos
        if identificar_edif and tipo == "edificio" and len(x0):
            fig.add_trace(go.Scatter3d(
                x=[x0.mean()], y=[y0.mean()], z=[z0.max() + 3],
                mode="text", text=[f"<b>{chave}</b>"],
                textfont=dict(size=12, color=cor),
                showlegend=False, hoverinfo="skip"))

        # casa simples ilustrativa do edificio vizinho
        if mostrar_casas and tipo == "edificio" and len(x0):
            _casa_edificio(fig, grp, cor, chave)

    # desenhar todas as hastes das setas de uma vez (por cor, para poucos traces)
    for c in set(seg_cor):
        xs, ys, zs = [], [], []
        for j, cc in enumerate(seg_cor):
            if cc == c:
                xs += seg_x[3*j:3*j+3]; ys += seg_y[3*j:3*j+3]; zs += seg_z[3*j:3*j+3]
        fig.add_trace(go.Scatter3d(x=xs, y=ys, z=zs, mode="lines",
                                   line=dict(color=c, width=4),
                                   showlegend=False, hoverinfo="skip"))
    # pontas das setas (cones), num unico trace
    if cone_x:
        fig.add_trace(go.Cone(
            x=cone_x, y=cone_y, z=cone_z, u=cone_u, v=cone_v, w=cone_w,
            sizemode="absolute", sizeref=1.2, anchor="tip",
            showscale=False, colorscale=[[0, "#555"], [1, "#555"]],
            hoverinfo="skip", showlegend=False, opacity=0.9,
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
    c2.metric("Desl. horizontal máx. (mm)", f"{np.nanmax(desl_h_all):.1f}")
    c3.metric("Em alarme", n_alarme)
    c4.metric("Em alerta", n_alerta)
    idx = int(np.nanargmax(desl_h_all))
    # leitura frente vs lateral do Edifício 1, se houver dados
    sc = campanha[campanha["Fachada"] != ""]
    linha_sc = ""
    if len(sc):
        frente = sc[sc["Fachada"] == "Frente escavação"][COLS["desl_h"]]
        lateral = sc[sc["Fachada"] == "Lateral (mar)"][COLS["desl_h"]]
        if len(frente) and len(lateral):
            linha_sc = (f" No Edifício 1, a fachada frontal (losangos, média "
                        f"{frente.mean():.0f} mm) move-se mais que a lateral "
                        f"(círculos, {lateral.mean():.0f} mm) — coerente com a "
                        f"exposição direta a escavação.")
    st.caption(
        f"O alvo mais afetado ({nomes_all[idx]}, {np.nanmax(desl_h_all):.1f} mm) "
        f"pertence a: {edif_all[idx]}. Setas e contornos: vermelho = alarme, "
        f"laranja = alerta, verde = regular (critérios oficiais recalculados). "
        f"A caixa mostra a profundidade real de escavação (16,3 m, do projeto) "
        f"como distância abaixo da cortina — os alvos e o projeto usam "
        f"referenciais de cota diferentes, por isso e profundidade, não cota "
        f"absoluta.{linha_sc}")


# =========================================================================
# SEPARADOR 2 — INCLINOMETROS
# =========================================================================
def separador_inclinometros(dados, limiar_vel, fator_acel):
    perfis, resumo = dados["perfis"], dados["resumo"]
    ok = validar_colunas(perfis, [COLS["data"], COLS["inclinometro"],
                                  COLS["profundidade"], COLS["desl_total"]],
                         "Inclinómetros / perfis") \
        and validar_colunas(resumo, [COLS["data"], COLS["inclinometro"],
                                     COLS["desl_max_global"], COLS["prof_do_max"]],
                            "Inclinómetros / resumo")
    if not ok:
        return

    inc = st.selectbox("Inclinómetro",
                       sorted(perfis[COLS["inclinometro"]].dropna().unique()))
    p_inc = perfis[perfis[COLS["inclinometro"]] == inc].copy()
    r_inc = resumo[resumo[COLS["inclinometro"]] == inc].copy().sort_values(COLS["data"])
    datas_inc = sorted(p_inc[COLS["data"]].dropna().unique())
    if not datas_inc:
        st.warning("Sem leituras com data válida.")
        return

    tab_perfil, tab_evol = st.tabs(
        ["📐 Perfil deformado", "📈 Evolução do deslocamento"])
    with tab_perfil:
        st.subheader("Perfil deformado")
        st.caption("Deslocamento acumulado ao longo da profundidade. Base fixa.")
        idx = sorted(set([0, len(datas_inc) // 2, len(datas_inc) - 1]))
        sel = st.multiselect("Leituras", datas_inc,
                             default=[datas_inc[i] for i in idx],
                             format_func=lambda d: pd.to_datetime(d).strftime("%d/%m/%Y"))

        # opcao de sobrepor a geologia de uma sondagem (peca-chave do back-analysis)
        geo_on = st.checkbox("Sobrepor geologia + SPT da sondagem", value=False,
                             help="Mostra a litologia, o nível freático e o "
                                  "perfil SPT de uma sondagem no mesmo eixo de "
                                  "profundidade, para relacionar a deformação "
                                  "com a resistência do terreno.")
        sond_sel = None
        if geo_on:
            # sugestao por proximidade (inferida das plantas)
            meta = INC_META.get(inc)
            sonds = list(GEO_LITOLOGIA.keys())
            default_idx = 0
            if meta and meta["sondagem"] in sonds:
                default_idx = sonds.index(meta["sondagem"])
            sond_sel = st.selectbox(
                "Sondagem de referência", sonds, index=default_idx)
            if meta:
                rumo = _rumo_cardeal(meta["azimute"])
                sugerida = meta["sondagem"]
                nota = (f"Sugerida por proximidade: **{sugerida}** "
                        f"(confiança {meta['confianca']}; {inc} fica em "
                        f"{meta['posicao']}). Eixo A+ orientado a "
                        f"{meta['azimute']}° ({rumo}). ")
                if sond_sel != sugerida:
                    nota += f"Estas a ver **{sond_sel}**, diferente da sugerida."
                st.caption(nota)
                st.caption("Associação inclinómetro-sondagem inferida da "
                           "sobreposição das plantas — a confirmar com a "
                           "instrumentação.")

        fig = go.Figure()

        # se geologia ligada, desenhar faixas litologicas de fundo + SPT
        if geo_on and sond_sel:
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
            # base da sondagem (abaixo disto nao ha dado geologico)
            base_sond = GEO_LITOLOGIA[sond_sel][-1][1]
            prof_inc_max = float(p_inc[COLS["profundidade"]].max())
            if prof_inc_max > base_sond + 0.5:
                fig.add_hline(y=base_sond, line=dict(color="gray", width=1, dash="dot"),
                              annotation_text=f"base {sond_sel}",
                              annotation_position="left")
            # perfil SPT sobreposto num eixo X secundario (N pancadas)
            ensaios = GEO_SPT.get(sond_sel, [])
            if ensaios:
                sp_prof = [e[0] for e in ensaios]
                sp_n = [e[1] for e in ensaios]
                fig.add_trace(go.Scatter(
                    x=sp_n, y=sp_prof, mode="lines+markers",
                    name=f"SPT {sond_sel} (N)", xaxis="x2",
                    line=dict(color="rgba(70,70,70,0.7)", width=1.5, dash="dot"),
                    marker=dict(size=5, color="rgba(70,70,70,0.8)")))
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
        # eixo X secundario para o SPT (0-65), no topo
        if geo_on and sond_sel and GEO_SPT.get(sond_sel):
            fig.update_layout(xaxis2=dict(title="N (SPT)", overlaying="x",
                                          side="top", range=[0, 65],
                                          showgrid=False))
        # ecra cheio: legenda vertical a direita (litologia, campanhas e fases
        # numa so coluna). Ha largura de sobra, por isso nao precisa de ir para
        # baixo. As 20 campanhas continuam todas clicaveis para isolar.
        fig.update_layout(
            height=760,
            legend=dict(
                title="Leitura / geologia",
                orientation="v", yanchor="top", y=1,
                xanchor="left", x=1.02,
                font=dict(size=11),
                traceorder="normal"),
            margin=dict(r=60, t=60, b=40))
        st.plotly_chart(fig, use_container_width=True)
        if geo_on and sond_sel:
            st.caption(f"Litologia, NF e SPT da sondagem {sond_sel} sobrepostos "
                       f"(SPT no eixo de cima). A leitura central do "
                       f"back-analysis: ve se o 'joelho' de maior deformação do "
                       f"perfil coincide com uma subida do SPT (grés a "
                       f"consolidar) ou com o nível freático. Onde o "
                       f"inclinómetro passa da base da sondagem, não há dado "
                       f"geológico.")

    with tab_evol:
        st.subheader("Evolução do deslocamento")
        st.caption("Máximo global vs. profundidade fixa.")
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
                                  mode="lines+markers", name="Máximo global"))
        fig2.add_trace(go.Scatter(x=s_fix[COLS["data"]], y=s_fix[COLS["desl_total"]],
                                  mode="lines+markers", name=f"A {prof_fixa:.1f} m"))
        fig2.update_xaxes(title="Data")
        fig2.update_yaxes(title="Deslocamento (mm)")
        fig2.update_layout(
            height=780,
            legend=dict(title="Série", orientation="v",
                        yanchor="top", y=1, xanchor="left", x=1.02,
                        font=dict(size=11)),
            margin=dict(r=60, t=60, b=40))
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()
    st.subheader("Velocidade e sinais precursores")
    base = st.radio("Série de base", ["Máximo global", f"Profundidade fixa ({prof_fixa:.1f} m)"],
                    horizontal=True)
    if base == "Máximo global":
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
    c1.metric("Últ. desloc. máx. (mm)",
              f"{s_max[COLS['desl_max_global']].iloc[-1]:.2f}" if len(s_max) else "-")
    vmax = vdf["velocidade"].max()
    c2.metric("Velocidade máx. (mm/dia)", f"{vmax:.3f}" if pd.notna(vmax) else "-")
    c3.metric("Precursores", n)
    (st.warning if n else st.success)(
        f"{n} leitura(s) com aceleração acima dos critérios." if n
        else "Nenhuma aceleração acima dos critérios.")


# =========================================================================
# SEPARADOR 3 — ALVOS (2D, series temporais)
# =========================================================================
def _plotar_alcado_esquematico(sub_edi, edi, selecionados, cores_estado):
    """
    Desenha um alçado ESQUEMÁTICO de um edifício a partir das coordenadas
    reais dos alvos. Usa a coordenada ao longo da fachada (M ou P, conforme
    a orientação dominante) no eixo horizontal e o Z no eixo vertical. As
    POSIÇÕES RELATIVAS são fieis aos dados; a ESCALA e esquemática (não se
    afirmam cotas absolutas). Devolve uma figura plotly ou None.
    """
    import plotly.graph_objects as go
    d = sub_edi.dropna(subset=[COLS["M0"], COLS["P0"], COLS["Z0"]]).copy()
    if len(d) < 2:
        return None
    # orientacao da fachada: escolher o eixo (M ou P) com maior amplitude
    span_m = d[COLS["M0"]].max() - d[COLS["M0"]].min()
    span_p = d[COLS["P0"]].max() - d[COLS["P0"]].min()
    eixo = COLS["M0"] if span_m >= span_p else COLS["P0"]
    horiz_lbl = "Posição ao longo da fachada (m, relativo)"

    fig = go.Figure()
    # moldura da fachada (retangulo de fundo)
    x0, x1 = d[eixo].min(), d[eixo].max()
    z0, z1 = d[COLS["Z0"]].min(), d[COLS["Z0"]].max()
    mx = (x1 - x0) * 0.15 + 0.5
    mz = (z1 - z0) * 0.15 + 0.5
    fig.add_shape(type="rect", x0=x0 - mx, x1=x1 + mx, y0=z0 - mz, y1=z1 + mz,
                  line=dict(color="#999", width=1),
                  fillcolor="rgba(200,200,200,0.12)", layer="below")

    for _, r in d.iterrows():
        a = str(r[COLS["alvo"]])
        est = r.get("Estado calculado", "Regular")
        cor = cores_estado.get(est, "#1f9e55")
        realce = a in selecionados
        fig.add_trace(go.Scatter(
            x=[r[eixo]], y=[r[COLS["Z0"]]], mode="markers+text",
            marker=dict(size=20 if realce else 13, color=cor,
                        line=dict(color="black" if realce else "white",
                                  width=2 if realce else 1),
                        symbol="star" if realce else "circle"),
            text=[a], textposition="middle right" if realce else "top center",
            textfont=dict(size=12 if realce else 9,
                          color="black" if realce else "#444"),
            showlegend=False,
            hovertemplate=f"{a}<br>Estado: {est}<extra></extra>",
        ))
    fig.update_xaxes(title=horiz_lbl, showticklabels=False)
    fig.update_yaxes(title="Altura relativa (Z)", showticklabels=False)
    fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=30),
                      plot_bgcolor="white")
    return fig


def _slug_edificio(edi):
    """Nome de ficheiro seguro (sem acentos, minusculas) para a foto do edifício."""
    import unicodedata
    txt = unicodedata.normalize("NFKD", str(edi)).encode("ascii", "ignore").decode()
    slug = "".join(c if c.isalnum() else "_" for c in txt).strip("_").lower()
    while "__" in slug:
        slug = slug.replace("__", "_")
    return slug[:60]


def mostrar_localizacao_alvo(sub_edi, edi, selecionados, cores_estado):
    """
    Mostra onde estão os alvos selecionados no edifício. Prioridade:
      1) foto real do relatório, se existir em fotos_alvos/<slug>.{png,jpg}
      2) alçado esquemático das coordenadas reais (fallback)
    A foto e propriedade do relatório de instrumentação (33GRADOS) — creditar.
    """
    import os
    slug = _slug_edificio(edi)
    # pasta de fotos ancorada ao diretorio do script (robusto ao CWD do deploy)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    foto = None
    for ext in ("png", "jpg", "jpeg"):
        caminho = os.path.join(base_dir, "fotos_alvos", f"{slug}.{ext}")
        if os.path.exists(caminho):
            foto = caminho
            break
    if foto:
        st.image(foto, use_container_width=True,
                 caption=f"Localização dos alvos — {edi}. "
                         f"Fonte: relatório de instrumentação (33GRADOS).")
        return
    # fallback: esquema das coordenadas
    fig = _plotar_alcado_esquematico(sub_edi, edi, selecionados, cores_estado)
    if fig is not None:
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Esquema a partir das coordenadas reais dos alvos (posições "
                   "relativas fieis; escala esquemática). O alvo selecionado "
                   "aparece em estrela. Para uma foto real, coloca a imagem em "
                   f"'fotos_alvos/{slug}.png'.")
    else:
        st.caption("Sem coordenadas suficientes para esquematizar este edifício.")


# -------------------------------------------------------------------------
# PERIODOS DE ESCAVACAO — os seis periodos da leitura global da obra,
# delimitados pelo inicio da escavacao de cada nivel e pelo fim da escavacao
# segundo o PLANO DE TRABALHOS IMPACTADO (ver FASES_COMPARACAO e
# ESCAVACAO_COTAS). Sao uma referencia comum: as frentes avancaram
# desfasadas (a do alcado GH adiantada face ao plano), pelo que nao
# representam o estado de cada frente.
# (rotulo curto no grafico, descricao, inicio ISO, fim ISO | None = ultima leitura)
# -------------------------------------------------------------------------
PERIODOS_ESCAVACAO = [
    ("Antes da<br>escavação", "antes da escavação geral", "2025-05-12", "2025-07-24"),
    ("Pisos 2 e 1", "pisos 2 e 1", "2025-07-24", "2025-11-12"),
    ("Piso -1", "piso -1, com a escavação de 4,95 m da frente poente",
     "2025-11-12", "2025-12-15"),
    ("Piso -2", "piso -2 e escavação até à viga do piso -3", "2025-12-15", "2026-02-25"),
    ("Piso -3<br>e fundo", "piso -3 e fundo", "2026-02-25", "2026-04-16"),
    ("Laje de fundo<br>e estrutura", "laje de fundo e estrutura", "2026-04-16", None),
]
_PERIODOS_LINHAS = 3   # linhas do rotulo mais alto (numero + 2 linhas de texto)

# cores dos alvos: cada grupo de alvos a mesma posicao em planta partilha a
# cor (tom escuro no superior, claro no inferior); alvos isolados, tom escuro
_PARES_CORES = [("#1f4e9c", "#86aee6"), ("#b2182b", "#f4a0a0"),
                ("#1b7837", "#8fd18f"), ("#6a3d9a", "#c2a5e0"),
                ("#d95f02", "#fdb46b"), ("#01665e", "#76cbbf"),
                ("#8c510a", "#dcbc7c"), ("#404040", "#a6a6a6")]
# faixas de comparacao com outros grupos (preenchimento, contorno)
_COMP_CORES = [("rgba(120,120,120,0.20)", "rgba(95,95,95,0.75)"),
               ("rgba(176,141,87,0.24)", "rgba(140,105,50,0.85)"),
               ("rgba(90,150,170,0.20)", "rgba(50,110,130,0.80)"),
               ("rgba(170,110,160,0.20)", "rgba(130,70,120,0.80)")]


def _chave_natural(s):
    import re as _re
    return [(0, int(t), "") if t.isdigit() else (1, 0, t)
            for t in _re.split(r"(\d+)", str(s)) if t != ""]


def _num_pt(v, casas=1):
    """Número com vírgula decimal (o tradutor passa-o a ponto em inglês)."""
    return f"{v:.{casas}f}".replace("-", "−").replace(".", ",")


def rotulos_posicao_alvos(sub_grupo, dist_max=1.0, dz_min=0.5):
    """
    Identifica os PARES de alvos instalados à mesma posição em planta (a
    menos de `dist_max` m), a cotas diferentes (desnível > `dz_min` m) e com
    leituras no mesmo período (um alvo de substituição não forma par com o
    que substituiu). Devolve {alvo: (sufixo, cor)}: nos pares, o sufixo é
    «(sup.)» ou «(inf.)» e a cor é a mesma, escura no superior e clara no
    inferior; os restantes alvos (isolados ou em alinhamentos de três ou mais
    prismas, como nos alçados) têm cor própria e não têm sufixo.
    """
    d = sub_grupo.dropna(subset=[COLS["M0"], COLS["P0"], COLS["Z0"]])
    lido = d[d[COLS["desl_h"]].notna()]
    info = {}
    for a, g in d.groupby(COLS["alvo"]):
        gl = lido[lido[COLS["alvo"]] == a][COLS["data"]]
        r = g.iloc[0]
        info[a] = (float(r[COLS["M0"]]), float(r[COLS["P0"]]), float(r[COLS["Z0"]]),
                   gl.min() if len(gl) else pd.NaT, gl.max() if len(gl) else pd.NaT)
    nomes = sorted(info, key=_chave_natural)
    pai = {a: a for a in nomes}

    def raiz(a):
        while pai[a] != a:
            a = pai[a]
        return a
    for i, a in enumerate(nomes):
        ma, pa, za, a0, a1 = info[a]
        for b in nomes[i + 1:]:
            mb, pb, zb, b0, b1 = info[b]
            if np.hypot(ma - mb, pa - pb) >= dist_max or abs(za - zb) <= dz_min:
                continue
            if pd.isna(a0) or pd.isna(b0) or a1 < b0 or b1 < a0:
                continue
            pai[raiz(b)] = raiz(a)
    grupos = {}
    for a in nomes:
        grupos.setdefault(raiz(a), []).append(a)
    out = {}
    ordem = sorted(grupos.values(), key=lambda m: _chave_natural(m[0]))
    k = 0
    for membros in ordem:
        if len(membros) == 2:
            esc, cla = _PARES_CORES[k % len(_PARES_CORES)]
            sup, inf = sorted(membros, key=lambda a: -info[a][2])
            out[sup] = (" (sup.)", esc)
            out[inf] = (" (inf.)", cla)
            k += 1
            continue
        for a in membros:
            out[a] = ("", _PARES_CORES[k % len(_PARES_CORES)][0])
            k += 1
    return out


def adicionar_periodos_escavacao(fig, dt_min, dt_max):
    """Faixas alternadas com os períodos de escavação e o rótulo de cada um
    por cima da área de dados. Devolve [(n, descrição, início, fim|None)]."""
    dt_min, dt_max = pd.to_datetime(dt_min), pd.to_datetime(dt_max)
    vis = []
    for i, (rot, desc, ini, fim) in enumerate(PERIODOS_ESCAVACAO):
        t0 = pd.to_datetime(ini)
        t1 = pd.to_datetime(fim) if fim else dt_max
        a, b = max(t0, dt_min), min(t1, dt_max)
        if b <= a:
            continue
        n = i + 1
        if n % 2 == 1:
            fig.add_shape(type="rect", xref="x", yref="y domain", x0=a, x1=b,
                          y0=0, y1=1, fillcolor="rgba(110,110,110,0.10)",
                          line_width=0, layer="below")
        if t0 > dt_min:
            fig.add_shape(type="line", xref="x", yref="y domain", x0=t0, x1=t0,
                          y0=0, y1=1, line=dict(color="rgba(90,90,90,0.45)", width=1),
                          layer="below")
        # numero na 1.a linha e texto por baixo: todos os numeros ao mesmo
        # nivel (ancorados pelo topo, acima da area de dados)
        fs = 10
        fig.add_annotation(x=a + (b - a) / 2, y=1.0, xref="x", yref="y domain",
                           yanchor="top", xanchor="center", showarrow=False,
                           yshift=int(_PERIODOS_LINHAS * 1.3 * fs) + 6,
                           text=f"<b>{n}</b><br>{rot}", align="center",
                           font=dict(size=fs, color="#4d4d4d"), hovertext=desc)
        vis.append((n, desc, t0, pd.to_datetime(fim) if fim else None))
    return vis


def _fs_periodos():
    return 10


def legenda_periodos(vis):
    if not vis:
        return
    itens = []
    for n, desc, t0, t1 in vis:
        fim = t1.strftime("%d/%m/%Y") if t1 is not None else "última leitura"
        itens.append(f"**{n}.** {desc} ({t0.strftime('%d/%m/%Y')}–{fim})")
    st.caption("Períodos de escavação, segundo o plano de trabalhos impactado "
               "(referência comum; as frentes avançaram desfasadas): "
               + " · ".join(itens))


def _valor_na_data(s, col, d, tol_dias=10):
    """Última leitura até à data `d` (no máximo `tol_dias` antes)."""
    ss = s[(s[COLS["data"]] <= d) & s[col].notna()].sort_values(COLS["data"])
    if ss.empty:
        return None, None
    t = pd.to_datetime(ss[COLS["data"]].iloc[-1])
    if (d - t).days > tol_dias:
        return None, None
    return float(ss[col].iloc[-1]), t


def faixa_grupo(alvos, grupo, col, j0, j1):
    """Mínimo e máximo, em cada campanha, dos alvos de `grupo` zerados na
    primeira campanha do grupo (os de substituição ficam de fora)."""
    g = alvos[(alvos[COLS["edificio"]] == grupo) & alvos[col].notna()]
    if g.empty:
        return None, []
    ini = g.groupby(COLS["alvo"])[COLS["data"]].min()
    orig = ini[ini <= ini.min() + pd.Timedelta(days=3)].index
    g = g[g[COLS["alvo"]].isin(orig)]
    g = g[(g[COLS["data"]] >= j0) & (g[COLS["data"]] <= j1)]
    if g.empty:
        return None, []
    agg = (g.groupby(COLS["data"])[col].agg(["min", "max"]).reset_index()
           .sort_values(COLS["data"]))
    return agg, sorted(orig, key=_chave_natural)


def _texto_lista_alvos(nomes):
    if not nomes:
        return ""
    if len(nomes) <= 2:
        return " e ".join(nomes)
    return f"{nomes[0]}–{nomes[-1]}"


def grafico_alvos_tempo(sub, alvos_todos, sel, col, rotulos, *, titulo_y,
                        limiares, fases, datas_ref, comp, marcar_max, modo_max,
                        granul, j0, j1, rezer, altura=780, projeto=None):
    """Série temporal de uma componente (δH ou ΔZ) dos alvos escolhidos, com
    as sobreposições opcionais. Devolve (figura, fases visíveis, períodos)."""
    fig = go.Figure()
    # faixas dos outros grupos primeiro, para ficarem por baixo das séries
    for k, grupo in enumerate(comp):
        agg, nomes = faixa_grupo(alvos_todos, grupo, col, j0, j1)
        if agg is None or agg.empty:
            continue
        cf, cl = _COMP_CORES[k % len(_COMP_CORES)]
        nome_faixa = f"{_texto_lista_alvos(nomes)} (mín.–máx.)"
        xs = [pd.to_datetime(v) for v in agg[COLS["data"]].tolist()]
        fig.add_trace(go.Scatter(
            x=xs, y=[float(v) for v in agg["min"].tolist()], mode="lines",
            line=dict(color=cl, width=0.8), legendgroup=f"comp{k}",
            showlegend=False, name=nome_faixa,
            hovertemplate=f"{grupo}<br>mín.: %{{y:.1f}} mm<extra></extra>"))
        fig.add_trace(go.Scatter(
            x=xs, y=[float(v) for v in agg["max"].tolist()], mode="lines",
            line=dict(color=cl, width=0.8), fill="tonexty", fillcolor=cf,
            legendgroup=f"comp{k}", name=nome_faixa,
            legendgrouptitle_text=("Outros grupos" if k == 0 else None),
            hovertemplate=f"{grupo}<br>máx.: %{{y:.1f}} mm<extra></extra>"))
    # series dos alvos escolhidos
    for a in sel:
        s = sub[sub[COLS["alvo"]] == a].sort_values(COLS["data"])
        suf, cor = rotulos.get(a, ("", None))
        fig.add_trace(go.Scatter(
            x=s[COLS["data"]], y=s[col], mode="lines+markers", name=a + suf,
            legendgroup="alvos",
            line=dict(color=cor, width=2) if cor else None,
            marker=dict(size=5, color=cor) if cor else None))
    # limiares
    (va, ta), (vm, tm) = limiares
    # rotulos dentro da area de dados, a esquerda (a direita fica a legenda)
    fig.add_hline(y=va, line_dash="dash", line_color="orange",
                  annotation_text=ta, annotation_position="top left",
                  annotation_font=dict(size=11, color="#b36b00"))
    fig.add_hline(y=vm, line_dash="dash", line_color="red",
                  annotation_text=tm, annotation_position="top left",
                  annotation_font=dict(size=11, color="#b00000"))
    # re-zeragem dos alvos de substituicao escolhidos (so quando existem)
    for t, nomes in rezer:
        fig.add_shape(type="line", xref="x", yref="y domain", x0=t, x1=t,
                      y0=0, y1=1, line=dict(color="gray", width=1.5, dash="dot"))
        fig.add_annotation(x=t, y=0, xref="x", yref="y domain", xanchor="left",
                           yanchor="bottom", showarrow=False, xshift=3,
                           text=f"Re-zeragem {_texto_lista_alvos(nomes)}",
                           font=dict(size=10, color="#555"),
                           bgcolor="rgba(255,255,255,0.7)")
    # datas de referencia: linha, data e valor de cada alvo. O rotulo (na
    # vertical, junto a linha) vai para cima ou para baixo, onde nao houver
    # series nessa data.
    _ys = []
    for tr in fig.data:
        try:
            _ys += [float(v) for v in tr.y if v is not None and pd.notna(v)]
        except TypeError:
            pass
    _ys += [limiares[0][0], limiares[1][0]]
    _ylo, _yhi = min(_ys), max(_ys)
    _amp = (_yhi - _ylo) or 1.0
    for i_d, d in enumerate(datas_ref):
        fig.add_shape(type="line", xref="x", yref="y domain", x0=d, x1=d,
                      y0=0, y1=1, line=dict(color="#333333", width=1.2, dash="dash"))
        # com a data seguinte muito proxima, este rotulo passa para a esquerda
        _esq = (i_d + 1 < len(datas_ref)
                and (datas_ref[i_d + 1] - d) < 0.05 * (j1 - j0))
        _perto = []
        for tr in fig.data:
            try:
                for xv, yv in zip(tr.x, tr.y):
                    if (xv is not None and yv is not None and pd.notna(yv)
                            and abs((pd.to_datetime(xv) - d).days) <= 20):
                        _perto.append(float(yv))
            except TypeError:
                pass
        _topo_ocup = any(v > _yhi - 0.25 * _amp for v in _perto)
        _base_ocup = any(v < _ylo + 0.25 * _amp for v in _perto)
        _em_baixo = _topo_ocup and not _base_ocup
        fig.add_annotation(x=d, y=(0 if _em_baixo else 1), xref="x", yref="y domain",
                           xanchor=("right" if _esq else "left"),
                           yanchor=("bottom" if _em_baixo else "top"),
                           showarrow=False, xshift=(-2 if _esq else 2), textangle=-90,
                           text=d.strftime("%d/%m/%Y"),
                           font=dict(size=10, color="#333333"),
                           bgcolor="rgba(255,255,255,0.75)")
        for a in sel:
            s = sub[sub[COLS["alvo"]] == a]
            v, t = _valor_na_data(s, col, d)
            if v is None:
                continue
            suf, cor = rotulos.get(a, ("", None))
            fig.add_trace(go.Scatter(
                x=[t], y=[v], mode="markers", showlegend=False,
                marker=dict(size=11, symbol="circle-open", color=cor or "#333333",
                            line=dict(width=2)),
                hovertemplate=(f"{a}{suf} — {t.strftime('%d/%m/%Y')}: "
                               f"%{{y:.1f}} mm<extra>Data de referência</extra>")))
    # estimativa do projeto (so no grafico horizontal dos alcados): linha e
    # data em que o maximo dos alvos escolhidos a ultrapassou
    if projeto is not None:
        _sp = sub[sub[COLS["alvo"]].isin(sel)]
        t_p, v_p = data_ultrapassagem(_sp[COLS["data"]], _sp[col], projeto)
        _nm = f"Estimativa do projeto: {_num_pt(projeto, 0)} mm"
        if t_p is not None:
            _nm += f"<br>(ultrapassada a {t_p.strftime('%d/%m/%Y')})"
        fig.add_trace(go.Scatter(
            x=[j0, j1], y=[projeto, projeto], mode="lines", legendgroup="proj",
            name=_nm, line=dict(color="#2e7d32", width=2, dash="dashdot"),
            hoverinfo="skip"))
        if t_p is not None:
            fig.add_trace(go.Scatter(
                x=[t_p], y=[v_p], mode="markers", legendgroup="proj",
                showlegend=False,
                marker=dict(symbol="diamond", size=13, color="#2e7d32",
                            line=dict(color="white", width=1.5)),
                hovertemplate=(f"Estimativa do projeto ultrapassada a "
                               f"{t_p.strftime('%d/%m/%Y')}: %{{y:.1f}} mm"
                               "<extra></extra>")))
    # maximo dos alvos escolhidos no periodo visivel
    if marcar_max and sel:
        s = sub[sub[COLS["alvo"]].isin(sel) & sub[col].notna()]
        if not s.empty:
            idx = s[col].idxmax() if modo_max == "max" else s[col].abs().idxmax()
            r = s.loc[idx]
            v = float(r[col])
            t = pd.to_datetime(r[COLS["data"]])
            # estrela no ponto e valor na legenda (sem caixa sobre as séries)
            pref = "Máx. δH" if modo_max == "max" else "Máx. |ΔZ|"
            fig.add_trace(go.Scatter(
                x=[t], y=[v], mode="markers", legendgroup="max",
                name=(f"{pref} {r[COLS['alvo']]}: {_num_pt(v)} mm"
                      f"<br>({t.strftime('%d/%m/%Y')})"),
                marker=dict(symbol="star", size=16, color="#ffd400",
                            line=dict(color="#222222", width=1.2)),
                hovertemplate=(f"{r[COLS['alvo']]} — {t.strftime('%d/%m/%Y')}: "
                               f"%{{y:.1f}} mm<extra>{pref}</extra>")))
    # faseamento
    fases_vis, periodos_vis = [], []
    if fases == "Períodos de escavação":
        periodos_vis = adicionar_periodos_escavacao(fig, j0, j1)
    elif fases == "Fases da obra (plano)":
        fases_vis = adicionar_fases_obra(fig, j0, j1)
    fig.update_xaxes(title="Data", type="date")
    fig.update_yaxes(title=titulo_y)
    configurar_eixo_tempo(fig, granul)
    _mt = 60
    if periodos_vis:
        _mt = max(60, int(_PERIODOS_LINHAS * 1.3 * _fs_periodos()) + 22)
    fig.update_layout(
        height=altura,
        margin=dict(r=60, t=_mt, b=40),
        legend=dict(orientation="v", yanchor="top", y=1, traceorder="grouped",
                    xanchor="left", x=1.02, font=dict(size=11)))
    return fig, fases_vis, periodos_vis


def tabela_datas_referencia(sub, sel, datas_ref, col, rotulos):
    """Valores dos alvos nas datas de referência e incrementos entre elas."""
    if not datas_ref or not sel:
        return
    linhas = []
    for a in sel:
        s = sub[sub[COLS["alvo"]] == a].sort_values(COLS["data"])
        lin = {"Alvo": a + rotulos.get(a, ("", None))[0]}
        ant = None
        for d in datas_ref:
            v, _t = _valor_na_data(s, col, d)
            lin[d.strftime("%d/%m/%Y")] = round(v, 1) if v is not None else None
            if ant is not None:
                d0, v0 = ant
                lin[f"Δ {d0.strftime('%d/%m/%y')}→{d.strftime('%d/%m/%y')}"] = (
                    round(v - v0, 1) if (v is not None and v0 is not None) else None)
            ant = (d, v)
        linhas.append(lin)
    st.caption("Valores nas datas de referência (mm; última leitura até cada data, "
               "no máximo 10 dias antes) e incrementos entre datas consecutivas:")
    tab = pd.DataFrame(linhas)
    st.dataframe(tab, use_container_width=True, hide_index=True,
                 column_config={c: st.column_config.NumberColumn(format="%.1f")
                                for c in tab.columns if c != "Alvo"})



def separador_alvos_2d(dados):
    alvos = dados["alvos"]
    if not validar_colunas(alvos, [COLS["data"], COLS["alvo"], COLS["edificio"],
                                   COLS["desl_h"], COLS["dZ"]], "Alvos topográficos"):
        return
    st.subheader("Alvos topográficos — evolução temporal")

    # recalcular estado de ΔH/ΔV com os criterios oficiais (auditoria)
    alvos = anexar_estado_calculado(alvos)

    # ---- painel de estado da ultima campanha ----------------------------
    ult = alvos[alvos[COLS["data"]] == alvos[COLS["data"]].max()].copy()
    data_ult = pd.to_datetime(alvos[COLS["data"]].max()).strftime("%d/%m/%Y")
    n_alarme = int((ult["Estado calculado"] == "Alarme").sum())
    n_alerta = int((ult["Estado calculado"] == "Alerta").sum())
    n_reg = int((ult["Estado calculado"] == "Regular").sum())
    n_sl = int((ult["Estado calculado"] == "Sem leitura").sum())

    st.markdown(f"**Estado na última campanha ({data_ult})** — recalculado dos "
                f"deslocamentos ΔH/ΔV com os critérios oficiais:")
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
            st.success(f"Auditoria: o estado recalculado pela app coincide com a "
                       f"coluna 'Estado' do Excel (calculada por fórmula) em "
                       f"{n_conf}/{n_tot} leituras (100%). As duas implementações "
                       f"dos critérios são coerentes; a comparação com o estado "
                       f"publicado nos relatórios 33GRADOS é feita à parte.")
        else:
            st.warning(f"Auditoria: divergência em {n_tot - n_conf}/{n_tot} leituras "
                       f"entre o estado recalculado e a coluna 'Estado' do Excel. "
                       f"Ver tabela de divergências abaixo.")
            with st.expander("Ver divergências estado recalculado vs. relatório"):
                div = comp[~comp["Confere"]]
                st.dataframe(
                    div[[COLS["data"], COLS["alvo"], COLS["edificio"],
                         COLS["desl_h"], COLS["dZ"], COLS["estado"],
                         "Estado calculado", "Critério"]],
                    use_container_width=True, hide_index=True)

    # alvos em alarme/alerta agora, para leitura rapida
    crit_now = ult[ult["Estado calculado"].isin(["Alarme", "Alerta"])].copy()
    if len(crit_now):
        crit_now = crit_now.sort_values("Estado calculado")
        st.caption("Alvos em alerta ou alarme na última campanha:")
        st.dataframe(
            crit_now[[COLS["alvo"], COLS["edificio"], "Fachada",
                      COLS["desl_h"], COLS["dZ"], "Estado calculado", "Critério"]]
            .rename(columns={COLS["desl_h"]: "Desl. H (mm)", COLS["dZ"]: "ΔZ (mm)"}),
            use_container_width=True, hide_index=True)
    st.divider()

    edificios = sorted(alvos[COLS["edificio"]].dropna().unique())
    edi = st.selectbox("Edifício / elemento", edificios)
    sub = alvos[alvos[COLS["edificio"]] == edi]

    # se for o Edifício 1, permitir filtrar por fachada e avisar da substituicao
    e_edificio1 = isinstance(edi, str) and "Edifício 1" in edi
    if e_edificio1:
        st.info(
            "O Edifício 1 tem duas fachadas instrumentadas: **Frente a escavação** "
            "(A1–A4) e **Lateral, virada ao mar** (A5–A8). Os alvos A5–A8 foram "
            "tapados por um painel publicitário e substituídos por **A5b–A8b**, "
            "RE-ZERADOS em 20/10/2025 — por isso os acumulados dos 'b' não são "
            "comparáveis diretamente com A1–A4 (arrancam de zero mais tarde).")
        fach = st.radio("Fachada", ["Ambas", "Frente escavação", "Lateral (mar)"],
                        horizontal=True)
        if fach != "Ambas":
            sub = sub[sub["Fachada"] == fach]

    lista = sorted(sub[COLS["alvo"]].dropna().unique(), key=_chave_natural)
    sel = st.multiselect("Alvos", lista, default=lista[:min(5, len(lista))])

    # localizacao fisica dos alvos (foto real ou esquema das coordenadas)
    if st.checkbox("Mostrar localização dos alvos no edifício", value=False,
                   help="Foto real do relatório, se disponível; caso contrário "
                        "um alçado esquemático a partir das coordenadas."):
        COR_ESTADO = {"Alarme": "#c0140f", "Alerta": "#e67e00",
                      "Regular": "#1f9e55"}
        mostrar_localizacao_alvo(sub, edi, sel, COR_ESTADO)

    # criterio aplicavel a este edificio (para desenhar as linhas de limiar)
    crit_edi, rotulo_edi, ac_edi = criterios_do_alvo(edi)
    Ha, Hm, Va, Vm = crit_edi
    st.caption(f"Critério aplicado: **{rotulo_edi}**. As linhas tracejadas nos "
               f"gráficos marcam os limiares de alerta e alarme.")
    if ac_edi:
        st.warning("Este alçado não consta da lista de critérios dos relatórios "
                   "de alvos; o critério de 17 m é assumido.")
    # rotulos de posicao (sup./inf.) e cores por par, a partir de TODOS os
    # alvos do grupo (o rotulo nao depende dos alvos escolhidos)
    rotulos = rotulos_posicao_alvos(alvos[alvos[COLS["edificio"]] == edi])

    # --- controlo de zoom temporal (util quando as fases da obra estao ligadas)
    cz1, cz2 = st.columns([1, 2])
    with cz1:
        granul = st.selectbox("Detalhe do eixo temporal",
                              ["Automático", "Mensal", "Quinzenal", "Semanal"],
                              index=0,
                              help="Marcas mais finas ajudam a ler as "
                                   "campanhas (~8 em 8 dias).")
    with cz2:
        # janela de datas para focar um periodo (ex. onde os deslocamentos disparam)
        d_min = pd.to_datetime(sub[COLS["data"]].min()).date()
        d_max = pd.to_datetime(sub[COLS["data"]].max()).date()
        janela = st.slider("Janela temporal", min_value=d_min, max_value=d_max,
                           value=(d_min, d_max), format="DD/MM/YY")
    j0 = pd.to_datetime(janela[0])
    j1 = pd.to_datetime(janela[1])
    sub = sub[(sub[COLS["data"]] >= j0) & (sub[COLS["data"]] <= j1)]

    # --- sobreposicoes: faseamento, maximo, datas de referencia, comparacao
    _datas_disp = [pd.to_datetime(d).strftime("%d/%m/%Y")
                   for d in sorted(sub[COLS["data"]].dropna().unique())]
    co1, co2, co3 = st.columns([1.2, 1, 1])
    with co1:
        fases = st.radio(
            "Faseamento no gráfico",
            ["Nenhum", "Períodos de escavação", "Fases da obra (plano)"],
            index=1, horizontal=True, key="alvos_fases",
            help="Períodos de escavação: seis períodos delimitados pelo início "
                 "da escavação de cada nível e pelo fim da escavação, segundo o "
                 "plano de trabalhos impactado (faixas alternadas). Fases da "
                 "obra: barra de faseamento do plano no topo do gráfico. São "
                 "uma referência comum: as frentes avançaram desfasadas.")
        marcar_max = st.checkbox(
            "Assinalar o máximo", value=True, key="alvos_max",
            help="Assinala o maior deslocamento horizontal e o maior |ΔZ| dos "
                 "alvos escolhidos, no período visível.")
        _prev_edi = previsto_do_grupo(edi)
        mostrar_proj = False
        if _prev_edi[0] is not None and _prev_edi[2].startswith("direta"):
            mostrar_proj = st.checkbox(
                "Estimativa do projeto", value=False, key=f"alvos_proj_{edi}",
                help="Desenha no gráfico horizontal a estimativa da memória "
                     "descritiva para a última fase de escavação na zona deste "
                     "alçado (20 mm a poente e norte; 10 mm a nascente e sul; "
                     "atribuição inferida) e a data em que o máximo dos alvos "
                     "escolhidos a ultrapassou.")
    with co2:
        _dref = st.multiselect(
            "Datas de referência", _datas_disp, key=f"alvos_dref_{edi}",
            placeholder="Escolher campanhas",
            help="Marca as campanhas escolhidas (linha vertical e valor de "
                 "cada alvo) e lista, por baixo de cada gráfico, os valores "
                 "nessas datas e os incrementos entre elas.")
        datas_ref = sorted(pd.to_datetime(x, format="%d/%m/%Y") for x in _dref)
    with co3:
        _outros = [e for e in edificios if e != edi]
        comp = st.multiselect(
            "Comparar com (intervalo mín.–máx.)",
            _outros, key=f"alvos_comp_{edi}", placeholder="Escolher grupos",
            help="Sobrepõe, em faixa, o intervalo entre o mínimo e o máximo "
                 "dos alvos de outro grupo em cada campanha. Só entram os "
                 "alvos com a zeragem inicial do grupo; os de substituição, "
                 "zerados mais tarde, ficam de fora.")
        # ordem estável das faixas: pela designação do primeiro alvo do grupo
        comp = sorted(comp, key=lambda g: _chave_natural(min(
            alvos.loc[alvos[COLS["edificio"]] == g, COLS["alvo"]].astype(str),
            key=_chave_natural)))

    # re-zeragem: alvos de substituicao escolhidos, zerados depois do grupo
    _ini_grupo = alvos[(alvos[COLS["edificio"]] == edi)
                       & alvos[COLS["desl_h"]].notna()][COLS["data"]].min()
    _rz = {}
    for a in sel:
        sa = sub[(sub[COLS["alvo"]] == a) & sub[COLS["desl_h"]].notna()]
        if sa.empty or not str(a).endswith("b"):
            continue
        t_a = pd.to_datetime(sa[COLS["data"]].min())
        if t_a > _ini_grupo + pd.Timedelta(days=3) and t_a >= j0:
            _rz.setdefault(t_a, []).append(a)
    rezer = [(t, sorted(n, key=_chave_natural)) for t, n in sorted(_rz.items())]

    def _notas_grafico(fases_vis, periodos_vis):
        legenda_fases(fases_vis)
        legenda_periodos(periodos_vis)
        if any(rotulos.get(a, ("",))[0] for a in sel):
            st.caption("Os pares de alvos à mesma posição em planta (a menos de "
                       "1 m) e a cotas diferentes partilham a cor: tom escuro no "
                       "superior (sup.) e claro no inferior (inf.).")
        if comp:
            st.caption("Faixas: intervalo entre o mínimo e o máximo, em cada "
                       "campanha, dos alvos de cada grupo comparado (só os alvos "
                       "com a zeragem inicial do grupo).")

    tab_h, tab_v = st.tabs(
        ["↔ Deslocamento horizontal (H)", "↕ Assentamento vertical (ΔZ)"])
    with tab_h:
        st.markdown(
            "<h4 style='text-align:center; margin-bottom:0; color:#1f2a44;'>"
            "Deslocamento horizontal acumulado (mm)</h4>",
            unsafe_allow_html=True)
        fig, _fv, _pv = grafico_alvos_tempo(
            sub, alvos, sel, COLS["desl_h"], rotulos,
            titulo_y="Desl. horizontal (mm)",
            limiares=((Ha, f"Alerta {Ha}"), (Hm, f"Alarme {Hm}")),
            fases=fases, datas_ref=datas_ref, comp=comp, marcar_max=marcar_max,
            modo_max="max", granul=granul, j0=j0, j1=j1, rezer=rezer,
            altura=780, projeto=(_prev_edi[0] if mostrar_proj else None))
        st.plotly_chart(fig, use_container_width=True)
        _notas_grafico(_fv, _pv)
        tabela_datas_referencia(sub, sel, datas_ref, COLS["desl_h"], rotulos)
    with tab_v:
        st.markdown(
            "<h4 style='text-align:center; margin-bottom:0; color:#1f2a44;'>"
            "Assentamento vertical acumulado, ΔZ (mm)</h4>",
            unsafe_allow_html=True)
        # limiares verticais: o assentamento e negativo -> desenhar em -Va e -Vm
        fig2, _fv, _pv = grafico_alvos_tempo(
            sub, alvos, sel, COLS["dZ"], rotulos,
            titulo_y="ΔZ (mm)",
            limiares=((-Va, f"Alerta -{Va}"), (-Vm, f"Alarme -{Vm}")),
            fases=fases, datas_ref=datas_ref, comp=comp, marcar_max=marcar_max,
            modo_max="abs", granul=granul, j0=j0, j1=j1, rezer=rezer,
            altura=780)
        st.plotly_chart(fig2, use_container_width=True)
        _notas_grafico(_fv, _pv)
        tabela_datas_referencia(sub, sel, datas_ref, COLS["dZ"], rotulos)

    # ---- direção (roseta) e razão ΔZ/δh na última campanha --------------
    st.divider()
    st.markdown("#### Direção e proporção do movimento (última campanha)")
    _ultv = []
    for a in sel:
        sa = sub[sub[COLS["alvo"]] == a].sort_values(COLS["data"])
        sa = sa[sa[COLS["desl_h"]].notna()]
        if sa.empty:
            continue
        r = sa.iloc[-1]
        _ultv.append((a, float(r[COLS["desl_h"]]), r.get(COLS["dZ"]),
                      r.get(COLS["dM"]), r.get(COLS["dP"]),
                      pd.to_datetime(r[COLS["data"]])))
    if _ultv:
        colr, colt = st.columns([1, 1])
        with colr:
            figp = go.Figure()
            _tem_dir = False
            for a, dh, dz, dm, dp, dt in _ultv:
                if pd.notna(dm) and pd.notna(dp) and dh > 0:
                    az = float(np.degrees(np.arctan2(dm, dp)) % 360)
                    _suf, _cor = rotulos.get(a, ("", None))
                    figp.add_trace(go.Scatterpolar(
                        r=[0, dh], theta=[az, az], mode="lines+markers",
                        name=a + _suf,
                        line=dict(color=_cor) if _cor else None,
                        marker=dict(color=_cor) if _cor else None))
                    _tem_dir = True
            if _tem_dir:
                figp.update_layout(
                    height=420, margin=dict(t=30, b=20),
                    polar=dict(radialaxis=dict(title="δh (mm)"),
                               angularaxis=dict(rotation=90, direction="clockwise")),
                    legend=dict(font=dict(size=10)))
                st.plotly_chart(figp, use_container_width=True)
                st.caption("Vetores do deslocamento horizontal (componentes ΔM/ΔP) "
                           "na última campanha. 0° = Norte (+P), 90° = Este (+M). "
                           "O desvio do rumo face à direção do alçado é leitura do "
                           "autor (a tese obtém ≤ 26°).")
            else:
                st.info("Sem componentes ΔM/ΔP para desenhar a roseta destes alvos.")
        with colt:
            st.markdown("**Razão assentamento/horizontal (|ΔZ|/δh)**")
            _rows = []
            for a, dh, dz, dm, dp, dt in _ultv:
                razao = (abs(float(dz)) / dh) if (pd.notna(dz) and dh > 0) else np.nan
                _rows.append({"Alvo": a + rotulos.get(a, ("", None))[0],
                              "δh (mm)": round(dh, 1),
                              "ΔZ (mm)": (round(float(dz), 1) if pd.notna(dz) else None),
                              "|ΔZ|/δh": (round(razao, 2) if pd.notna(razao) else None)})
            st.dataframe(pd.DataFrame(_rows), use_container_width=True,
                         hide_index=True)
            st.caption("No Edifício 1, a tese obtém ≈ 0,2 em A1/A2 e 0,4–0,5 nos "
                       "restantes alvos.")

    fis = dados.get("fissuro")
    if e_edificio1 and fis is not None and not fis.empty and "Fissurómetro" in fis.columns:
        st.markdown("#### Fissurómetros do Edifício 1 (F1–F5)")
        f_ = fis.copy(); f_["Data"] = pd.to_datetime(f_["Data"], errors="coerce")
        fig3 = go.Figure()
        for nome, g in f_.groupby("Fissurómetro"):
            loc = str(g["Localização"].iloc[0]).split(" - ")[-1] if "Localização" in g.columns else ""
            fig3.add_trace(go.Scatter(x=g["Data"], y=g["Leitura acumulada (mm)"], mode="lines+markers",
                                      name=f"{nome} — {loc}", line=dict(shape="hv")))
        for d, z, o, e in MEDIDAS_OBRA:
            if d and pd.to_datetime(d) >= f_["Data"].min() - pd.Timedelta(days=7):
                fig3.add_vline(x=pd.to_datetime(d), line=dict(color="#7a1fa0", width=1, dash="dot"))
        fig3.update_layout(height=420, margin=dict(r=60, t=30, b=40),
                           yaxis_title="Leitura acumulada (mm)", xaxis_title="Data",
                           legend=dict(orientation="h", yanchor="top", y=-0.2))
        st.plotly_chart(fig3, use_container_width=True)
        st.caption("Zeragem a 19/01/2026. A grandeza medida e a convenção de sinal não constam do quadro "
                   "de leituras (a confirmar); variações de 0,5 mm estão no limite da resolução. Linhas "
                   "roxas: medidas com data indicadas pela obra (separador «Obra»).")


# =========================================================================
# SEPARADOR 4 — CELULAS DE CARGA
# =========================================================================
def _mostrar_imagens_celula(cel_sel, key_suffix=""):
    """Mostra, num expander, as imagens de localização da célula (peças do
    projeto). Reutilizada no separador Células (Dados) e na Análise."""
    imgs = IMAGENS_CELULAS.get(cel_sel, [])
    if not imgs:
        return
    with st.expander(f"📍 Ver localização da célula {cel_sel} (peças do projeto)"):
        base_dir = Path(__file__).resolve().parent
        mostrou = False
        for fich, legenda in imgs:
            fp = base_dir / "fotos_celulas" / fich
            if not fp.exists():
                fp = Path("fotos_celulas") / fich
            if fp.exists():
                st.image(str(fp), caption=legenda, use_container_width=True)
                mostrou = True
        if not mostrou:
            st.info("Imagens de localização não encontradas (pasta "
                    "'fotos_celulas/'). Verifica que foram publicadas com a app.")
        loc = LOCALIZACAO_CELULAS.get(cel_sel)
        if loc:
            st.caption(f"Localização: {loc[0]}. Imagens das peças desenhadas do "
                       f"projeto de contenção (JETsj).")


def separador_celulas(dados):
    cc = dados["celulas"]
    if not validar_colunas(cc, [COLS["data"], COLS["celula"], COLS["carga_atual"],
                                COLS["variacao"]], "Células de carga"):
        return
    cc = cc.sort_values(COLS["data"])
    st.subheader("Células de carga")
    cel = st.selectbox("Célula", sorted(cc[COLS["celula"]].dropna().unique()))
    _mostrar_imagens_celula(cel, key_suffix="_inputs")
    sub = cc[cc[COLS["celula"]] == cel].sort_values(COLS["data"]).copy()
    ref = st.radio("Referência da variação de carga",
                   ["Blocagem indicada nos relatórios", "Primeira leitura da monitorização"],
                   horizontal=True, key="cc_ref",
                   help="Os relatórios calculam a variação sobre a «blocagem». Na A26, esse valor "
                        "(833,5 kN) é a leitura da célula no patamar de 1,0P do ensaio, e não a "
                        "carga de cravação: a primeira leitura da monitorização é 957,4 kN.")
    if ref.startswith("Primeira"):
        c0 = float(sub[COLS["carga_atual"]].iloc[0])
        sub[COLS["variacao"]] = (sub[COLS["carga_atual"]] - c0) / c0
        sub[COLS["estado"]] = np.where(sub[COLS["variacao"]] > CC_ALARME, "Alarme",
                                       np.where(sub[COLS["variacao"]] >= CC_ALERTA, "Alerta", "Regular"))
        st.caption(f"Variação calculada em relação à primeira leitura ({c0:.1f} kN, "
                   f"{sub[COLS['data']].iloc[0]:%d/%m/%Y}).")
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
                                  mode="lines+markers", name="Variação (%)"))
        fig2.add_hline(y=CC_ALERTA * 100, line_dash="dash", line_color="orange",
                       annotation_text="Alerta 15%")
        fig2.add_hline(y=CC_ALARME * 100, line_dash="dash", line_color="red",
                       annotation_text="Alarme 25%")
        fig2.update_xaxes(title="Data"); fig2.update_yaxes(title="Variação (%)")
        fig2.update_layout(height=460)
        st.plotly_chart(fig2, use_container_width=True)
    cols = [c for c in [COLS["data"], COLS["ancoragem"], COLS["carga_atual"],
                        COLS["blocagem"], COLS["variacao"], COLS["estado"]]
            if c in sub.columns]
    st.dataframe(sub[cols], use_container_width=True)

    # ---- carga × temperatura (coluna no modelo; hoje sem dados) ---------
    _tcol = "Temperatura (°C)"
    st.markdown("##### Efeito térmico")
    if _tcol in sub.columns and sub[_tcol].notna().sum() >= 3:
        _stc = sub[[COLS["carga_atual"], _tcol]].dropna()
        figt = go.Figure(go.Scatter(
            x=_stc[_tcol], y=_stc[COLS["carga_atual"]], mode="markers",
            marker=dict(color="#c0140f", size=8), name="Leituras"))
        if _stc[_tcol].nunique() >= 2:
            _b1, _b0 = np.polyfit(_stc[_tcol], _stc[COLS["carga_atual"]], 1)
            _xs = np.linspace(_stc[_tcol].min(), _stc[_tcol].max(), 20)
            figt.add_trace(go.Scatter(
                x=_xs, y=_b1 * _xs + _b0, mode="lines",
                line=dict(dash="dash", color="#1f3a5f"),
                name=f"Deriva ≈ {_b1:.2f} kN/°C"))
        figt.update_layout(height=380, xaxis_title="Temperatura (°C)",
                           yaxis_title="Carga (kN)", margin=dict(t=30, b=40),
                           legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(figt, use_container_width=True)
        st.caption("Carga da célula em função da temperatura; a deriva ajustada "
                   "compara-se com a do certificado de calibração.")
    else:
        st.info(
            "A coluna **Temperatura (°C)** existe no modelo mas está vazia, por "
            "isso a app não consegue separar o efeito térmico do estrutural. É a "
            "dúvida da Secção 5.2.8: as duas células subiram em simultâneo entre "
            "as leituras de final de junho e a de 06/07/2026 — a A26 +27,6 kN e a "
            "da ancoragem DE +24,2 kN. Com a deriva do sensor da A26 (0,75 kN/°C, "
            "certificado de calibração), seria preciso ≈ 37 °C de variação para "
            "explicar o salto só por temperatura. Este gráfico (carga × "
            "temperatura) aparece automaticamente quando carregares as "
            "temperaturas das células.")


# =========================================================================
# SEPARADOR 5 — PIEZOMETROS
# =========================================================================
def separador_piezometros(dados):
    pz = dados["piezo"]
    if not validar_colunas(pz, [COLS["data"], COLS["piezometro"], COLS["cota_agua"]],
                           "Piezometros"):
        return
    pz = pz.sort_values(COLS["data"])
    st.subheader("Piezómetros — cota da água vs. escavação")
    st.caption("O eixo vertical é a COTA (m), partilhada com as cotas de "
               "projeto da escavação/contenção e com o nível freático de "
               "repouso. Assim ve-se a que profundidade anda a água face a "
               "cada piso e ao fundo de escavação.")

    c1, c2 = st.columns([1, 2])
    with c1:
        p = st.selectbox("Piezómetro",
                         sorted(pz[COLS["piezometro"]].dropna().unique()))
    with c2:
        mostrar_pisos = st.checkbox("Cotas dos pisos e fundo de escavação",
                                    value=True)
        mostrar_nf = st.checkbox("Nível freático de repouso (2022)", value=True)

    sub = pz[pz[COLS["piezometro"]] == p].sort_values(COLS["data"])
    fig = go.Figure()

    # cotas de projeto como linhas horizontais (referencia geometrica)
    if mostrar_pisos:
        for nome, cota in COTAS_PISOS:
            fig.add_hline(y=cota, line=dict(color="rgba(120,120,120,0.5)",
                          width=1, dash="dot"),
                          annotation_text=f"{nome} ({cota:.2f})",
                          annotation_position="right",
                          annotation_font_size=9)
        fig.add_hline(y=COTA_FUNDO_ESCAVACAO,
                      line=dict(color="#b45309", width=2),
                      annotation_text=f"Fundo de escavação ({COTA_FUNDO_ESCAVACAO:.2f})",
                      annotation_position="right", annotation_font_size=10)

    # nivel freatico de repouso (faixa entre min e max das sondagens)
    if mostrar_nf:
        nfs = [c for _, c in NF_REPOUSO]
        fig.add_hrect(y0=min(nfs), y1=max(nfs),
                      fillcolor="rgba(37,99,235,0.10)", line_width=0,
                      annotation_text="NF de repouso (2022)",
                      annotation_position="top left", annotation_font_size=9)

    # a serie do piezometro por cima
    fig.add_trace(go.Scatter(x=sub[COLS["data"]], y=sub[COLS["cota_agua"]],
                             mode="lines+markers", name="Cota da água (PZ)",
                             line=dict(color="#2563eb", width=2)))
    fig.update_xaxes(title="Data")
    fig.update_yaxes(title="Cota (m)")
    fig.update_layout(
        height=560,
        margin=dict(r=180, t=40, b=40),
        legend=dict(orientation="v", yanchor="top", y=1,
                    xanchor="left", x=1.01, font=dict(size=10)))
    st.plotly_chart(fig, use_container_width=True)
    # nota: o faseamento aparece na barra no topo do grafico e na legenda a
    # direita; a caption redundante foi removida.

    # leitura cruzada quantitativa
    if len(sub):
        c_ini = sub[COLS["cota_agua"]].iloc[0]
        c_fim = sub[COLS["cota_agua"]].iloc[-1]
        desc = c_ini - c_fim
        nf_med = sum(c for _, c in NF_REPOUSO) / len(NF_REPOUSO)
        st.markdown(
            f"**Leitura:** a água no {p} desceu de **{c_ini:.2f}** para "
            f"**{c_fim:.2f} m** ({desc:+.2f} m) no período monitorizado. "
            f"O fundo de escavação (**{COTA_FUNDO_ESCAVACAO:.2f} m**) fica "
            f"{c_fim - COTA_FUNDO_ESCAVACAO:.1f} m abaixo da água atual e "
            f"~{nf_med - COTA_FUNDO_ESCAVACAO:.0f} m abaixo do nível freático "
            f"de repouso de 2022 (~{nf_med:.0f} m).")
    st.caption("Cotas de projeto: escavação e contenção periférica (JETsj, "
               "PRO/2023/368). NF de repouso: piezómetros das sondagens "
               "(ENGGEO, Quadro III, 24/11/2022).")


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
    Transformação de semelhanca 2D (translação + rotação + escala uniforme)
    que leva pontos do sistema do DESENHO (src) para o sistema dos ALVOS (dst),
    a partir de 2 pares de pontos correspondentes.

    Dois pontos chegam para fixar as 4 incognitas (2 translação, 1 rotação,
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

    st.caption("Carrega uma planta em DXF (por exemplo a planta de escavação "
               "e contenção de um piso). Podes ve-la sozinha, ou ALINHA-LA com "
               "os alvos topográficos indicando 2 pontos de referência — a app "
               "calcula a transformação e sobrepõe os alvos coloridos pelo "
               "deslocamento.")

    dxf = st.file_uploader("Ficheiro DXF", type=["dxf"])
    if dxf is None:
        st.info("Carrega um DXF para ver a planta. Sugestão: a planta de "
                "escavação do piso -1, -2 ou -3 mostra bem o contorno do "
                "recinto de contenção.")
        return

    conteudo = dxf.getvalue()

    try:
        layers = dxf_listar_layers(conteudo)
    except Exception as e:
        st.error(f"Não consegui ler as layers do DXF. Detalhe: {e}")
        return

    sugeridas = [l for l in layers
                 if any(l.upper().startswith(p) for p in LAYERS_ESTRUTURAIS_SUGERIDAS)]

    st.write(f"O desenho tem **{len(layers)} layers**. Estão pre-selecionadas "
             f"as estruturais (contorno, estacas, grelha). Ajusta se quiseres:")
    layers_sel = st.multiselect("Layers a desenhar", sorted(layers),
                                default=sorted(sugeridas) if sugeridas else [])
    if not layers_sel:
        st.warning("Escolhe pelo menos uma layer para desenhar.")
        return

    try:
        segmentos = dxf_extrair_segmentos(conteudo, set(layers_sel))
    except Exception as e:
        st.error(f"Não consegui extrair a geometria. Detalhe: {e}")
        return
    if not segmentos:
        st.warning("Não encontrei geometria nas layers escolhidas. Tenta outras.")
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
            "Alinhar a planta com os alvos (sobreposição)", value=False,
            help="Precisas de indicar, para 2 alvos, onde eles estão no "
                 "desenho. A coordenada no sistema dos alvos já vêm do Excel.")

    if alinhar:
        # coordenadas dos alvos (sistema dos alvos) — uma linha por alvo (usar campanha mais recente)
        ult = alvos[alvos[COLS["data"]] == alvos[COLS["data"]].max()]
        lista_alvos = sorted(ult[COLS["alvo"]].astype(str).unique())

        st.markdown("**Ponto de referência 1**")
        c1, c2, c3 = st.columns(3)
        a1 = c1.selectbox("Alvo 1", lista_alvos, key="a1")
        x1d = c2.number_input("X no desenho", value=0.0, key="x1d", format="%.2f")
        y1d = c3.number_input("Y no desenho", value=0.0, key="y1d", format="%.2f")

        st.markdown("**Ponto de referência 2** (escolhe um bem afastado do 1)")
        d1, d2, d3 = st.columns(3)
        a2 = d1.selectbox("Alvo 2", lista_alvos,
                          index=min(len(lista_alvos) - 1, 1), key="a2")
        x2d = d2.number_input("X no desenho", value=0.0, key="x2d", format="%.2f")
        y2d = d3.number_input("Y no desenho", value=0.0, key="y2d", format="%.2f")

        st.caption("Como obter o X,Y no desenho: abre o DXF no AutoCAD ou num "
                   "visualizador, aponta o cursor ao sítio onde o alvo esta, e "
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
                m2.metric("Rotação", f"{ang:.1f}°")
                m3.metric("Estado", "Alinhado")
                if not (0.5 < escala < 2.0):
                    st.warning("A escala calculada e invulgar. Confirma as "
                               "coordenadas dos pontos no desenho — pode haver "
                               "troca de X/Y ou de ponto.")
            except Exception as e:
                st.error(f"Não consegui calcular o alinhamento: {e}")

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
                   "contorno bate certo com a nuvem de alvos; se não, ajusta os "
                   "pontos de referência.")
    else:
        st.caption("Planta em coordenadas locais do desenho. Ativa o "
                   "alinhamento acima para a sobrepor aos alvos.")


# =========================================================================
# SEPARADOR 7 — GEOLOGIA
# =========================================================================
def desenhar_coluna_litologica(fig, sondagem, x_centro=0, largura=0.8,
                               em_cota=False, cota_terreno=0):
    """Desenha a coluna litológica de uma sondagem como retangulos coloridos.
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
    Ponte de LEITURA N -> zona geotécnica, segundo os intervalos do próprio
    relatório (Quadros V-VII): ZG5 (SPT 11-30), ZG4 (31-56), ZG3/ZG2/ZG1 (>=60,
    distinguidas pelo RQD que NÃO temos por ponto). Serve só para colorir o
    ponto SPT e dar leitura rápida; NÃO define fronteiras de camada.
    Devolve a chave da zona (ou grupo) para indexar ZONA_CORES.
    """
    if n < 11:
        return "ZG6"
    if n <= 30:
        return "ZG5"
    if n <= 56:
        return "ZG4"
    return "ZG3-ZG1"          # nega: SPT nao separa ZG3/ZG2/ZG1 (so o RQD separa)


# =========================================================================
# PALETA DE ZONAMENTO GEOTECNICO — alinhada ao relatorio ENGGEO (Quadros
# V-VII). Gradiente que comunica a CONSOLIDACAO CRESCENTE do gres: laranja
# (aterro) -> verdes progressivamente mais escuros ate quase preto (ZG1).
# A luminancia desce monotonicamente do ZG5a ao ZG1 (validado).
# ZONA_CORES_FULL: as 6 zonas + pontuais, para a tabela/legenda.
# ZONA_CORES: as chaves que a classificacao por N produz (grupos), para
# colorir os pontos SPT — a nega fica numa cor unica porque o SPT nao a
# separa.
# =========================================================================
ZONA_CORES_FULL = {
    "ZG6":  "#c0641e",   # aterro
    "ZG5a": "#e8efe0",   # zona pontual (verde quase branco)
    "ZG5":  "#d3e2c4",   # gres N 11-30
    "ZG4":  "#a9c47f",   # gres N 31-56
    "ZG3a": "#7fa860",   # zona pontual
    "ZG3":  "#5c8a45",   # nega RQD 0-25%
    "ZG2":  "#3f6f3f",   # nega RQD 45-75%
    "ZG1":  "#26401f",   # nega RQD 76-100% (quase preto)
}

# cor de cada grupo produzido por _zona_por_N (a nega usa um verde escuro
# intermedio, representando o conjunto ZG3-ZG1 que o SPT nao distingue)
ZONA_CORES = {
    "ZG6": ZONA_CORES_FULL["ZG6"],
    "ZG5": ZONA_CORES_FULL["ZG5"],
    "ZG4": ZONA_CORES_FULL["ZG4"],
    "ZG3-ZG1": "#3a5f34",   # nega (conjunto), verde escuro
}

# litologia alinhada a mesma familia de cores, para coerencia visual entre
# o perfil litologico e o zonamento (mesmo verde-base para o gres)
GEO_CORES_LITO_V2 = {
    "Aterro": "#c0641e",
    "Grés (C1As)": "#a9c47f",
    "Calcário (C1A)": "#5c8a45",
}


def separador_geologia(dados):
    st.subheader("Geologia do terreno (Relatório ENGGEO, proc. 220216)")
    st.caption("Leitura integrada por sondagem: a coluna litológica, os ensaios "
               "SPT e o zonamento geotécnico partilham o eixo de profundidade, "
               "para se lerem em conjunto. Sob ~0,5 m de aterro, o terreno e "
               "essencialmente grés dos 'Grés Superiores' (C1As), com calcário "
               "(C1A) apenas no fundo do SC8. A resistência não é uniforme: há "
               "troços de menor N(SPT) intercalados com nega (por exemplo, N 19-47 "
               "no SC8 entre as cotas ~24 e ~18 e N 22-38 no SC6 entre as cotas "
               "~14 e ~6,6) e o RQD é nulo em grande parte das sondagens. Compare "
               "sondagens por COTA, não por profundidade (as bocas estão a cotas "
               "diferentes).")

    sonds = list(GEO_LITOLOGIA.keys())

    col_sel, col_info = st.columns([1, 2])
    with col_sel:
        modo = st.radio("Vista", ["Uma sondagem (detalhe)", "As quatro (comparar)"])
        if modo == "Uma sondagem (detalhe)":
            sond = st.selectbox("Sondagem", sonds)
        else:
            sond = None
    with col_info:
        st.caption("A coluna colorida à esquerda de cada sondagem é a litologia; "
                   "os pontos e a linha são o SPT (N pancadas), no mesmo eixo de "
                   "profundidade. A cor do ponto SPT indica a zona geotécnica "
                   "provável pelo valor de N (ver tabela em baixo). A linha azul "
                   "tracejada é o nível freático.")

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
                             name="Nível freático"))
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
    st.caption("Como ler: para cada sondagem, a barra colorida à esquerda é a "
               "litologia; a curva à direita é o SPT (escala 0–60, com a linha "
               "pontilhada na nega). Onde o SPT sobe, o grés está mais "
               "consolidado — e e ai que a rigidez do maciço aumenta. "
               "Profundidades de cada zona ZG NÃO estão definidas no relatório "
               "por ponto; a cor do SPT é a zona PROVÁVEL pelo valor de N.")

    st.divider()

    st.markdown("#### Zonamento geotécnico e parâmetros de projeto")
    st.caption("Parâmetros propostos (Quadros V-VII do relatório) que alimentam "
               "a modelação numérica da contenção. A coluna 'SPT típico' é a "
               "ponte para o perfil acima: é por ela que se lê em que zona está "
               "cada troço de terreno.")

    faixa_spt = {
        "ZG6": "aterro", "ZG5": "11–30", "ZG4": "31–56",
        "ZG3": "≥60 (RQD 0–25%)", "ZG2": "≥60 (RQD 45–75%)",
        "ZG1": "≥60 (RQD 76–100%)",
    }
    zt = pd.DataFrame(GEO_ZONAMENTO)
    zt.insert(2, "SPT típico (N)", zt["Zona"].map(faixa_spt))

    # colorir a celula da zona com a cor oficial (gradiente de consolidacao)
    def _estilo_zona(v):
        cor = ZONA_CORES_FULL.get(v, "")
        if not cor:
            return ""
        # texto claro sobre fundos escuros
        r = int(cor[1:3], 16); g = int(cor[3:5], 16); b = int(cor[5:7], 16)
        lum = 0.299*r + 0.587*g + 0.114*b
        txt = "#ffffff" if lum < 140 else "#1a1a1a"
        return f"background-color: {cor}; color: {txt}; font-weight: 600;"

    st.dataframe(zt.style.map(_estilo_zona, subset=["Zona"]),
                 use_container_width=True, hide_index=True)
    st.caption("gama: peso volúmico | c': coesão | fi': ângulo de atrito | "
               "E': módulo de deformabilidade. Zonas ZG3-ZG1 (rocha) com c' e E' "
               "em MPa/GPa; ZG6-ZG4 (solo/grés brando) em kPa/MPa. Nota: as três "
               "zonas de nega (ZG3-ZG1) distinguem-se pelo RQD, que o SPT sozinho "
               "não mede — por isso o perfil agrupa-as como 'nega'.")



# =========================================================================
# SEPARADOR 8 — CRONOGRAMA DA OBRA
# =========================================================================
def separador_obra(dados):
    st.subheader("Cronograma da obra — plano contratual e plano impactado")
    st.caption("O plano IMPACTADO (22/04/2026; data de estado 21/08/2026) é o plano "
               "CONTRATUAL com três suspensões de frentes inseridas durante a "
               "execução das estacas, recalculado com as durações contratuais. As "
               "datas são as do plano — não um registo independente da execução. "
               "A janela de instrumentação está assinalada.")

    vista = st.radio("Cronograma a mostrar",
                     ["Plano impactado", "Comparação impactado vs. contratual"],
                     horizontal=True)

    fig = go.Figure()
    # impactos (suspensoes) — so existem no plano impactado
    for nome, ini, fim, du, oque, prop in IMPACTOS_PLANO:
        fig.add_trace(go.Scatter(
            x=[pd.to_datetime(ini), pd.to_datetime(fim)], y=[nome, nome],
            mode="lines", line=dict(color="#c0140f", width=12),
            hovertemplate=f"{nome}<br>{ini} a {fim} ({du} dias úteis)<br>{oque}"
                          f"<br>Propagação: {prop}<extra></extra>",
            showlegend=False))
    if vista == "Plano impactado":
        for i, (nome, ci, cf, ii, fi) in enumerate(FASES_COMPARACAO):
            cor = CORES_FASES[i % len(CORES_FASES)]
            fig.add_trace(go.Scatter(
                x=[pd.to_datetime(ii), pd.to_datetime(fi)], y=[nome, nome],
                mode="lines", line=dict(color=cor, width=16),
                hovertemplate=f"{nome}<br>impactado: {ii} a {fi}<br>contratual: "
                              f"{ci} a {cf}<extra></extra>",
                showlegend=False))
    else:
        for i, (nome, ci, cf, ii, fi) in enumerate(FASES_COMPARACAO):
            cor = CORES_FASES[i % len(CORES_FASES)]
            fig.add_trace(go.Scatter(
                x=[pd.to_datetime(ii), pd.to_datetime(fi)], y=[f"{nome} ", f"{nome} "],
                mode="lines", line=dict(color="#1f6fb4", width=11),
                hovertemplate=f"IMPACTADO: {ii} a {fi}<extra></extra>", showlegend=False))
            fig.add_trace(go.Scatter(
                x=[pd.to_datetime(ci), pd.to_datetime(cf)], y=[f" {nome}", f" {nome}"],
                mode="lines", line=dict(color="rgba(120,120,120,0.6)", width=11),
                hovertemplate=f"CONTRATUAL: {ci} a {cf}<extra></extra>", showlegend=False))
        st.caption("Em cada fase: barra azul = plano impactado; barra cinzenta = "
                   "plano contratual. Barras vermelhas = suspensões (impactos).")

    alvos = dados.get("alvos")
    if alvos is not None and not alvos.empty and COLS["data"] in alvos.columns:
        d0 = alvos[COLS["data"]].min()
        d1 = alvos[COLS["data"]].max()
        fig.add_vrect(x0=d0, x1=d1, fillcolor="crimson", opacity=0.10,
                      line_width=0, annotation_text="Instrumentação",
                      annotation_position="top left")
    fig.update_xaxes(title="Data")
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(height=560, margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig, use_container_width=True)

    # ---- impactos e desfasamento por frente ------------------------------
    st.markdown("#### Impactos e desfasamento por frente")
    st.table(pd.DataFrame([{
        "Impacto": n, "Suspensão": f"{pd.to_datetime(a):%d/%m} a {pd.to_datetime(b):%d/%m/%Y}",
        "Dias úteis": du, "O que ficou suspenso": oq, "Propagação no plano": pr}
        for n, a, b, du, oq, pr in IMPACTOS_PLANO]).set_index("Impacto"))
    st.table(pd.DataFrame([{
        "Frente": f, "Edifício em frente": e, "Origem do desfasamento": o,
        "Dias úteis": du, "Dias corridos": dc,
        "Fundo de escavação (contratual)": fc, "Fundo de escavação (impactado)": fi}
        for f, e, o, du, dc, fc, fi in DESFASAMENTO_FRENTES]).set_index("Frente"))
    st.markdown(
        "**Leitura:** nenhuma atividade tem, no plano impactado, duração diferente "
        "da contratual — as 180 tarefas elementares de escavação e travamento "
        "mantêm a duração e foram deslocadas. Todo o atraso vem das suspensões de "
        "frentes durante as estacas: o impacto 1 (caneiro) desloca a frente poente, "
        "sob o Edifício 1, em 36 dias úteis; o impacto 2 (cabos BT/MT) desloca a "
        "frente norte e a escavação da frente nascente/sul em 32. A escavação "
        "termina a **15/04/2026** em vez de 02/03/2026 e a conclusão da 1.ª fase "
        "passa de 11/01/2027 para 01/03/2027.")
    anc = dados.get("ancoragens")
    if anc is not None and not anc.empty and "Lote" in anc.columns:
        st.markdown("#### Ancoragens — plano impactado vs. registos de execução")
        a = anc.copy()
        for c in ("Furação e selagem (adotada)", "Tensionamento e ensaio (adotado)", "Fim da injeção"):
            a[c] = pd.to_datetime(a[c], errors="coerce")
        lin = []
        for lote, g in a.groupby("Lote", sort=False):
            f0, f1 = g["Furação e selagem (adotada)"].min(), g["Furação e selagem (adotada)"].max()
            t0, t1 = g["Tensionamento e ensaio (adotado)"].min(), g["Tensionamento e ensaio (adotado)"].max()
            dd = g["Dias entre o fim da injeção e o tensionamento"].dropna()
            fmt = lambda x, y: (f"{x:%d/%m/%Y}" if x == y else f"{x:%d/%m} a {y:%d/%m/%Y}")
            pl = ANC_PLANO.get(lote, ("—", "—"))
            lin.append({"Lote": lote, "Alçado e nível": g[_col_nivel(g)].iloc[0],
                        "Plano: execução": pl[0], "Plano: tensionamento": pl[1],
                        "Registo: furação": fmt(f0, f1), "Registo: tensionamento": fmt(t0, t1),
                        "Dias fim da injeção → tensionamento":
                            (f"{int(dd.min())} a {int(dd.max())}" if len(dd) else "—")})
        st.table(pd.DataFrame(lin).set_index("Lote"))
        st.caption("Registos do empreiteiro (furação e selagem, multi-injeção e ensaios). Alçado GH "
                   "segundo o alçado com a numeração das ancoragens; no troço poente do AB, o nível é "
                   "inferido e a correspondência com o plano (?) está por confirmar. O plano trata as "
                   "ancoragens dos dois lados de cada nível numa só tarefa. O procedimento de execução "
                   "prevê o tensionamento 7 a 8 dias após o fim das injeções. Datas incoerentes nos "
                   "registos foram estimadas (folha Qualidade_Dados).")
    # ---- CRONOLOGIA VISUAL DAS ANCORAGENS --------------------------------
    anc2 = dados.get("ancoragens")
    if anc2 is not None and not anc2.empty and "Ancoragem" in anc2.columns:
        import re as _re
        st.markdown("#### Cronologia de execução das ancoragens")
        st.caption(
            "Para cada ancoragem: furação/selagem (○), injeção (barra roxa, do "
            "início ao fim) e tensionamento/ensaio (◆). A cor do tensionamento "
            "indica o intervalo entre o FIM da injeção e o tensionamento — o "
            "procedimento prevê 7 a 8 dias. Datas adotadas (folha Qualidade_Dados "
            "corrige registos incoerentes).")

        b = anc2.copy()
        for _c in ("Furação e selagem (adotada)", "Início da injeção",
                   "Fim da injeção", "Tensionamento e ensaio (adotado)"):
            b[_c] = pd.to_datetime(b[_c], errors="coerce")
        b["_dias"] = pd.to_numeric(
            b["Dias entre o fim da injeção e o tensionamento"], errors="coerce")

        def _alc_anc(x):
            s = str(x)
            m = _re.match(r"\s*([A-Z]{2})\b", s)
            if m:
                return m.group(1)
            return "Viga da grua" if "grua" in s.lower() else "Outros"

        b["_alc"] = b["Alçado e nível"].apply(_alc_anc)
        _opts = ["Todos"] + sorted(b["_alc"].dropna().unique())
        _fa = st.selectbox("Alçado", _opts, key="anc_tl_alc")
        bb = b if _fa == "Todos" else b[b["_alc"] == _fa]
        bb = bb.sort_values("Início da injeção")

        if bb.empty:
            st.info("Sem ancoragens para o alçado escolhido.")
        else:
            ordem = bb["Ancoragem"].astype(str).tolist()
            figA = go.Figure()
            # barras de injeção (uma linha roxa por ancoragem, com None a separar)
            xi, yi = [], []
            for _, r in bb.iterrows():
                i0, i1 = r["Início da injeção"], r["Fim da injeção"]
                if pd.notna(i0) and pd.notna(i1):
                    xi += [i0, i1, None]
                    yi += [str(r["Ancoragem"]), str(r["Ancoragem"]), None]
            if xi:
                figA.add_trace(go.Scatter(
                    x=xi, y=yi, mode="lines",
                    line=dict(color="#7c3aed", width=7),
                    name="Injeção", hoverinfo="skip"))
            # furação (círculo)
            fm = bb[bb["Furação e selagem (adotada)"].notna()]
            figA.add_trace(go.Scatter(
                x=fm["Furação e selagem (adotada)"],
                y=fm["Ancoragem"].astype(str), mode="markers",
                marker=dict(symbol="circle-open", size=9, color="#1f2a44",
                            line=dict(width=2)),
                name="Furação/selagem",
                hovertemplate="%{y} · furação %{x|%d/%m/%Y}<extra></extra>"))
            # tensionamento (losango), cor por intervalo
            _buckets = [("≤ 8 dias", "#15803d", lambda d: d <= 8),
                        ("9–20 dias", "#ea8a00", lambda d: 8 < d <= 20),
                        ("> 20 dias", "#c0140f", lambda d: d > 20)]
            for _nm, _cor, _cond in _buckets:
                sub = bb[bb["_dias"].notna() & bb["_dias"].apply(
                    lambda d: _cond(d) if pd.notna(d) else False)]
                if not sub.empty:
                    figA.add_trace(go.Scatter(
                        x=sub["Tensionamento e ensaio (adotado)"],
                        y=sub["Ancoragem"].astype(str), mode="markers",
                        marker=dict(symbol="diamond", size=11, color=_cor,
                                    line=dict(width=1, color="white")),
                        name=f"Tensionamento ({_nm})",
                        customdata=sub["_dias"],
                        hovertemplate="%{y} · tensionamento %{x|%d/%m/%Y}"
                                      "<br>%{customdata:.0f} dias após injeção"
                                      "<extra></extra>"))
            semd = bb[bb["_dias"].isna() & bb["Tensionamento e ensaio (adotado)"].notna()]
            if not semd.empty:
                figA.add_trace(go.Scatter(
                    x=semd["Tensionamento e ensaio (adotado)"],
                    y=semd["Ancoragem"].astype(str), mode="markers",
                    marker=dict(symbol="diamond-open", size=10, color="#888",
                                line=dict(width=1)),
                    name="Tensionamento (sem registo do intervalo)",
                    hovertemplate="%{y} · tensionamento %{x|%d/%m/%Y}<extra></extra>"))
            figA.update_layout(
                height=140 + 20 * len(ordem),
                xaxis=dict(title="Data"),
                yaxis=dict(categoryarray=ordem, categoryorder="array",
                           autorange="reversed", title="Ancoragem"),
                margin=dict(l=10, r=10, t=30, b=40),
                legend=dict(orientation="h", yanchor="bottom", y=1.02))
            st.plotly_chart(figA, use_container_width=True)

            _d = bb["_dias"].dropna()
            if len(_d):
                _n8 = int((_d > 8).sum())
                _n20 = int((_d > 20).sum())
                st.caption(
                    f"Nos registos mostrados, **{_n8} de {len(_d)}** ancoragens "
                    f"ultrapassam a janela de 7–8 dias entre injeção e "
                    f"tensionamento e **{_n20}** ultrapassam 20 dias (até "
                    f"{int(_d.max())}). Os maiores intervalos concentram-se no "
                    f"alçado GH (viga de distribuição do piso −2) e no troço "
                    f"poente do AB (piso −1), onde o tensionamento foi diferido.")

    st.markdown("#### Ocorrências e medidas indicadas pela obra")
    med = pd.DataFrame([{"Data": (pd.to_datetime(d).strftime("%d/%m/%Y") if d else "—"), "Zona": z,
                         "Ocorrência ou medida": o, "Estado": e} for d, z, o, e in MEDIDAS_OBRA])
    st.table(med.set_index("Data"))
    st.caption("As medidas com data são assinaladas nos gráficos dos fissurómetros do Edifício 1 "
               "(separador «Alvos (2D)»).")
    st.caption("Dias úteis contados de segunda a sexta. O impacto 3 (Edifício 2) "
               "não está ligado a tarefas seguintes no plano, e as estacas "
               "E6_21–E6_27 aparecem executadas a partir de 07/08/2025, antes do "
               "fim dessa suspensão (21/08/2025) — incoerência a confirmar com o "
               "empreiteiro. Por manter as durações contratuais, o plano não mede o "
               "ritmo efetivo de cada atividade.")


# =========================================================================
# SEPARADOR 0 — PAGINA INICIAL (HOME)
# =========================================================================
def separador_pressupostos(dados):
    """
    Reúne num só sítio TODOS os pressupostos, inferências e limitações da
    ferramenta e dos dados — o que e medido vs. o que e assumido, e o estado de
    confirmação de cada um. Assumir as limitações explicitamente e uma forca:
    antecipa as perguntas difíceis da defesa.
    """
    st.subheader("Pressupostos, inferências e qualidade dos dados")
    st.caption("Transparência sobre o que é MEDIDO e o que é ASSUMIDO. Cada "
               "pressuposto tem a sua origem e o estado de confirmação. Esta "
               "secção reúne, num só sítio, as ressalvas assinaladas ao longo "
               "da aplicação.")

    st.markdown("#### Pressupostos e inferências da ferramenta")
    press = pd.DataFrame([
        {"Item": "Referencial de cota (3D)",
         "O que se assume": "A caixa de escavação usa a PROFUNDIDADE real "
                            "(16,3 m), não a cota absoluta — os sistemas de cota "
                            "dos alvos (Z 42-63) e do projeto (4,5-24,6) não "
                            "estão relacionados.",
         "Estado": "A confirmar (falta 1 par de cotas nos 2 sistemas)"},
        {"Item": "Associação inclinómetro-sondagem",
         "O que se assume": "I1<->SC6, I2<->SC9, I3<->SC6, inferido por "
                            "sobreposição das plantas de prospeção e de "
                            "instrumentação. A sobreposição gráfica é feita "
                            "por PROFUNDIDADE, não por cota (cota da boca dos "
                            "tubos desconhecida).",
         "Estado": "A confirmar com a equipa de instrumentação"},
        {"Item": "Critério de alerta e alarme de cada alçado",
         "O que se assume": "Valores da memória descritiva (17 m: 20/40 · 10/15 mm; 24 m: "
                            "30/40 · 10/15 mm), atribuídos a cada alçado segundo os relatórios "
                            "de alvos: 17 m em AB, CD, DE, BF e PQ; 24 m em FG, GH, JK, KL, LM, "
                            "MNO e OP (DE e LM a partir do R36, de 16/03/2026). MNO e OP têm "
                            "17 a 20 m de cortina, mas são tratados como 24 m.",
         "Estado": "Documentado nos relatórios; motivo da classificação de MNO e OP a confirmar"},
        {"Item": "Alçado e nível de cada ancoragem",
         "O que se assume": "Alçado GH (A19–A42): segundo o alçado com a numeração das ancoragens "
                            "(A26 no piso −1, cota 15,90). Troço poente do AB (A1–A18): nível "
                            "inferido da cronologia e do comprimento livre.",
         "Estado": "GH confirmado; AB a confirmar"},
        {"Item": "Fissurómetros do Edifício 1",
         "O que se assume": "Leituras acumuladas em mm desde 19/01/2026, tal como recebidas.",
         "Estado": "Grandeza medida e convenção de sinal a confirmar"},
        {"Item": "Referência das células de carga",
         "O que se assume": "Os relatórios calculam a variação sobre a «blocagem»; na A26 "
                            "esse valor (833,5 kN) é a leitura no patamar de ensaio. A app "
                            "permite usar a primeira leitura (957,4 kN).",
         "Estado": "Pressuposto declarado"},
        {"Item": "Faseamento da obra",
         "O que se assume": "Datas do plano de trabalhos impactado (o plano "
                            "contratual com três suspensões de frentes, mantendo as "
                            "durações contratuais); não há registo diário de obra. "
                            "As ancoragens usam as datas dos registos de execução "
                            "(folha «Ancoragens»), que diferem do plano em semanas.",
         "Estado": "Pressuposto declarado"},
        {"Item": "Zona geotécnica no perfil SPT",
         "O que se assume": "A cor da zona (ZG) é a provável pelo valor de N; o "
                            "relatório não define fronteiras de zona por "
                            "profundidade.",
         "Estado": "Leitura qualitativa (não fronteiras reais)"},
    ])
    st.dataframe(press, use_container_width=True, hide_index=True)

    st.markdown("#### Tratamento de dados (critérios aplicados)")
    trat = pd.DataFrame([
        {"Item": "Estado dos alvos",
         "Tratamento": "Recalculado de ΔH/ΔV com os critérios do relatório; "
                       "comparado com a coluna Estado do Excel (fórmula) — "
                       "coerência entre implementações, não confirmação do "
                       "estado publicado nos relatórios."},
        {"Item": "Leituras nulas de A5-A8 após 20/10/2025",
         "Tratamento": "Os alvos tapados tem leitura 0,0000 no relatório; a app "
                       "trata-as como ausência de leitura (a série não cai a "
                       "zero)."},
        {"Item": "Alvos A5-A8 -> A5b-A8b",
         "Tratamento": "Tratados como séries separadas; os 'b' foram re-zerados "
                       "em 20/10/2025 e os acumulados não são comparáveis "
                       "diretamente com A1-A4."},
        {"Item": "Fachadas do Edifício 1",
         "Tratamento": "Distinguidas frontal (A1-A4) e lateral (A5-A8) pela "
                       "posição em planta."},
    ])
    st.dataframe(trat, use_container_width=True, hide_index=True)

    # questoes de qualidade do proprio ficheiro (folha Qualidade_Dados)
    try:
        base_dir = Path(__file__).resolve().parent
        fpath = base_dir / FICHEIRO_EXCEL
        if not fpath.exists():
            fpath = FICHEIRO_EXCEL
        xq = pd.ExcelFile(fpath)
        if "Qualidade_Dados" in xq.sheet_names:
            qd = pd.read_excel(xq, "Qualidade_Dados")
            st.markdown("#### Questões de qualidade dos dados (do relatório)")
            st.caption("Ressalvas identificadas no próprio modelo de dados, "
                       "transcritas da folha Qualidade_Dados.")
            st.dataframe(qd, use_container_width=True, hide_index=True)
    except Exception:
        pass

    st.info("Nota metodológica: estas limitações não invalidam a análise — "
            "delimitam o seu alcance. A app apresenta medições e correlações; a "
            "interpretação das causas (escavação, furação das ancoragens, "
            "rebaixamento da água, rigidez do maciço) é feita pelo engenheiro e "
            "uma correlação temporal não demonstra, por si só, causalidade.")


@st.cache_data
def _carregar_terreno():
    """Le o modelo digital de terreno (MDT) extraido do levantamento topográfico."""
    import json
    base_dir = Path(__file__).resolve().parent
    fp = base_dir / "terreno_mdt.json"
    if not fp.exists():
        fp = "terreno_mdt.json"
        if not Path(fp).exists():
            return None
    try:
        return json.load(open(fp, encoding="utf-8"))
    except Exception:
        return None


def _cor_lados_por_movimento(dados, esc):
    """
    Coloração ILUSTRATIVA e QUALITATIVA dos lados do recinto de escavação
    conforme o nível de movimento dos alvos, associado por ORIENTAÇÃO.

    Como o terreno (coords nacionais) e os alvos (coords locais) estão em
    referenciais diferentes, não há correspondencia ponto-a-ponto. A associação
    e feita por lado/orientação: divide-se o contorno em 4 quadrantes (N, S, E,
    O) relativos ao centro, e a cada quadrante atribui-se a cor do pior estado
    dos alvos cuja orientação (relativa ao centro dos alvos) corresponde a esse
    lado. E uma leitura de tendência, não uma medida exata.

    Devolve lista de (indice_inicio, indice_fim, cor, etiqueta) para desenhar
    o contorno por troços coloridos.
    """
    import numpy as np
    alvos = dados.get("alvos")
    if alvos is None or alvos.empty:
        return None
    alvos = anexar_estado_calculado(alvos)
    ult = alvos[alvos[COLS["data"]] == alvos[COLS["data"]].max()].copy()

    # orientacao de cada alvo relativa ao centro dos alvos (referencial local)
    cxa, cya = ult[COLS["M0"]].mean(), ult[COLS["P0"]].mean()
    sev = {"Alarme": 3, "Alerta": 2, "Regular": 1, "Sem leitura": 0}
    cor_sev = {3: "#c0140f", 2: "#e67e00", 1: "#1f9e55", 0: "#9e9e9e"}
    # pior estado por quadrante de orientacao (N, E, S, O)
    pior = {"N": 0, "E": 0, "S": 0, "O": 0}
    for _, r in ult.iterrows():
        ang = np.degrees(np.arctan2(r[COLS["P0"]] - cya, r[COLS["M0"]] - cxa))
        if -45 <= ang < 45:      qd = "E"
        elif 45 <= ang < 135:    qd = "N"
        elif ang >= 135 or ang < -135: qd = "O"
        else:                    qd = "S"
        s = sev.get(r["Estado calculado"], 0)
        pior[qd] = max(pior[qd], s)

    # dividir o contorno da escavacao em 4 quadrantes pela orientacao ao centro
    esc_arr = np.array(esc)
    cxe, cye = esc_arr[:, 0].mean(), esc_arr[:, 1].mean()
    segmentos = []
    ini = 0
    qd_ant = None
    etiquetas_usadas = set()
    for idx in range(len(esc)):
        ang = np.degrees(np.arctan2(esc_arr[idx, 1] - cye, esc_arr[idx, 0] - cxe))
        if -45 <= ang < 45:      qd = "E"
        elif 45 <= ang < 135:    qd = "N"
        elif ang >= 135 or ang < -135: qd = "O"
        else:                    qd = "S"
        if qd_ant is None:
            qd_ant = qd
        if qd != qd_ant:
            sev_q = pior.get(qd_ant, 0)
            nome_q = {"N": "Lado N", "S": "Lado S", "E": "Lado E", "O": "Lado O"}[qd_ant]
            estado_q = {3: "alarme", 2: "alerta", 1: "regular", 0: "s/ dados"}[sev_q]
            lbl = f"{nome_q} ({estado_q})"
            # so mostrar cada etiqueta uma vez na legenda
            segmentos.append((ini, idx, cor_sev[sev_q],
                              lbl if lbl not in etiquetas_usadas else None))
            etiquetas_usadas.add(lbl)
            ini = idx
            qd_ant = qd
    # ultimo troco
    sev_q = pior.get(qd_ant, 0)
    nome_q = {"N": "Lado N", "S": "Lado S", "E": "Lado E", "O": "Lado O"}[qd_ant]
    estado_q = {3: "alarme", 2: "alerta", 1: "regular", 0: "s/ dados"}[sev_q]
    lbl = f"{nome_q} ({estado_q})"
    segmentos.append((ini, len(esc) - 1, cor_sev[sev_q],
                      lbl if lbl not in etiquetas_usadas else None))
    return segmentos


def _bbox_alinhamento(dados, terreno):
    """
    Calcula os parâmetros de alinhamento CAIXA-ENVOLVENTE entre o referencial
    dos alvos (local, M~5000/P~5050) e o do terreno (nacional, X~-110900/
    Y~-106300). Centra o conjunto dos alvos no centro do terreno e aplica uma
    escala ÚNICA proporcional (a menor razão de extensão, para não distorcer).
    Devolve um dicionario com centros, escala e extensão — a transformação em si
    e aplicada por _transformar_xy, que aceita ajustes finos do utilizador.

    NOTA IMPORTANTE (honestidade para a tese): este alinhamento e ILUSTRATIVO.
    Não há pontos de controlo comuns aos dois referenciais, por isso a posição
    de cada alvo sobre o terreno e APROXIMADA (erro típico de metros). Serve
    para ver a tendência do movimento sobre o relevo, não para medir. A precisão
    esta no separador Alvos (2D).
    """
    import numpy as np
    alvos = dados.get("alvos")
    if alvos is None or alvos.empty:
        return None
    ult = alvos[alvos[COLS["data"]] == alvos[COLS["data"]].max()]
    aM, aP = ult[COLS["M0"]].to_numpy(), ult[COLS["P0"]].to_numpy()
    faces = np.array(terreno["faces"]).reshape(-1, 3)
    aW = max(aM.max() - aM.min(), 1e-6)
    aH = max(aP.max() - aP.min(), 1e-6)
    tW = faces[:, 0].max() - faces[:, 0].min()
    tH = faces[:, 1].max() - faces[:, 1].min()
    return {
        "alvo_cx": float(aM.mean()), "alvo_cy": float(aP.mean()),
        "ter_cx": float(faces[:, 0].mean()), "ter_cy": float(faces[:, 1].mean()),
        "escala": float(min(tW / aW, tH / aH)),
    }


def _transformar_xy(M, P, par, rot_deg=0, flip_x=False, flip_y=False,
                    off_x=0.0, off_y=0.0):
    """Aplica a transformação caixa-envolvente (par de _bbox_alinhamento) a
    coordenadas de alvos, com ajustes finos opcionais (rotação em torno do
    centro, espelhamento e desvio manual). Devolve (X, Y) no referencial do
    terreno."""
    import numpy as np
    M = np.asarray(M, dtype=float); P = np.asarray(P, dtype=float)
    x = M - par["alvo_cx"]; y = P - par["alvo_cy"]
    if flip_x: x = -x
    if flip_y: y = -y
    th = np.radians(rot_deg)
    xr = x * np.cos(th) - y * np.sin(th)
    yr = x * np.sin(th) + y * np.cos(th)
    return (par["ter_cx"] + xr * par["escala"] + off_x,
            par["ter_cy"] + yr * par["escala"] + off_y)


def _transformar_vetor(dM, dP, par, rot_deg=0, flip_x=False, flip_y=False):
    """Transforma um VETOR de deslocamento (sem translação nem centro): só
    aplica espelhamento, rotação e escala. Para as setas de movimento assentarem
    coerentes com os alvos transformados."""
    import numpy as np
    dM = np.asarray(dM, dtype=float); dP = np.asarray(dP, dtype=float)
    x = -dM if flip_x else dM
    y = -dP if flip_y else dP
    th = np.radians(rot_deg)
    xr = x * np.cos(th) - y * np.sin(th)
    yr = x * np.sin(th) + y * np.cos(th)
    return xr * par["escala"], yr * par["escala"]


@st.cache_data
def _superficie_terreno(_terreno):
    """Prepara os pontos (XY -> Z) do MDT para interpolar a cota. Em cache
    porque o griddata reconstroi a triangulacao a cada chamada."""
    import numpy as np
    v = np.array(_terreno["faces"]).reshape(-1, 3)
    return v[:, :2], v[:, 2]


def _cota_terreno_em(X, Y, pts_xy, pts_z):
    """Interpola a cota da superfície do terreno nos pontos (X,Y) — para
    'drapejar' os alvos sobre o relevo. Linear dentro do casco; nearest fora."""
    import numpy as np
    from scipy.interpolate import griddata
    X = np.atleast_1d(np.asarray(X, dtype=float))
    Y = np.atleast_1d(np.asarray(Y, dtype=float))
    z = griddata(pts_xy, pts_z, (X, Y), method="linear")
    m = np.isnan(z)
    if m.any():
        z[m] = griddata(pts_xy, pts_z, (X[m], Y[m]), method="nearest")
    return z


def _casa_edificio_xy(fig, X, Y, zbase, cor, nome):
    """Como _casa_edificio, mas recebe coordenadas JÁ TRANSFORMADAS (X,Y no
    referencial do terreno) e a cota de base do terreno — para desenhar a casa
    ilustrativa do edifício vizinho sobre a superfície fundida."""
    import numpy as np
    x0, x1 = float(np.min(X)), float(np.max(X))
    y0, y1 = float(np.min(Y)), float(np.max(Y))
    mx = max((x1 - x0) * 0.2, 2.0); my = max((y1 - y0) * 0.2, 2.0)
    x0 -= mx; x1 += mx; y0 -= my; y1 += my
    h_parede, h_telhado = 5.0, 3.0
    zt = zbase + h_parede; zc = zt + h_telhado
    xs = [x0, x1, x1, x0, x0, x1, x1, x0]
    ys = [y0, y0, y1, y1, y0, y0, y1, y1]
    zs = [zbase, zbase, zbase, zbase, zt, zt, zt, zt]
    fig.add_trace(go.Mesh3d(
        x=xs, y=ys, z=zs,
        i=[0, 0, 0, 4, 1, 1, 2, 3, 0, 3],
        j=[1, 2, 4, 5, 2, 5, 3, 7, 3, 7],
        k=[2, 3, 5, 7, 5, 6, 7, 4, 7, 4],
        color=cor, opacity=0.30, name=nome, hoverinfo="name",
        showlegend=False, flatshading=True))
    ym = (y0 + y1) / 2.0
    fig.add_trace(go.Mesh3d(
        x=[x0, x1, x1, x0, x0, x1], y=[y0, y0, y1, y1, ym, ym],
        z=[zt, zt, zt, zt, zc, zc],
        i=[0, 1, 3, 2, 0, 1], j=[1, 4, 2, 5, 4, 4], k=[4, 5, 5, 4, 3, 0],
        color=cor, opacity=0.45, hoverinfo="skip",
        showlegend=False, flatshading=True))
    fig.add_trace(go.Scatter3d(
        x=[x0, x1], y=[ym, ym], z=[zc, zc], mode="lines",
        line=dict(color=cor, width=4), showlegend=False, hoverinfo="skip"))


def separador_terreno3d(dados):
    """
    3D do terreno REAL, a partir do levantamento topográfico (MDT triangulado),
    no referencial da obra e com cotas Z reais. Mostra a superfície do terreno,
    a linha de escavação (cota 4,55) e o volume escavado. Ao contrário do 3D dos
    alvos (referencial local, cotas relativas), este assenta em cotas absolutas.
    """
    st.subheader("Terreno 3D — modelo do levantamento topográfico")
    st.caption("Superfície real do terreno (modelo digital triangulado) do "
               "levantamento topográfico, no referencial da obra e com COTAS "
               "REAIS. Mostra a topografia original e o volume a escavar até a "
               "cota de fundo (4,55 m). Complementa o 3D dos alvos — que usa um "
               "referencial local e cotas relativas — com geometria absoluta.")

    terreno = _carregar_terreno()
    if terreno is None:
        st.info("Modelo de terreno não disponível (falta terreno_mdt.json).")
        return

    import numpy as np
    faces = np.array(terreno["faces"])           # (N,3,3)
    z_fundo = terreno.get("cota_fundo", 4.55)
    z_max_terreno = float(faces.reshape(-1, 3)[:, 2].max())  # cota mais alta

    c1, c2 = st.columns(2)
    with c1:
        mostrar_escav = st.checkbox("Mostrar volume de escavação", value=True)
        mostrar_curvas = st.checkbox("Curvas de nível", value=False,
                                     help="Curvas de nível do levantamento, que "
                                          "dao a leitura do relevo (alçados).")
        colorir_lados = st.checkbox(
            "Colorir alçados por movimento (ilustrativo)", value=False,
            help="Colore os lados do recinto conforme o nível de movimento dos "
                 "alvos próximos, por orientação. É QUALITATIVO e ilustrativo: "
                 "o terreno e os alvos estão em referenciais diferentes, pelo "
                 "que a associação e por lado, não por ponto exato.")
    with c2:
        exagero = st.slider("Exagero vertical", 1.0, 4.0, 1.5, 0.5,
                            help="Amplia a escala vertical para realçar o "
                                 "relevo. 1.0 = escala real.")
        # filtro por fase de escavacao: ate que piso ja se escavou
        fases_esc = ["Terreno original (sem escavar)"] + \
            [f"Até {nome} (cota {cota:.2f})" for nome, cota in COTAS_PISOS
             if cota < z_max_terreno] + \
            [f"Escavação completa (fundo {z_fundo})"]
        fase_esc = st.selectbox(
            "Fase de escavação a representar", fases_esc,
            index=len(fases_esc) - 1,
            help="Mostra a escavação até a cota do piso correspondente. As "
                 "cotas são as do projeto; a escavação real progride por fases.")

    # cota de escavacao correspondente a fase escolhida
    if fase_esc.startswith("Terreno original"):
        cota_escav_fase = None                 # nao escavado
    elif fase_esc.startswith("Escavação completa"):
        cota_escav_fase = z_fundo
    else:
        # extrair a cota do texto "...(cota XX.XX)"
        import re
        m = re.search(r"cota ([\d.]+)", fase_esc)
        cota_escav_fase = float(m.group(1)) if m else z_fundo

    # construir a malha Mesh3d a partir dos triangulos
    verts = faces.reshape(-1, 3)
    i = np.arange(0, len(verts), 3)
    j = i + 1
    k = i + 2
    x, y, z = verts[:, 0], verts[:, 1], verts[:, 2]
    z_base = z.min()
    z_plot = z_base + (z - z_base) * exagero

    fig = go.Figure()
    fig.add_trace(go.Mesh3d(
        x=x, y=y, z=z_plot, i=i, j=j, k=k,
        intensity=z, colorscale="earth", opacity=0.92,
        colorbar=dict(title="Cota (m)"),
        name="Terreno", hovertemplate="Cota: %{intensity:.1f} m<extra></extra>"))

    # ---- volume de escavacao ate a cota da FASE escolhida ----------------
    esc = terreno.get("escavacao", [])
    if esc and mostrar_escav and cota_escav_fase is not None:
        ex = [p[0] for p in esc]
        ey = [p[1] for p in esc]
        zf = z_base + (cota_escav_fase - z_base) * exagero

        # coloracao ILUSTRATIVA dos lados por movimento dos alvos
        cor_fundo = "#b45309"
        segmentos_cor = None
        if colorir_lados:
            segmentos_cor = _cor_lados_por_movimento(dados, esc)

        # contorno do fundo (a cota da fase)
        if segmentos_cor is None:
            fig.add_trace(go.Scatter3d(
                x=ex, y=ey, z=[zf] * len(ex), mode="lines",
                line=dict(color=cor_fundo, width=4),
                name=f"Escavação até cota {cota_escav_fase:.2f}"))
        else:
            # desenhar o contorno por segmentos coloridos conforme o lado
            for (i0, i1, cor, lbl) in segmentos_cor:
                fig.add_trace(go.Scatter3d(
                    x=ex[i0:i1+1], y=ey[i0:i1+1], z=[zf]*(i1-i0+1),
                    mode="lines", line=dict(color=cor, width=6),
                    name=lbl))

        # paredes verticais da escavacao (do terreno ate a cota da fase)
        for idx in range(0, len(esc) - 1, 3):
            fig.add_trace(go.Scatter3d(
                x=[ex[idx], ex[idx]], y=[ey[idx], ey[idx]],
                z=[z_plot.max(), zf], mode="lines",
                line=dict(color="rgba(180,83,9,0.25)", width=1),
                showlegend=False, hoverinfo="skip"))

    # curvas de nivel (dao a leitura do relevo / alcados do terreno)
    if mostrar_curvas:
        curvas = terreno.get("curvas_nivel", [])
        cx, cy, cz = [], [], []
        for c in curvas:
            zc = c[0][2] if len(c[0]) > 2 else 0
            if zc < 1:            # ignorar curvas sem cota valida
                continue
            for p in c:
                cx.append(p[0]); cy.append(p[1])
                cz.append(z_base + (p[2] - z_base) * exagero)
            cx.append(None); cy.append(None); cz.append(None)
        if cx:
            fig.add_trace(go.Scatter3d(
                x=cx, y=cy, z=cz, mode="lines",
                line=dict(color="rgba(60,40,20,0.5)", width=1),
                name="Curvas de nível", hoverinfo="skip"))

    fig.update_layout(
        height=640,
        scene=dict(
            xaxis_title="M (m)", yaxis_title="P (m)", zaxis_title="Cota (m)",
            aspectmode="data"),
        margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig, use_container_width=True)

    st.caption(f"Superfície do terreno entre as cotas {z.min():.1f} e "
               f"{z.max():.1f} m; fundo de escavação a {z_fundo} m — "
               f"aproximadamente {z.max()-z_fundo:.0f} m de altura escavada no "
               f"ponto mais alto. Fonte: levantamento topográfico (DXF). Notas: "
               f"o exagero vertical é apenas visual; o filtro de fase mostra a "
               f"escavação até a cota do piso do projeto (a progressão real e "
               f"por fases). A coloração dos alçados por movimento, quando "
               f"ativa, e ILUSTRATIVA e qualitativa — associa o estado dos "
               f"alvos ao lado do recinto por orientação, não por coordenada "
               f"exata, porque terreno e alvos usam referenciais distintos.")


def _detetar_cantos_contorno(pts, n_cantos=4):
    """Deteta os cantos de um contorno fechado (ângulo interno mais fechado),
    espacados entre si. Usado para posicionar as escoras de canto."""
    import numpy as np
    pts = np.asarray(pts)
    n = len(pts)
    angs = []
    for i in range(n):
        a = pts[(i - 3) % n]; b = pts[i]; c = pts[(i + 3) % n]
        v1 = a - b; v2 = c - b
        cos = np.dot(v1, v2) / (np.linalg.norm(v1) * np.linalg.norm(v2) + 1e-9)
        angs.append(np.arccos(np.clip(cos, -1, 1)))
    angs = np.array(angs)
    cantos = []
    for i in np.argsort(angs):
        if all(min(abs(i - j), n - abs(i - j)) > n // 8 for j in cantos):
            cantos.append(int(i))
        if len(cantos) >= n_cantos:
            break
    return sorted(cantos)


def _elementos_contencao_esquematicos(fig, esc, zex, cotas_pisos,
                                      coroamento, fundo):
    """
    Desenha, de forma ESQUEMÁTICA mas INFORMADA PELO PROJETO, as soluções de
    contenção periférica: viga de coroamento/distribuição, bandas de laje e
    escoramentos. Baseado nas peças desenhadas da JETsj (EDN-JET-ZZ-ZZ-DR-U-
    0021..0027): cotas dos pisos confirmadas com o projeto, cores das lajes
    conforme a legenda oficial, e escoras HORIZONTAIS entre lados opostos as
    cotas dos pisos (como nos cortes tipo), não diagonais de canto.

    ATENÇÃO — HONESTIDADE: a FORMA em planta (o contorno curvo real, a geometria
    exata de cada banda) não foi extraída do PDF — a sua reprodução vetorial não
    era fiável. Por isso as bandas são aneis perimetrais aproximados sobre o
    contorno disponível, e as escoras são troços representativos. As COTAS e as
    CORES são reais; a forma e esquemática. A representação rigorosa esta nas
    peças desenhadas do projeto e nas VISTAS 3D do projeto (mostradas no fim
    deste separador).
    """
    import numpy as np
    esc = np.asarray(esc, dtype=float)
    n = len(esc)
    ex, ey = esc[:, 0], esc[:, 1]
    cx, cy = ex.mean(), ey.mean()

    # 1) VIGA DE COROAMENTO / DISTRIBUICAO — anel no topo (cota real)
    zc = float(zex(coroamento))
    fig.add_trace(go.Scatter3d(
        x=list(ex) + [ex[0]], y=list(ey) + [ey[0]], z=[zc] * (n + 1),
        mode="lines", line=dict(color="#3a3a3a", width=6),
        name="Viga de coroamento/distribuição (projeto: cota real)",
        hovertemplate="Viga de coroamento/distribuição<br>"
                      f"cota {coroamento:.2f} m (projeto JETsj)<extra></extra>"))

    # 2) BANDAS DE LAJE — aneis perimetrais as cotas reais dos pisos, com as
    #    CORES OFICIAIS do projeto (ciclo pelas cores da legenda, ja que o
    #    mapeamento exato piso->espessura consta das pecas desenhadas).
    interior = np.column_stack([cx + (ex - cx) * 0.90, cy + (ey - cy) * 0.90])
    ix, iy = interior[:, 0], interior[:, 1]
    cores_laje = list(CORES_LAJE_PROJETO.values())
    pisos_dentro = [(nm, ct) for nm, ct in cotas_pisos if fundo < ct < coroamento]
    for idx_p, (nome, cota) in enumerate(pisos_dentro):
        cor = cores_laje[idx_p % len(cores_laje)]
        zp = float(zex(cota))
        wx, wy, wz, wi, wj, wk = [], [], [], [], [], []
        for i in range(n - 1):
            b = len(wx)
            wx += [ex[i], ex[i+1], ix[i+1], ix[i]]
            wy += [ey[i], ey[i+1], iy[i+1], iy[i]]
            wz += [zp, zp, zp, zp]
            wi += [b, b]; wj += [b + 1, b + 2]; wk += [b + 2, b + 3]
        fig.add_trace(go.Mesh3d(
            x=wx, y=wy, z=wz, i=wi, j=wj, k=wk,
            color=cor, opacity=0.55, flatshading=True,
            name=f"Banda de laje {nome} (cor do projeto)",
            hovertext=f"Banda de laje {nome} — cota {cota:.2f} m "
                      f"(cor conforme legenda do projeto)",
            hoverinfo="text"))

    # 3) ESCORAMENTOS METALICOS PROVISORIOS — HORIZONTAIS entre lados opostos,
    #    as cotas dos pisos (como nos cortes tipo do projeto). Esquematico na
    #    posicao (nº e vao representativos), fiel no tipo (horizontal, a verde).
    #    Ligam cada ponto de um lado ao ponto oposto (atravessando o recinto).
    ang = np.arctan2(ey - cy, ex - cx)
    ordem = np.argsort(ang)
    escora_verde = "#1f9e2f"
    xs, ys, zs = [], [], []
    # 2 cotas representativas para nao poluir: pisos -1 e -3 (intermedios)
    cotas_escora = [ct for nm, ct in pisos_dentro
                    if any(t in nm for t in ("-1", "-3"))]
    if not cotas_escora and pisos_dentro:
        cotas_escora = [pisos_dentro[len(pisos_dentro) // 2][1]]
    for cota in cotas_escora:
        ze = float(zex(cota))
        # 4 escoras a atravessar, ligando pontos opostos do contorno
        for f in np.linspace(0.15, 0.85, 4):
            a = ordem[int(f * (n - 1))]
            b = ordem[int(((f + 0.5) % 1.0) * (n - 1))]
            xs += [ex[a], ex[b], None]
            ys += [ey[a], ey[b], None]
            zs += [ze, ze, None]
    fig.add_trace(go.Scatter3d(
        x=xs, y=ys, z=zs, mode="lines",
        line=dict(color=escora_verde, width=5),
        name="Escoras metálicas provisórias (esquemático)",
        hovertemplate="Escora metálica provisória (horizontal)<br>"
                      "posição esquemática<extra></extra>"))


def separador_terreno_alvos_3d(dados):
    """
    SEPARADOR FUNDIDO: terreno real (MDT) + alvos com movimento drapejados
    sobre a superfície. Junta o melhor dos dois 3D antigos — o relevo real com
    curvas de nível e volume de escavação, e os alvos coloridos por estado com
    setas de deslocamento, casas dos edifícios e destaque do Edifício 1.

    Referenciais diferentes: os alvos são alinhados ao terreno por CAIXA-
    ENVOLVENTE (ilustrativo, ver _bbox_alinhamento) e assentes na cota do
    terreno (drapejados). A posição e APROXIMADA; a precisão esta no Alvos 2D.
    """
    import numpy as np
    st.subheader("Terreno + Alvos 3D — movimento sobre o relevo real")
    st.caption("Funde o modelo real do terreno (levantamento topográfico, com "
               "curvas de nível e volume de escavação) com os alvos e o seu "
               "movimento. Os alvos estão num referencial diferente do terreno, "
               "por isso são POUSADOS sobre a superfície por alinhamento "
               "aproximado (caixa-envolvente) — a posição é ILUSTRATIVA, para "
               "ler a tendência do movimento sobre o relevo. A medição precisa "
               "está no separador Alvos (2D).")

    terreno = _carregar_terreno()
    if terreno is None:
        st.info("Modelo de terreno não disponível (falta terreno_mdt.json).")
        return
    alvos = dados.get("alvos")
    if alvos is None or alvos.empty:
        st.info("Sem dados de alvos.")
        return

    par = _bbox_alinhamento(dados, terreno)
    if par is None:
        st.info("Não foi possível alinhar alvos e terreno.")
        return

    faces = np.array(terreno["faces"])
    v = faces.reshape(-1, 3)
    z_fundo = terreno.get("cota_fundo", 4.55)
    z_max_terreno = float(v[:, 2].max())
    pts_xy, pts_z = _superficie_terreno(terreno)

    datas = sorted(alvos[COLS["data"]].dropna().unique())
    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        data_sel = st.select_slider(
            "Campanha", options=datas, value=datas[-1],
            format_func=lambda d: pd.to_datetime(d).strftime("%d/%m/%Y"),
            key="fus_campanha")
        fator = st.slider("Amplificação do deslocamento", 50, 2000, 500, 50,
                          key="fus_fator",
                          help="Deslocamentos milimétricos sobre coordenadas em "
                               "metros; amplia-se para se verem.")
        exagero = st.slider("Exagero vertical do terreno", 1.0, 4.0, 1.5, 0.5,
                            key="fus_exag",
                            help="Amplia só a escala vertical do relevo.")
    with col_b:
        mostrar_curvas = st.checkbox("Curvas de nível", value=True,
                                     key="fus_curvas")
        mostrar_escav = st.checkbox("Volume de escavação", value=True,
                                    key="fus_escav")
        mostrar_casas = st.checkbox("Casas dos edifícios", value=True,
                                    key="fus_casas")
        identificar_edif = st.checkbox("Identificar edifícios", value=True,
                                       key="fus_idedif")
        mostrar_contencao = st.checkbox(
            "Elementos de contenção (esquemático)", value=False,
            key="fus_contencao",
            help="Desenha, de forma ESQUEMÁTICA e ILUSTRATIVA, a viga de "
                 "coroamento, as bandas de laje (as cotas reais dos pisos) e "
                 "os escoramentos de canto. A geometria de projeto destes "
                 "elementos não consta dos dados — a forma é assumida, só as "
                 "cotas das lajes/viga são reais. Os escoramentos são "
                 "totalmente ilustrativos.")
    with col_c:
        destacar_alarmes = st.checkbox("Destacar alarmes/alertas", value=True,
                                       key="fus_alarmes")
        destacar_sc = st.checkbox("Realçar Edifício 1", value=True,
                                  key="fus_sc")
        # ajustes finos do alinhamento (o utilizador valida a olho)
        with st.expander("Ajuste fino do alinhamento"):
            rot = st.select_slider("Rotação", [0, 90, 180, 270], value=0,
                                   key="fus_rot")
            flip_x = st.checkbox("Espelhar horizontal", value=False,
                                 key="fus_flipx")
            flip_y = st.checkbox("Espelhar vertical", value=False,
                                 key="fus_flipy")

    # fase de escavacao segue a campanha? Mantemos escavacao completa por
    # simplicidade (a progressao real ve-se no Terreno 3D antigo / Alvos 2D).
    cota_escav_fase = z_fundo

    campanha = alvos[alvos[COLS["data"]] == data_sel].copy()
    campanha = anexar_estado_calculado(campanha)

    # ---- cotas Z do terreno com exagero ----
    z_base = v[:, 2].min()

    def _zex(zabs):
        return z_base + (np.asarray(zabs) - z_base) * exagero

    fig = go.Figure()

    # ---- superficie do terreno ----
    i = np.arange(0, len(v), 3); j = i + 1; k = i + 2
    fig.add_trace(go.Mesh3d(
        x=v[:, 0], y=v[:, 1], z=_zex(v[:, 2]), i=i, j=j, k=k,
        intensity=v[:, 2], colorscale="earth", opacity=0.9,
        colorbar=dict(title="Cota (m)"), name="Terreno",
        hovertemplate="Cota: %{intensity:.1f} m<extra></extra>"))

    # ---- volume de escavacao (paredes SOLIDAS que encaixam no terreno) ----
    esc = terreno.get("escavacao", [])
    if esc and mostrar_escav:
        ex = np.array([p[0] for p in esc], dtype=float)
        ey = np.array([p[1] for p in esc], dtype=float)
        zf = float(_zex(cota_escav_fase))
        # topo de cada ponto do contorno = cota REAL do terreno ali (para a
        # parede encaixar na superficie em vez de partir de uma cota unica).
        ztopo = _cota_terreno_em(ex, ey, pts_xy, pts_z)
        ztopo = _zex(ztopo)
        # parede como faixa continua de triangulos (Mesh3d): para cada par de
        # pontos consecutivos, um quad (2 triangulos) do topo ate ao fundo.
        n = len(ex)
        wx, wy, wz, wi, wj, wk = [], [], [], [], [], []
        for idx in range(n - 1):
            b = len(wx)
            wx += [ex[idx], ex[idx+1], ex[idx+1], ex[idx]]
            wy += [ey[idx], ey[idx+1], ey[idx+1], ey[idx]]
            wz += [ztopo[idx], ztopo[idx+1], zf, zf]
            wi += [b + 0, b + 0]; wj += [b + 1, b + 2]; wk += [b + 2, b + 3]
        fig.add_trace(go.Mesh3d(
            x=wx, y=wy, z=wz, i=wi, j=wj, k=wk,
            color="#b45309", opacity=0.45, flatshading=True,
            name="Paredes de escavação", hoverinfo="skip"))
        # contorno do fundo (linha fechada, cota de fundo)
        fig.add_trace(go.Scatter3d(
            x=list(ex) + [ex[0]], y=list(ey) + [ey[0]], z=[zf] * (n + 1),
            mode="lines", line=dict(color="#8B4513", width=4),
            name=f"Fundo de escavação ({cota_escav_fase:.2f} m)"))
        # aresta de topo (onde a escavacao corta a superficie)
        fig.add_trace(go.Scatter3d(
            x=list(ex), y=list(ey), z=list(ztopo), mode="lines",
            line=dict(color="rgba(139,69,19,0.6)", width=2),
            showlegend=False, hoverinfo="skip"))

    # ---- elementos de contencao periferica (ESQUEMATICOS) ----
    if mostrar_contencao and esc:
        _elementos_contencao_esquematicos(
            fig, np.array([[p[0], p[1]] for p in esc]), _zex,
            COTAS_PISOS, COTA_COROAMENTO_PADRAO, z_fundo)

    # ---- curvas de nivel ----
    if mostrar_curvas:
        cx, cy, cz = [], [], []
        for c in terreno.get("curvas_nivel", []):
            zc = c[0][2] if len(c[0]) > 2 else 0
            if zc < 1:
                continue
            for p in c:
                cx.append(p[0]); cy.append(p[1]); cz.append(float(_zex(p[2])))
            cx.append(None); cy.append(None); cz.append(None)
        if cx:
            fig.add_trace(go.Scatter3d(
                x=cx, y=cy, z=cz, mode="lines",
                line=dict(color="rgba(60,40,20,0.5)", width=1),
                name="Curvas de nível", hoverinfo="skip"))

    # ---- alvos drapejados + movimento ----
    COR_ESTADO = {"Alarme": "#c0140f", "Alerta": "#e67e00", "Regular": "#1f9e55"}
    seg_x, seg_y, seg_z, seg_cor = [], [], [], []
    cone_x, cone_y, cone_z, cone_u, cone_v, cone_w = ([] for _ in range(6))
    dz_alt = 2.0  # levantar os alvos um pouco acima da superficie, p/ se verem

    for chave, grp in campanha.groupby(COLS["edificio"]):
        tipo, etiqueta = classificar_grupo(chave)
        M0 = grp[COLS["M0"]].to_numpy(); P0 = grp[COLS["P0"]].to_numpy()
        # transformar posicao para o referencial do terreno
        X, Y = _transformar_xy(M0, P0, par, rot, flip_x, flip_y)
        Zs = _cota_terreno_em(X, Y, pts_xy, pts_z)
        Zs = _zex(Zs) + dz_alt
        # transformar vetores de deslocamento (mm -> m -> amplificado)
        dM = grp[COLS["dM"]].to_numpy() / 1000.0 * fator
        dP = grp[COLS["dP"]].to_numpy() / 1000.0 * fator
        dZ = grp[COLS["dZ"]].to_numpy() / 1000.0 * fator * exagero
        dX, dY = _transformar_vetor(dM, dP, par, rot, flip_x, flip_y)

        dh = grp[COLS["desl_h"]].to_numpy()
        nomes = grp[COLS["alvo"]].astype(str).to_numpy()
        estados = grp["Estado calculado"].to_numpy()
        fachadas = grp["Fachada"].to_numpy()
        e_edificio1 = isinstance(chave, str) and "Edifício 1" in chave

        if tipo == "edificio":
            cor = CORES_EDIFICIO.get(chave, "#7f7f7f"); nome_leg = chave
        else:
            cor = "#ff7f0e"; nome_leg = f"Contenção — Alçado {etiqueta}"

        if destacar_sc and e_edificio1:
            simbolos = ["diamond" if f == "Frente escavação" else "circle"
                        for f in fachadas]
        else:
            simbolos = "circle"
        if destacar_alarmes:
            cor_borda = [COR_ESTADO.get(e, "rgba(0,0,0,0.2)") for e in estados]
            larg = 4 if any(e in ("Alarme", "Alerta") for e in estados) else 1
        else:
            cor_borda = "rgba(0,0,0,0.2)"; larg = 1

        cd = np.column_stack([dh, estados, fachadas])
        fig.add_trace(go.Scatter3d(
            x=X + dX, y=Y + dY, z=Zs + dZ, mode="markers+text",
            marker=dict(size=6, color=cor, symbol=simbolos,
                        line=dict(color=cor_borda, width=larg)),
            text=nomes, textposition="top center", textfont=dict(size=8),
            name=nome_leg, customdata=cd,
            hovertemplate="Alvo %{text}<br>Desl. h: %{customdata[0]:.1f} mm"
                          "<br>Estado: %{customdata[1]}<br>%{customdata[2]}"
                          "<extra>" + nome_leg + "</extra>"))

        for n in range(len(X)):
            c = (COR_ESTADO.get(estados[n], "#888") if destacar_alarmes
                 else "crimson")
            seg_x += [X[n], X[n] + dX[n], None]
            seg_y += [Y[n], Y[n] + dY[n], None]
            seg_z += [Zs[n], Zs[n] + dZ[n], None]
            seg_cor.append(c)
            cone_x.append(X[n] + dX[n]); cone_y.append(Y[n] + dY[n])
            cone_z.append(Zs[n] + dZ[n])
            cone_u.append(dX[n]); cone_v.append(dY[n]); cone_w.append(dZ[n])

        if identificar_edif and tipo == "edificio" and len(X):
            fig.add_trace(go.Scatter3d(
                x=[X.mean()], y=[Y.mean()], z=[Zs.max() + 6],
                mode="text", text=[f"<b>{chave}</b>"],
                textfont=dict(size=12, color=cor),
                showlegend=False, hoverinfo="skip"))
        if mostrar_casas and tipo == "edificio" and len(X):
            _casa_edificio_xy(fig, X, Y, float(Zs.min()) - dz_alt, cor, chave)

    # hastes das setas por cor
    for c in set(seg_cor):
        xs, ys, zs = [], [], []
        for jj, cc in enumerate(seg_cor):
            if cc == c:
                xs += seg_x[3*jj:3*jj+3]; ys += seg_y[3*jj:3*jj+3]
                zs += seg_z[3*jj:3*jj+3]
        fig.add_trace(go.Scatter3d(x=xs, y=ys, z=zs, mode="lines",
                                   line=dict(color=c, width=4),
                                   showlegend=False, hoverinfo="skip"))
    if cone_x:
        fig.add_trace(go.Cone(
            x=cone_x, y=cone_y, z=cone_z, u=cone_u, v=cone_v, w=cone_w,
            sizemode="absolute", sizeref=1.2, anchor="tip", showscale=False,
            colorscale=[[0, "#555"], [1, "#555"]], hoverinfo="skip",
            showlegend=False, opacity=0.9))

    fig.update_layout(
        height=760,
        scene=dict(xaxis_title="M (m)", yaxis_title="P (m)",
                   zaxis_title="Cota (m)", aspectmode="data"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    font=dict(size=9)),
        margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig, use_container_width=True)

    # metricas
    desl_h_all = campanha[COLS["desl_h"]].to_numpy()
    nomes_all = campanha[COLS["alvo"]].astype(str).to_numpy()
    n_alarme = int((campanha["Estado calculado"] == "Alarme").sum())
    n_alerta = int((campanha["Estado calculado"] == "Alerta").sum())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alvos na campanha", len(campanha))
    c2.metric("Desl. horizontal máx. (mm)", f"{np.nanmax(desl_h_all):.1f}")
    c3.metric("Em alarme", n_alarme)
    c4.metric("Em alerta", n_alerta)
    idx = int(np.nanargmax(desl_h_all))
    st.caption(
        f"O alvo mais afetado ({nomes_all[idx]}, {np.nanmax(desl_h_all):.1f} mm) "
        f"está drapejado sobre o relevo real. Vermelho = alarme, laranja = "
        f"alerta, verde = regular. IMPORTANTE: a posição dos alvos sobre o "
        f"terreno é aproximada (alinhamento por caixa-envolvente entre "
        f"referenciais diferentes) — para leitura tendencial do movimento sobre "
        f"o relevo, não métrica. A medição precisa está no separador Alvos (2D). "
        f"Se algum alvo cair no lado errado, usa o 'Ajuste fino do alinhamento'.")
    if mostrar_contencao:
        st.warning(
            "⚠ Os elementos de contenção (viga de coroamento/distribuição, "
            "bandas de laje, escoras metálicas) são ESQUEMÁTICOS mas INFORMADOS "
            "PELO PROJETO: as cotas dos pisos e as cores das lajes seguem as "
            "peças desenhadas da JETsj (EDN-JET-ZZ-ZZ-DR-U-0021..0027), e as "
            "escoras são horizontais entre lados (como nos cortes tipo). A FORMA "
            "em planta (contorno curvo, geometria exata de cada banda) e "
            "aproximada — não foi extraída do projeto. Para a representação "
            "rigorosa, ver as Vistas 3D do projeto no fim deste separador.")

    # ---- Vistas 3D do proprio projeto (referencia fiel, JETsj) ----
    with st.expander("📐 Vistas 3D do projeto de contenção (JETsj) — referência"):
        st.caption("Imagens das peças desenhadas do projeto de contenção "
                   "periférica (JETsj, pranchas EDN-JET-ZZ-ZZ-DR-U-0002 a 0005). "
                   "São a representação AUTORITATIVA e rigorosa da solução — "
                   "cortina de estacas, bandas de laje por cota, ancoragens, "
                   "escoras metálicas e faseamento. O 3D interativo acima é uma "
                   "leitura do movimento dos alvos; estas vistas são a geometria "
                   "de projeto.")
        base_dir = Path(__file__).resolve().parent
        vistas = [
            ("vista3d_1de4.jpg", "Vistas 3D (1/4) — vista global"),
            ("vista3d_2de4.jpg", "Vistas 3D (2/4)"),
            ("vista3d_3de4.jpg", "Vistas 3D (3/4)"),
            ("vista3d_4de4.jpg", "Vista 3D (4/4)"),
        ]
        alguma = False
        for fich, legenda in vistas:
            fp = base_dir / "projeto_vistas" / fich
            if not fp.exists():
                fp = Path("projeto_vistas") / fich
            if fp.exists():
                st.image(str(fp), caption=legenda, use_container_width=True)
                alguma = True
        if not alguma:
            st.info("Imagens das vistas 3D do projeto não encontradas "
                    "(pasta 'projeto_vistas/'). Verifica que foi publicada "
                    "junto com a app.")


def separador_resultados(dados):
    """
    RESULTADOS — quantificação factual das movimentações observadas pela
    instrumentação. Magnitude, distribuição espacial, ritmo, profundidade e
    comparação com a estimativa de projeto. São ILUSTRACOES DOS RESULTADOS,
    sem interpretação de causas: a correlação entre instrumentos esta no
    separador Correlações; a interpretação do PORQUE e do autor no texto.
    """
    import numpy as np
    st.subheader("Resultados — movimentações observadas")
    st.caption("Quantificação factual do que a instrumentação mediu: quanto, "
               "onde, a que ritmo e a que profundidade, e como se compara com a "
               "estimativa de projeto. Ilustração de resultados, sem "
               "interpretação de causas.")

    alvos = dados.get("alvos")
    if alvos is None or alvos.empty:
        st.info("Sem dados de alvos.")
        return
    alvos = alvos.copy()
    alvos[COLS["data"]] = pd.to_datetime(alvos[COLS["data"]], errors="coerce")
    ult_data = alvos[COLS["data"]].max()
    ult = anexar_estado_calculado(alvos[alvos[COLS["data"]] == ult_data].copy())

    # ---- 1. QUANTO ----
    st.markdown("#### 1. Quanto — magnitude das movimentações")
    n_alarme = int((ult["Estado calculado"] == "Alarme").sum())
    n_alerta = int((ult["Estado calculado"] == "Alerta").sum())
    dh = ult[COLS["desl_h"]]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alvos monitorizados", ult[COLS["alvo"]].nunique())
    c2.metric("Desl. H máximo (mm)", f"{dh.max():.1f}")
    c3.metric("Em alarme", n_alarme)
    c4.metric("Em alerta", n_alerta)

    top = ult.nlargest(10, COLS["desl_h"])[
        [COLS["alvo"], COLS["edificio"], COLS["desl_h"], COLS["dZ"],
         "Estado calculado"]].copy()
    top.columns = ["Alvo", "Edifício / elemento", "Desl. H (mm)",
                   "Desl. V ΔZ (mm)", "Estado"]
    top["Desl. H (mm)"] = top["Desl. H (mm)"].round(1)
    top["Desl. V ΔZ (mm)"] = top["Desl. V ΔZ (mm)"].round(1)
    st.markdown("**Ranking dos 10 alvos mais deslocados** (última campanha, "
                f"{ult_data.strftime('%d/%m/%Y')}):")
    st.dataframe(top, use_container_width=True, hide_index=True)
    top1 = top.iloc[0]
    st.caption(f"Maior deslocamento horizontal: alvo {top1['Alvo']} "
               f"({top1['Desl. H (mm)']:.1f} mm), no elemento "
               f"«{top1['Edifício / elemento']}».")

    # ---- 2. ONDE ----
    st.divider()
    st.markdown("#### 2. Onde — distribuição espacial")
    g = ult.groupby(COLS["edificio"])[COLS["desl_h"]].agg(
        media="mean", maximo="max", n="count").reset_index()
    g = g.sort_values("maximo", ascending=False)
    g_plot = g.head(12)
    fig_b = go.Figure()
    fig_b.add_trace(go.Bar(
        y=g_plot[COLS["edificio"]], x=g_plot["maximo"].round(1),
        orientation="h", name="Máximo",
        marker=dict(color="#c0140f"),
        text=g_plot["maximo"].round(1), textposition="auto"))
    fig_b.add_trace(go.Bar(
        y=g_plot[COLS["edificio"]], x=g_plot["media"].round(1),
        orientation="h", name="Média",
        marker=dict(color="#f0a080")))
    fig_b.update_layout(
        height=460, barmode="overlay",
        xaxis=dict(title="Desl. horizontal (mm)"),
        yaxis=dict(autorange="reversed"),
        margin=dict(l=10, r=10, t=30, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig_b, use_container_width=True)
    lider = g.iloc[0]
    seg = g.iloc[1] if len(g) > 1 else None
    txt_onde = (f"Elemento com maior movimentação: «{lider[COLS['edificio']]}» "
                f"(máx {lider['maximo']:.1f} mm, média {lider['media']:.1f} mm).")
    if seg is not None:
        racio = lider["maximo"] / seg["maximo"] if seg["maximo"] else float("nan")
        txt_onde += (f" O segundo é «{seg[COLS['edificio']]}» (máx "
                     f"{seg['maximo']:.1f} mm) — o primeiro move-se "
                     f"{racio:.1f}x mais.")
    st.caption(txt_onde)

    # ---- 3. QUANDO / RITMO ----
    st.divider()
    st.markdown("#### 3. Quando e a que ritmo — velocidade de movimentação")
    st.caption("Velocidade (mm/dia entre campanhas) do deslocamento horizontal "
               "de um alvo. Mostra quando o movimento acelerou.")
    alvos_disp = sorted(ult.nlargest(15, COLS["desl_h"])[COLS["alvo"]]
                        .astype(str).tolist())
    alvo_sel = st.selectbox("Alvo a analisar (velocidade)", alvos_disp,
                            key="res_alvo_vel")
    s = alvos[alvos[COLS["alvo"]].astype(str) == alvo_sel].sort_values(
        COLS["data"])[[COLS["data"], COLS["desl_h"]]].dropna()
    s["dias"] = (s[COLS["data"]] - s[COLS["data"]].shift()).dt.days
    s["dincr"] = s[COLS["desl_h"]].diff()
    s["vel"] = s["dincr"] / s["dias"]
    s_v = s.dropna(subset=["vel"])
    if not s_v.empty:
        fig_c = go.Figure()
        fig_c.add_trace(go.Bar(
            x=[pd.to_datetime(v) for v in s_v[COLS["data"]].tolist()],
            y=[float(v) for v in s_v["vel"].tolist()],
            name="Velocidade (mm/dia)", marker=dict(color="#1f78d1")))
        fig_c.update_layout(
            height=340, xaxis=dict(title="Data", type="date"),
            yaxis=dict(title="Velocidade (mm/dia)"),
            margin=dict(t=30, b=40),
            legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(fig_c, use_container_width=True)
        pico = s_v.loc[s_v["vel"].idxmax()]
        st.caption(f"Velocidade máxima do alvo {alvo_sel}: {pico['vel']:.2f} "
                   f"mm/dia, no intervalo terminado a "
                   f"{pico[COLS['data']].strftime('%d/%m/%Y')}.")

    # ---- 4. PROFUNDIDADE ----
    st.divider()
    st.markdown("#### 4. A que profundidade — perfil inclinométrico")
    st.caption("Profundidade do deslocamento máximo de cada inclinómetro e "
               "evolução da deformada ao longo do tempo.")
    res = dados.get("resumo")
    if res is not None and not res.empty:
        res = res.copy()
        res[COLS["data"]] = pd.to_datetime(res[COLS["data"]], errors="coerce")
        linhas = []
        for inc in sorted(res[COLS["inclinometro"]].dropna().unique()):
            si = res[res[COLS["inclinometro"]] == inc].sort_values(
                COLS["data"]).iloc[-1]
            linhas.append({
                "Inclinómetro": inc,
                "Desl. máx (mm)": round(float(si[COLS["desl_max_global"]]), 1),
                "Profundidade do máx (m)": round(float(si[COLS["prof_do_max"]]), 1),
            })
        df_d = pd.DataFrame(linhas)
        st.dataframe(df_d, use_container_width=True, hide_index=True)
    else:
        st.info("Sem dados de inclinómetros para o perfil de profundidade.")

    # deformada do inclinometro ao longo do tempo (ilustracao do resultado)
    perfis = dados.get("perfis")
    if perfis is not None and not perfis.empty:
        perfis = perfis.copy()
        perfis[COLS["data"]] = pd.to_datetime(perfis[COLS["data"]], errors="coerce")
        incs = sorted(perfis[COLS["inclinometro"]].dropna().unique())
        inc_sel = st.selectbox("Inclinómetro (deformada no tempo)", incs,
                               key="res_inc_deformada")
        pi = perfis[perfis[COLS["inclinometro"]] == inc_sel]
        datas_i = sorted(pi[COLS["data"]].dropna().unique())
        if len(datas_i) > 6:
            idx = np.linspace(0, len(datas_i) - 1, 6).astype(int)
            datas_plot = [datas_i[i] for i in idx]
        else:
            datas_plot = datas_i
        fig8 = go.Figure()
        n = len(datas_plot)
        for k, d in enumerate(datas_plot):
            sp = pi[pi[COLS["data"]] == d].sort_values(COLS["profundidade"])
            frac = k / max(n - 1, 1)
            cor = f"rgb({int(150+105*frac)},{int(180-150*frac)},{int(200-150*frac)})"
            fig8.add_trace(go.Scatter(
                x=[float(v) for v in sp[COLS["desl_total"]].tolist()],
                y=[float(v) for v in sp[COLS["profundidade"]].tolist()],
                mode="lines+markers", name=pd.to_datetime(d).strftime("%d/%m/%Y"),
                line=dict(color=cor, width=2), marker=dict(size=3)))
        fig8.update_layout(
            height=560,
            xaxis=dict(title="Deslocamento acumulado (mm)"),
            yaxis=dict(title="Profundidade (m)", autorange="reversed"),
            margin=dict(t=30, b=40),
            legend=dict(title="Campanha", orientation="v", x=1.02, y=1))
        st.plotly_chart(fig8, use_container_width=True)
        st.caption(f"Deformada do {inc_sel} em várias campanhas (cor mais clara "
                   f"= mais antiga; mais escura = mais recente).")

    # ---- 5. OBSERVADO vs PREVISTO ----
    st.divider()
    st.markdown("#### 5. Observado vs. Previsto — comparação com o projeto")
    st.caption(
        "Deslocamento horizontal máximo OBSERVADO (todas as campanhas) vs. "
        "ESTIMATIVA DO PROJETO para a última fase de escavação (memória "
        "descritiva: ~20 mm na cortina poente/norte e ~10 mm na nascente/sul). "
        "Nos alçados, a estimativa é a da zona do projeto em que cada um se "
        "situa (atribuição inferida da posição; menos segura nos alçados BF, "
        "CD e DE, no canto sudoeste). Nos edifícios vizinhos, a comparação é "
        "INDIRETA: a estimativa refere-se à cortina adjacente, não ao "
        "edifício. A coluna «Ultrapassada a» indica a primeira campanha em que "
        "o máximo do grupo atingiu a estimativa.")
    linhas = []
    for chave, grp in alvos.groupby(COLS["edificio"]):
        prev, zona, tipo = previsto_do_grupo(chave)
        if prev is None:
            continue
        obs = grp[COLS["desl_h"]].max()
        if pd.isna(obs):
            continue
        t_ult, _v = data_ultrapassagem(grp[COLS["data"]], grp[COLS["desl_h"]], prev)
        linhas.append({
            "Grupo": nome_curto_grupo(chave), "Zona": zona, "Comparação": tipo,
            "Previsto (mm)": prev, "Observado máx (mm)": round(float(obs), 1),
            "Observado/previsto": round(float(obs) / prev, 1),
            "Ultrapassada a": (t_ult.strftime("%d/%m/%Y") if t_ult is not None else "–"),
        })
    df5 = pd.DataFrame(linhas).sort_values("Observado máx (mm)", ascending=False)
    st.dataframe(df5, use_container_width=True, hide_index=True,
                 height=int(35 * (len(df5) + 1) + 3),
                 column_config={
                     _k: st.column_config.NumberColumn(format=_f)
                     for _c, _f in (("Previsto (mm)", "%.0f"),
                                    ("Observado máx (mm)", "%.1f"),
                                    ("Observado/previsto", "%.1f"))
                     for _k in {_c, _tr(_c)}} | {
                     _k: st.column_config.TextColumn(width="medium")
                     for _k in {"Grupo", _tr("Grupo")}})
    top5 = df5.head(12)
    fig5 = go.Figure()
    fig5.add_trace(go.Bar(
        y=top5["Grupo"], x=top5["Previsto (mm)"], orientation="h",
        name="Previsto (projeto)", marker=dict(color="#9aa7b5")))
    fig5.add_trace(go.Bar(
        y=top5["Grupo"], x=top5["Observado máx (mm)"], orientation="h",
        name="Observado (máx)", marker=dict(color="#c0140f")))
    fig5.add_vline(x=40, line=dict(color="#8B0000", width=1, dash="dash"),
                   annotation_text="Alarme contenção (40)",
                   annotation_position="top")
    fig5.update_layout(
        height=460, barmode="group",
        xaxis=dict(title="Desl. horizontal (mm)"),
        yaxis=dict(autorange="reversed"),
        margin=dict(l=10, r=10, t=40, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig5, use_container_width=True)
    _dir = df5[df5["Comparação"].str.startswith("direta")]
    if not _dir.empty:
        pior = _dir.iloc[0]
        st.caption(
            f"Maior excedência na contenção: {pior['Grupo']} — "
            f"{pior['Observado máx (mm)']:.1f} mm, {pior['Observado/previsto']:.1f} "
            f"vezes a estimativa de {pior['Previsto (mm)']:.0f} mm "
            f"(zona {pior['Zona']}), ultrapassada a {pior['Ultrapassada a']}.")

    # ---- 6. ENVELOPES EMPIRICOS -----------------------------------------
    st.divider()
    st.markdown("#### 6. Envelopes empíricos — δ_h,máx em função de H")
    st.caption(
        "Deslocamento horizontal máximo de cada alçado (máximo dos prismas em "
        "todas as campanhas) e de cada inclinómetro (máximo do perfil), em "
        "função da profundidade escavada H (do topo das estacas ao fundo de "
        "escavação), face aos referenciais empíricos e às estimativas do "
        "projeto (memória descritiva, por zona do projeto). As linhas a "
        "tracejado são δ_h = p%·H; a sombra marca as faixas de Long e de Yoo.")

    import numpy as _np
    _pts_alc, _pts_inc = [], []
    _cont = dados["alvos"].copy()
    _cont["_alc"] = _cont[COLS["edificio"]].apply(_extrair_alcado)
    for _alc, _g in _cont[_cont["_alc"].notna()].groupby("_alc"):
        if _alc not in ALTURA_CORTINA:
            continue
        _dh = _g[COLS["desl_h"]].max()
        if pd.isna(_dh):
            continue
        _hmin, _hmax = ALTURA_CORTINA[_alc]
        _proj = previsto_do_grupo(f"Alçado {_alc}")[0]
        _pts_alc.append((_alc, (_hmin + _hmax) / 2, _hmin, _hmax, float(_dh), _proj))
    _res = dados.get("resumo")
    if _res is not None and not _res.empty:
        for _inc, _g in _res.groupby(COLS["inclinometro"]):
            if _inc not in H_INCLINOMETRO:
                continue
            _dh = _g[COLS["desl_max_global"]].max()
            if pd.isna(_dh):
                continue
            _hmin, _hmax = H_INCLINOMETRO[_inc]
            _pts_inc.append((_inc, (_hmin + _hmax) / 2, _hmin, _hmax, float(_dh)))

    if not _pts_alc and not _pts_inc:
        st.info("Sem dados suficientes para os envelopes empíricos.")
    else:
        _Hs = [p[1] for p in _pts_alc] + [p[1] for p in _pts_inc]
        _xx = _np.linspace(min(_Hs) - 1.0, max(_Hs) + 1.0, 50)
        figE = go.Figure()
        # posicao dos rotulos para nao se sobreporem (pontos proximos)
        _POS_ROT = {"AB": "middle left", "CD": "middle right", "BF": "bottom center",
                    "OP": "bottom center", "I2": "middle right"}
        for _rot, _pmin, _pmax, _rgb in REFERENCIAS_EMPIRICAS:
            if _pmin == _pmax:
                figE.add_trace(go.Scatter(
                    x=_xx, y=_pmax * _xx * 10.0, mode="lines",
                    line=dict(color=f"rgb({_rgb})", width=2, dash="dash"),
                    name=_rot))
            else:
                figE.add_trace(go.Scatter(
                    x=list(_xx) + list(_xx[::-1]),
                    y=list(_pmax * _xx * 10.0) + list((_pmin * _xx * 10.0)[::-1]),
                    fill="toself", fillcolor=f"rgba({_rgb},0.12)",
                    line=dict(color=f"rgba({_rgb},0.6)", width=1),
                    name=_rot, hoverinfo="skip"))
        if _pts_alc:
            figE.add_trace(go.Scatter(
                x=[p[1] for p in _pts_alc], y=[p[4] for p in _pts_alc],
                mode="markers+text", text=[p[0] for p in _pts_alc],
                textposition=[_POS_ROT.get(p[0], "top center") for p in _pts_alc], textfont=dict(size=10),
                marker=dict(size=11, color="#b45309", symbol="circle",
                            line=dict(width=1, color="white")),
                name="Alçados (prismas)",
                customdata=[[p[2], p[3], p[4] / p[1] / 10.0] for p in _pts_alc],
                hovertemplate=("Alçado %{text}<br>H = %{customdata[0]:.1f}"
                               "–%{customdata[1]:.1f} m<br>δ_h = %{y:.1f} mm"
                               "<br>δ_h/H = %{customdata[2]:.2f} %<extra></extra>")))
        if _pts_inc:
            figE.add_trace(go.Scatter(
                x=[p[1] for p in _pts_inc], y=[p[4] for p in _pts_inc],
                mode="markers+text", text=[p[0] for p in _pts_inc],
                textposition=[_POS_ROT.get(p[0], "top center") for p in _pts_inc], textfont=dict(size=10),
                marker=dict(size=12, color="#1f2a44", symbol="diamond",
                            line=dict(width=1, color="white")),
                name="Inclinómetros",
                customdata=[[p[4] / p[1] / 10.0] for p in _pts_inc],
                hovertemplate=("%{text}<br>H = %{x:.1f} m<br>δ_h = %{y:.1f} mm"
                               "<br>δ_h/H = %{customdata[0]:.2f} %<extra></extra>")))
        _projx = [p[1] for p in _pts_alc if p[5] is not None]
        _projy = [p[5] for p in _pts_alc if p[5] is not None]
        if _projx:
            figE.add_trace(go.Scatter(
                x=_projx, y=_projy, mode="markers",
                name="Estimativa do projeto (memória descritiva)",
                marker=dict(size=9, color="#888", symbol="x",
                            line=dict(width=1, color="#888"))))
        figE.update_layout(
            height=520, xaxis=dict(title="Profundidade escavada, H (m)"),
            yaxis=dict(title="δ_h,máx (mm)", rangemode="tozero"),
            margin=dict(l=10, r=10, t=30, b=40),
            legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(figE, use_container_width=True)
        _acima = [p[0] for p in _pts_alc if p[4] / p[1] / 10.0 > 0.25] + \
                 [p[0] for p in _pts_inc if p[4] / p[1] / 10.0 > 0.25]
        _txt = ("Acima do limite superior de Long (0,25 % de H): "
                + ", ".join(_acima) + ". ") if _acima else ""
        st.caption(
            _txt + "Os alçados mais deslocados situam-se em torno de 0,2 % de H "
            "(Clough & O'Rourke), acima das estimativas do projeto; os troços "
            "curtos travados por escoras de canto (FG, BF) e os alçados DE e LM, "
            "com prismas zerados só em 22/12/2025 e 05/02/2026, ficam abaixo de "
            "0,1 %. Valores subestimados: excluem o movimento anterior à "
            "zeragem de cada instrumento.")

    # ---- 7. ELEMENTOS ADJACENTES ----
    st.divider()
    _bloco_elementos_adjacentes(alvos)

    # ---- 8 e 9. DEFORMADA DA CONTENÇÃO ----
    st.divider()
    _G = _geometria_contencao()
    if _G is None:
        st.markdown("#### 8. Deformada da contenção")
        st.info(f"Falta o ficheiro de geometria da contenção ({GEOMETRIA_CONTENCAO}) ao lado da aplicação.")
    else:
        _E = _estagios_deformada(alvos, float(_G.get("desvio_cota_relatorios_projeto", 0.0)))
        _bloco_deformada_planta(_E, _G)
        st.divider()
        _bloco_deformada_desenvolvimento(_E, dados.get("perfis"), _G)


# =========================================================================
# DEFORMADA DA CONTENÇÃO (blocos 8 e 9 de «Resultados»)
# -------------------------------------------------------------------------
# A geometria da contenção (contorno, vértices, estacas, inclinómetros e
# topo das estacas por troço) vem da planta do projeto, num ficheiro
# auxiliar ao lado da aplicação. As posições dos prismas passam do
# referencial dos alvos para o dos desenhos por uma transformação de
# semelhança guardada no mesmo ficheiro.
# =========================================================================
GEOMETRIA_CONTENCAO = "geometria_contencao.json"
# fim de cada etapa de escavação, segundo o plano impactado (PERIODOS_ESCAVACAO)
ETAPAS_DEFORMADA = [
    (PERIODOS_ESCAVACAO[1][3], "fim dos pisos 2 e 1", "end of levels 2 and 1"),
    (PERIODOS_ESCAVACAO[2][3], "fim do piso -1", "end of level -1"),
    (PERIODOS_ESCAVACAO[3][3], "fim do piso -2", "end of level -2"),
    (PERIODOS_ESCAVACAO[4][3], "fim da escavação", "end of excavation"),
]
CORES_ETAPAS = ["#b9d6ec", "#7cb0d8", "#3f86c0", "#1b5ea8", "#0a2f63"]


@st.cache_data(show_spinner=False)
def _geometria_contencao():
    import json
    for fp in (Path(__file__).resolve().parent / GEOMETRIA_CONTENCAO, Path(GEOMETRIA_CONTENCAO)):
        if fp.exists():
            G = json.loads(fp.read_text(encoding="utf-8"))
            poly = np.asarray(G["contorno"], float)
            seg = np.diff(poly, axis=0)
            sl = np.hypot(seg[:, 0], seg[:, 1])
            ch = np.r_[0, np.cumsum(sl)]
            vch = {k: float(ch[int(np.argmin(np.hypot(*(poly - np.asarray(v)).T)))])
                   for k, v in G["vertices"].items()}
            vch[G["desenvolvimento"]["inicio"] + "2"] = float(ch[-1])
            G["_poly"] = poly
            G["_vch"] = vch
            return G
    return None


def _alvos_para_desenho(G, MP, vetor=False):
    """Coordenadas (M, P) dos alvos, em m, no referencial dos desenhos; com
    vetor=True, transforma um deslocamento (só escala e rotação)."""
    T = G["transformacao_alvos_para_desenho"]
    R = np.asarray(T["R"], float)
    MP = np.atleast_2d(np.asarray(MP, float))
    if vetor:
        return float(T["s"]) * MP @ R.T
    return (float(T["s"]) * (MP - np.asarray(T["origem_alvos"], float)) @ R.T
            + np.asarray(T["origem_desenho"], float))


def _desenvolvimento_contorno(G, xy):
    """Distância ao longo do contorno, a partir do canto inicial, da projeção
    de cada ponto (referencial dos desenhos) no contorno."""
    poly = G["_poly"]
    seg = np.diff(poly, axis=0)
    sl = np.hypot(seg[:, 0], seg[:, 1])
    ch = np.r_[0, np.cumsum(sl)]
    sl2 = np.where(sl > 0, sl ** 2, 1.0)
    out = []
    for p in np.atleast_2d(xy):
        t = np.clip(((p - poly[:-1]) * seg).sum(1) / sl2, 0, 1)
        q = poly[:-1] + t[:, None] * seg
        i = int(np.argmin(np.hypot(*(p - q).T)))
        out.append(ch[i] + t[i] * sl[i])
    return np.array(out)


def _estagios_deformada(alvos, desvio_cota):
    """Deslocamento de cada prisma da contenção no fim de cada etapa de
    escavação (interpolado entre campanhas) e na última leitura, se esta for
    posterior ao fim da escavação. Um prisma só entra nas etapas posteriores
    à sua zeragem."""
    import re as _re
    a = alvos[alvos[COLS["edificio"]].astype(str).str.contains("Alçado", na=False)]
    a = a.dropna(subset=[COLS["desl_h"]]).sort_values(COLS["data"])
    t0 = pd.Timestamp("2025-01-01")
    fim = pd.Timestamp(ETAPAS_DEFORMADA[-1][0])
    linhas = []
    for alvo, g in a.groupby(COLS["alvo"]):
        t = (g[COLS["data"]] - t0).dt.days.values.astype(float)
        h = g[COLS["desl_h"]].values.astype(float)
        dm = g[COLS["dM"]].values.astype(float)
        dp = g[COLS["dP"]].values.astype(float)
        base = dict(alvo=str(alvo), alc=_nome_elemento(g[COLS["edificio"]].iloc[0]).replace("Alçado ", ""),
                    rep=bool(_re.search(r"\d[b-d]$", str(alvo))),
                    m0=float(g[COLS["M0"]].iloc[0]), p0=float(g[COLS["P0"]].iloc[0]),
                    cota=float(g[COLS["Z0"]].iloc[0]) - desvio_cota)
        for k, et in enumerate(ETAPAS_DEFORMADA + [None]):
            if et is None:
                if g[COLS["data"]].iloc[-1] <= fim:
                    continue
                v, vm, vp, dd = h[-1], dm[-1], dp[-1], g[COLS["data"]].iloc[-1]
            else:
                dd = pd.Timestamp(et[0])
                x = float((dd - t0).days)
                if x < t[0] or x > t[-1]:
                    continue
                v, vm, vp = np.interp(x, t, h), np.interp(x, t, dm), np.interp(x, t, dp)
            linhas.append(dict(base, k=k, h=float(v), dm=float(vm), dp=float(vp), data=dd))
    return pd.DataFrame(linhas)


def _rotulos_etapas(E, en):
    rot = [f"{pd.Timestamp(d):%d/%m/%Y} – {e if en else p}" for d, p, e in ETAPAS_DEFORMADA]
    ult = E[E["k"] == len(ETAPAS_DEFORMADA)]["data"]
    fim = f" ({'to' if en else 'até'} {ult.max():%d/%m/%Y})" if len(ult) else ""
    rot.append(("last reading, after the end of excavation" if en
                else "última leitura, após o fim da escavação") + fim)
    return rot


def _bloco_deformada_planta(E, G):
    """Bloco 8 (Figura 5.9): trajetória em planta do topo de cada coluna de
    prismas no fim de cada etapa, com o deslocamento amplificado."""
    en = _lang_en()
    st.markdown("#### 8. Deformed shape of the retaining wall in plan" if en
                else "#### 8. Deformada da contenção em planta")
    st.caption(
        "Position of the top of each column of prisms (the highest original prism at each "
        "position) at the end of each excavation stage, according to the impacted works "
        "schedule, and at the last reading after the end of excavation, with the horizontal "
        "displacement magnified. Values at the stage dates are interpolated between campaigns; "
        "a prism only enters the stages after its zeroing. The geometry (outline, piles, "
        "corners and inclinometers) comes from the design plan, in an auxiliary file." if en else
        "Posição do topo de cada coluna de prismas (o prisma original mais alto em cada "
        "posição) no fim de cada etapa de escavação, segundo o plano impactado, e na última "
        "leitura após o fim da escavação, com o deslocamento horizontal amplificado. Os "
        "valores nas datas das etapas são interpolados entre campanhas; um prisma só entra "
        "nas etapas posteriores à sua zeragem. A geometria (contorno, estacas, vértices e "
        "inclinómetros) vem da planta do projeto, num ficheiro auxiliar.")
    amp = st.slider("Displacement magnification" if en else "Amplificação do deslocamento",
                    min_value=50, max_value=500, value=200, step=50, key="def_amp")
    S = E[~E["rep"]] if len(E) else E
    if S.empty:
        st.info("No wall prism readings." if en else "Sem leituras de prismas da contenção.")
        return
    # topo de cada coluna: prisma original de maior cota em cada posição (a menos de 1,5 m)
    u = S.drop_duplicates("alvo")[["alvo", "alc", "m0", "p0", "cota"]]
    topos = []
    for _, g in u.groupby("alc"):
        usados = []
        for _, r in g.sort_values("cota", ascending=False).iterrows():
            if any(np.hypot(r.m0 - q.m0, r.p0 - q.p0) < 1.5 for q in usados):
                continue
            usados.append(r)
            topos.append(r.alvo)
    T = S[S["alvo"].isin(topos)].copy()
    xy = _alvos_para_desenho(G, T[["m0", "p0"]].values)
    vec = _alvos_para_desenho(G, T[["dm", "dp"]].values / 1000.0, vetor=True)
    T["x"], T["y"] = xy[:, 0], xy[:, 1]
    T["xd"], T["yd"] = T["x"] + amp * vec[:, 0], T["y"] + amp * vec[:, 1]
    rot = _rotulos_etapas(E, en)

    fig = go.Figure()
    pl = np.asarray(G["estacas"], float)
    fig.add_trace(go.Scatter(x=pl[:, 0], y=pl[:, 1], mode="markers", hoverinfo="skip",
                             marker=dict(size=4, symbol="circle-open", color="#c4c4c4",
                                         line=dict(width=0.7)), showlegend=False))
    V = G["vertices"]
    vk = [k for k in "ABCDEFGHIJKLMNOPQ" if k in V]
    fig.add_trace(go.Scatter(x=[V[k][0] for k in vk], y=[V[k][1] for k in vk], mode="markers+text",
                             text=vk, textposition="middle center",
                             textfont=dict(size=10, color="#b03a2e"), hoverinfo="skip",
                             marker=dict(size=17, color="white", line=dict(color="#b03a2e", width=1)),
                             showlegend=False))
    xs, ys = [], []
    for _, g in T.groupby("alvo"):
        g = g.sort_values("k")
        xs += [g["x"].iloc[0]] + list(g["xd"]) + [None]
        ys += [g["y"].iloc[0]] + list(g["yd"]) + [None]
    fig.add_trace(go.Scatter(x=xs, y=ys, mode="lines", line=dict(color="#9aa7b4", width=1),
                             hoverinfo="skip", showlegend=False))
    for k in range(len(rot)):
        g = T[T["k"] == k]
        if g.empty:
            continue
        fig.add_trace(go.Scatter(
            x=g["xd"], y=g["yd"], mode="markers", name=rot[k],
            marker=dict(size=10 if k < len(rot) - 1 else 12, color=CORES_ETAPAS[k % 5],
                        line=dict(color="#333333", width=0.6)),
            customdata=np.c_[g["alvo"], g["h"].round(1)],
            hovertemplate="%{customdata[0]}<br>%{customdata[1]} mm<extra>" + rot[k] + "</extra>"))
    g0 = T.drop_duplicates("alvo")
    fig.add_trace(go.Scatter(x=g0["x"], y=g0["y"], mode="markers", marker=dict(size=6, color="#555555"),
                             name="initial position of the prism" if en else "posição inicial do prisma",
                             text=g0["alvo"], hovertemplate="%{text}<extra></extra>"))
    inc = G["inclinometros"]
    fig.add_trace(go.Scatter(x=[v["xy"][0] for v in inc.values()], y=[v["xy"][1] for v in inc.values()],
                             mode="markers+text", text=list(inc),
                             textposition=[v.get("rotulo_planta", "bottom right") for v in inc.values()],
                             textfont=dict(size=13, color="#2166ac"), hoverinfo="skip",
                             marker=dict(size=11, symbol="square", color="#2166ac",
                                         line=dict(color="white", width=0.6)),
                             name="inclinometer" if en else "inclinómetro"))
    for a, (lx, ly) in G["rotulos_alcados"].items():
        fig.add_annotation(x=lx, y=ly, text=f"<b>{a}</b>", showarrow=False,
                           font=dict(size=13, color="#1f4e79"), bgcolor="rgba(255,255,255,0.85)")
    x0, x1, y0, y1 = G["janela_planta"]
    # escala gráfica (0-20 m) e escala do deslocamento amplificado
    bx, by = x0 + 6, y1 - 4.8
    for i_ in range(4):
        fig.add_shape(type="rect", x0=bx + 5 * i_, x1=bx + 5 * (i_ + 1), y0=by, y1=by + 0.8,
                      fillcolor="black" if i_ % 2 == 0 else "white", line=dict(color="black", width=0.8))
    fig.add_annotation(x=bx, y=by - 0.4, text="0", showarrow=False, yanchor="top", font=dict(size=13, color="black"))
    fig.add_annotation(x=bx + 20, y=by - 0.4, text="20 m", showarrow=False, yanchor="top", font=dict(size=13, color="black"))
    fig.add_shape(type="line", x0=bx, x1=bx + amp * 0.020, y0=by - 5.6, y1=by - 5.6,
                  line=dict(color="#0a2f63", width=3))
    fig.add_annotation(x=bx, y=by - 7.4, xanchor="left", showarrow=False, font=dict(size=13, color="black"),
                       text=(f"20 mm of displacement (magnification {amp}×)" if en
                             else f"20 mm de deslocamento (amplificação de {amp}×)"))
    b = float(G["norte_rad"])
    nx, ny = x1 - 6, y1 - 12
    fig.add_annotation(x=nx + 2.6 * np.sin(b), y=ny + 4.4 * np.cos(b), ax=nx, ay=ny, axref="x", ayref="y",
                       xref="x", yref="y", text="", showarrow=True, arrowhead=2, arrowsize=1.3,
                       arrowwidth=2, arrowcolor="black")
    fig.add_annotation(x=nx + 3.2 * np.sin(b), y=ny + 6.6, text="<b>N</b>", showarrow=False,
                       font=dict(size=15, color="black"))
    fig.update_xaxes(range=[x0, x1], visible=False, constrain="domain", autorange=False)
    fig.update_yaxes(range=[y0, y1], visible=False, scaleanchor="x", scaleratio=1,
                     constrain="domain", autorange=False)
    fig.update_layout(
        height=820, margin=dict(l=5, r=5, t=5, b=5), plot_bgcolor="white",
        legend=dict(x=0.995, y=0.005, xanchor="right", yanchor="bottom", font=dict(size=13),
                    bgcolor="rgba(255,255,255,0.95)", bordercolor="#bbbbbb", borderwidth=1,
                    title=dict(text=("Magnified position of the top of each column of prisms" if en
                                     else "Posição amplificada do topo de cada coluna de prismas"),
                               font=dict(size=13))))
    st.plotly_chart(fig, use_container_width=True)


def _bloco_deformada_desenvolvimento(E, perfis, G):
    """Bloco 9 (Figura 5.10): deslocamento de cada prisma à sua cota e perfil
    dos inclinómetros, ao longo do contorno da escavação, por etapa."""
    from plotly.subplots import make_subplots
    en = _lang_en()
    st.markdown("#### 9. Deformed shape of the retaining wall along its outline" if en
                else "#### 9. Deformada da contenção em desenvolvimento")
    st.caption(
        "Cumulative horizontal displacement of each prism with a reading, at the elevation at "
        "which it is installed, and profile of the inclinometers, along the excavation outline "
        "from corner B, clockwise, at the end of each excavation stage and at the last reading. "
        "Each prism is shown from its zeroing; elevations in the design datum (report datum "
        "minus the offset in the geometry file); inclinometer collars assumed at the top of the "
        "piles. The profile of each inclinometer is the last valid reading up to the stage date." if en else
        "Deslocamento horizontal acumulado de cada prisma com leitura, à cota a que está "
        "instalado, e perfil dos inclinómetros, ao longo do contorno da escavação a partir do "
        "canto B, no sentido dos ponteiros do relógio, no fim de cada etapa de escavação e na "
        "última leitura. Cada prisma é representado desde a sua zeragem; cotas no referencial "
        "do projeto (referencial dos relatórios menos o desvio indicado no ficheiro de "
        "geometria); boca dos inclinómetros admitida no topo das estacas. O perfil de cada "
        "inclinómetro é a última leitura válida até à data da etapa.")
    if E.empty:
        st.info("No wall prism readings." if en else "Sem leituras de prismas da contenção.")
        return
    D = G["desenvolvimento"]
    vch = G["_vch"]
    fim = D["inicio"] + "2"
    L = vch[fim]
    u = E.drop_duplicates("alvo")[["alvo", "m0", "p0"]].copy()
    u["s"] = _desenvolvimento_contorno(G, _alvos_para_desenho(G, u[["m0", "p0"]].values))
    S = E.merge(u[["alvo", "s"]], on="alvo")
    # inclinómetros: leituras válidas
    I = perfis.copy() if perfis is not None and not perfis.empty else pd.DataFrame()
    if not I.empty:
        I[COLS["data"]] = pd.to_datetime(I[COLS["data"]], errors="coerce")
        if "Validade / observação" in I.columns:
            I = I[~I["Validade / observação"].fillna("").astype(str).str.startswith("Excluir")]
    incl = G["inclinometros"]
    ipos = {nm: float(_desenvolvimento_contorno(G, [v["xy"]])[0]) for nm, v in incl.items()}
    rot = _rotulos_etapas(E, en)
    n = len(rot)
    VMAX = 60
    fig = make_subplots(rows=n, cols=1, shared_xaxes=True, vertical_spacing=0.05)
    niv = D["niveis"]
    cf = float(D["cota_fundo"])
    for k in range(n):
        r = k + 1
        ax = "" if r == 1 else str(r)
        # parede: troços até ao topo das estacas (formas por baixo dos traços, porque os perfis dos
        # inclinómetros são heatmaps, que ficam sempre por baixo dos scatter) e, a tracejado, muros e
        # consolas acima do topo das estacas, interrompidos onde passa a coluna de um inclinómetro
        xr, yr = f"x{ax}", f"y{ax}"
        hx, hy = [], []
        for a, b_, t, tm in D["topo_estacas"]:
            s0, s1 = vch[a], vch[b_]
            fig.add_shape(type="rect", x0=s0, x1=s1, y0=cf, y1=t, xref=xr, yref=yr, layer="below",
                          fillcolor="#f1eee8", line=dict(width=0))
            fig.add_shape(type="line", x0=s0, x1=s1, y0=t, y1=t, xref=xr, yref=yr, layer="below",
                          line=dict(color="#b9b2a6", width=1))
            if tm > t:
                tro = [(s0, s1)]
                for nm_, v_ in incl.items():
                    c0 = ipos[nm_] + float(v_.get("desvio_desenvolvimento", 0.0)) - 1.6
                    c1 = c0 + 3.2
                    tro = [p for q0, q1 in tro for p in ((q0, min(q1, c0)), (max(q0, c1), q1)) if p[1] - p[0] > 0.05]
                for q0, q1 in tro:
                    hx += [q0, q1, q1, q0, q0, None]; hy += [t, t, tm, tm, t, None]
        if hx:
            fig.add_trace(go.Scatter(x=hx, y=hy, fill="toself", mode="lines", line=dict(width=0),
                                     fillpattern=dict(shape="/", fgcolor="#cfc8bb", bgcolor="rgba(0,0,0,0)",
                                                      size=6, solidity=0.15),
                                     hoverinfo="skip", showlegend=False), row=r, col=1)
        for z, _ in niv:
            fig.add_shape(type="line", x0=-3, x1=L + 3, y0=z, y1=z, xref=xr, yref=yr, layer="below",
                          line=dict(color="#d4d0c8", width=0.8, dash="dash"))
        for v in D["vertices_eixo"]:
            fig.add_shape(type="line", x0=vch[v], x1=vch[v], y0=2.5, y1=29.5, xref=xr, yref=yr, layer="below",
                          line=dict(color="#c9c3b8", width=0.8))
        # inclinómetros: última leitura válida até à data da etapa
        if not I.empty:
            for nm, v in incl.items():
                gi = I[I[COLS["inclinometro"]] == nm]
                if gi.empty:
                    continue
                datas = gi[COLS["data"]].dropna()
                if k < len(ETAPAS_DEFORMADA):
                    datas = datas[datas <= pd.Timestamp(ETAPAS_DEFORMADA[k][0])]
                if datas.empty:
                    continue
                g = gi[gi[COLS["data"]] == datas.max()].sort_values(COLS["profundidade"])
                z = float(v["cota_boca"]) - g[COLS["profundidade"]].values.astype(float)
                val = g[COLS["desl_total"]].values.astype(float)
                o = np.argsort(z); z, val = z[o], val[o]
                if len(z) > 1:
                    mid = (z[:-1] + z[1:]) / 2
                    ed = np.r_[z[0] - (mid[0] - z[0]), mid, z[-1] + (z[-1] - mid[-1])]
                else:
                    ed = np.r_[z[0] - 0.25, z[0] + 0.25]
                sx = ipos[nm] + float(v.get("desvio_desenvolvimento", 0.0))
                fig.add_trace(go.Heatmap(x=[sx - 1.3, sx + 1.3], y=ed, z=val[:, None], coloraxis="coloraxis",
                                         hovertemplate=f"{nm} ({datas.max():%d/%m/%Y})<br>%{{z:.1f}} mm<extra></extra>"),
                              row=r, col=1)
                fig.add_trace(go.Scatter(x=[sx - 1.3, sx + 1.3, sx + 1.3, sx - 1.3, sx - 1.3],
                                         y=[ed[0], ed[0], ed[-1], ed[-1], ed[0]], mode="lines",
                                         line=dict(color="#555555", width=0.8), hoverinfo="skip",
                                         showlegend=False), row=r, col=1)
                cb = float(v["cota_boca"])
                fig.add_trace(go.Scatter(x=[ipos[nm], sx], y=[cb + 1.0, cb + 1.0], mode="lines",
                                         line=dict(color="#555555", width=0.8), hoverinfo="skip",
                                         showlegend=False), row=r, col=1)
                fig.add_annotation(x=sx, y=cb + 1.3, text=nm, showarrow=False, yanchor="bottom",
                                   font=dict(size=12, color="#333333"), row=r, col=1)
        # prismas: substitutos por baixo, originais por cima
        g = S[S["k"] == k].sort_values("rep", ascending=False)
        if len(g):
            fig.add_trace(go.Scatter(x=g["s"], y=g["cota"], mode="markers", showlegend=False,
                                     marker=dict(size=10, color=g["h"], coloraxis="coloraxis",
                                                 line=dict(color="#333333", width=0.6)),
                                     customdata=np.c_[g["alvo"], g["h"].round(1)],
                                     hovertemplate="%{customdata[0]}<br>%{customdata[1]} mm<extra></extra>"),
                          row=r, col=1)
            mx = g.loc[g["h"].idxmax()]
            info = (f"{len(g)} prisms with a reading · maximum {_num_pt(mx['h'])} mm ({mx['alvo']})" if en
                    else f"{len(g)} prismas com leitura · máximo {_num_pt(mx['h'])} mm ({mx['alvo']})")
            fig.add_annotation(x=0.997, y=1.02, xref=f"x{ax} domain", yref=f"y{ax} domain", text=info,
                               showarrow=False, xanchor="right", yanchor="bottom",
                               font=dict(size=13, color="#333333"))
        fig.add_annotation(x=0.003, y=1.02, xref=f"x{ax} domain", yref=f"y{ax} domain",
                           text=f"<b>({'abcdefgh'[k]}) {rot[k]}</b>", showarrow=False,
                           xanchor="left", yanchor="bottom", font=dict(size=14))
        fig.update_yaxes(range=[2.5, 29.5], tickvals=[z for z, _ in niv],
                         ticktext=[(nv.replace("piso", "level").replace("fundo", "bottom") if en else nv)
                                   for _, nv in niv],
                         showgrid=False, zeroline=False, ticks="", tickfont=dict(size=13), row=r, col=1)
    eixo = D["vertices_eixo"]
    fig.update_xaxes(range=[-3, L + 3], tickvals=[vch[v] for v in eixo] + [L],
                     ticktext=list(eixo) + [D["inicio"]], showgrid=False, zeroline=False,
                     ticks="outside", ticklen=3, tickfont=dict(size=13))
    fig.update_xaxes(title_text=("Development of the excavation outline from corner B, clockwise (m)" if en
                                 else "Desenvolvimento do contorno da escavação, a partir do canto B, "
                                      "no sentido dos ponteiros do relógio (m)"),
                     title_font=dict(size=14), row=n, col=1)
    # faixas das frentes e nomes dos alçados por cima do primeiro painel
    CF = ["#efe6d4", "#e3ebe0", "#e2e7f0"]
    for i_, (nm, a, b_) in enumerate(D["frentes"]):
        if en:
            nm = (nm.replace("frente ", "").replace("poente", "west").replace("nascente", "east")
                  .replace("norte", "north").replace("sul", "south") + " front")
        fig.add_shape(type="rect", x0=vch[a], x1=vch[b_], y0=1.43, y1=1.60, xref="x", yref="y domain",
                      fillcolor=CF[i_ % 3], line=dict(color="white", width=1.5))
        fig.add_annotation(x=(vch[a] + vch[b_]) / 2, y=1.515, xref="x", yref="y domain", text=nm,
                           showarrow=False, font=dict(size=13, color="#444444"))
    for nm, a, b_ in D["alcados"]:
        fig.add_annotation(x=(vch[a] + vch[b_]) / 2, y=1.33, xref="x", yref="y domain",
                           text=nm.replace("anel", "ring") if en else nm, showarrow=False,
                           font=dict(size=12, color="#555555"))
    fig.update_layout(
        height=1100, margin=dict(l=100, r=10, t=105, b=125), plot_bgcolor="white", font=dict(size=13),
        coloraxis=dict(colorscale=[[0, "#d6e6f4"], [0.25, "#9ecae1"], [0.5, "#4292c6"],
                                   [0.75, "#2171b5"], [1, "#08306b"]], cmin=0, cmax=VMAX,
                       colorbar=dict(orientation="h", x=0.5, xanchor="center", y=-0.075, yanchor="top",
                                     len=0.5, thickness=10, tickvals=list(range(0, VMAX + 1, 10)),
                                     title=dict(text=("Cumulative horizontal displacement, resultant (mm)" if en
                                                      else "Deslocamento horizontal acumulado, resultante (mm)"),
                                                side="top", font=dict(size=14)))))
    fig.add_annotation(x=-0.105, y=0.5, xref="paper", yref="paper", textangle=-90, showarrow=False,
                       text=("Level (design elevation)" if en else "Nível (cota de projeto)"),
                       font=dict(size=14))
    st.plotly_chart(fig, use_container_width=True)


def _bloco_elementos_adjacentes(alvos):
    """Bloco 7 de «Resultados» (Figura 5.8 da dissertação): dois alvos de
    elementos diferentes (p. ex. um edifício vizinho e o alçado que lhe fica em
    frente), cada um com os limiares do SEU critério, e o estado de cada um no
    tempo. Mostra se os critérios dos dois elementos são coerentes."""
    st.markdown("#### 7. Elementos adjacentes — cada um com o seu critério")
    st.caption(
        "Compara dois alvos de elementos diferentes, por exemplo um edifício "
        "vizinho e o alçado da contenção que lhe fica em frente. Cada série é "
        "lida com os limiares do critério do seu elemento, e a faixa inferior "
        "mostra o estado de cada alvo segundo esse critério. Opcionalmente, o "
        "primeiro alvo é também representado a partir da data de zeragem do "
        "segundo, para comparar movimentos com a mesma origem.")
    a = alvos.dropna(subset=[COLS["desl_h"]]).copy()
    a = a.sort_values(COLS["data"]).drop_duplicates(
        subset=[COLS["alvo"], COLS["data"]], keep="last")
    nomes = sorted(a[COLS["alvo"]].unique(), key=_chave_natural)
    if len(nomes) < 2:
        st.info("São precisos pelo menos dois alvos.")
        return
    i1 = nomes.index("A3") if "A3" in nomes else 0
    i2 = nomes.index("GH3") if "GH3" in nomes else 1
    c1, c2, c3 = st.columns([1, 1, 2])
    t1 = c1.selectbox("Alvo 1", nomes, index=i1, key="adj_t1")
    t2 = c2.selectbox("Alvo 2", nomes, index=i2, key="adj_t2")
    rebase = c3.checkbox("Representar também o alvo 1 a partir da zeragem do alvo 2",
                         value=True, key="adj_reb")
    g1 = a[a[COLS["alvo"]] == t1]; g2 = a[a[COLS["alvo"]] == t2]
    d0 = max(g1[COLS["data"]].min(), g2[COLS["data"]].min())
    d1 = min(g1[COLS["data"]].max(), g2[COLS["data"]].max())
    if d0 >= d1:
        st.info("Os dois alvos não têm leituras num período comum.")
        return
    per = st.slider("Período", min_value=d0.date(), max_value=d1.date(),
                    value=(d0.date(), d1.date()), format="DD/MM/YYYY", key="adj_per")
    p0, p1 = pd.Timestamp(per[0]), pd.Timestamp(per[1])
    el1 = g1[COLS["edificio"]].iloc[0]; el2 = g2[COLS["edificio"]].iloc[0]
    cr1, _r1, _ = criterios_do_alvo(el1); cr2, _r2, _ = criterios_do_alvo(el2)
    n1 = _nome_elemento(el1); n2 = _nome_elemento(el2)
    z1 = g1[COLS["data"]].min(); z2 = g2[COLS["data"]].min()
    en = _lang_en()
    _EST_EN = {"Regular": "regular", "Alerta": "alert", "Alarme": "alarm"}
    _zer = "zeroed on" if en else "zerado a"
    _ale, _ala = ("alert", "alarm") if en else ("alerta", "alarme")

    def _estados(g, cr):
        return [estado_calculado(h, v, cr) for h, v in zip(g[COLS["desl_h"]], g[COLS["dZ"]])]
    g1 = g1.assign(_e=_estados(g1, cr1)); g2 = g2.assign(_e=_estados(g2, cr2))
    w1 = g1[(g1[COLS["data"]] >= p0) & (g1[COLS["data"]] <= p1)]
    w2 = g2[(g2[COLS["data"]] >= p0) & (g2[COLS["data"]] <= p1)]

    from plotly.subplots import make_subplots
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, row_heights=[0.78, 0.22],
                        vertical_spacing=0.06)
    fig.add_trace(go.Scatter(x=w1[COLS["data"]], y=w1[COLS["desl_h"]], mode="lines+markers",
                             line=dict(color="#1f5fa8", width=2.2), marker=dict(size=5),
                             name=f"{t1}, {n1} ({_zer} {z1:%d/%m/%Y})"), row=1, col=1)
    if rebase and z2 > z1:
        _t0 = g1[COLS["data"]].min()
        base = float(np.interp((z2 - _t0).days, (g1[COLS["data"]] - _t0).dt.days.values,
                               g1[COLS["desl_h"]].values))
        wr = w1[w1[COLS["data"]] >= z2]
        fig.add_trace(go.Scatter(x=wr[COLS["data"]], y=wr[COLS["desl_h"]] - base, mode="lines",
                                 line=dict(color="#5b8fd0", width=1.5, dash="dash"),
                                 name=(f"{t1}, from {z2:%d/%m/%Y} (origin of {t2})" if en
                                       else f"{t1}, desde {z2:%d/%m/%Y} (origem do {t2})")), row=1, col=1)
    fig.add_trace(go.Scatter(x=w2[COLS["data"]], y=w2[COLS["desl_h"]], mode="lines+markers",
                             line=dict(color="#2b2b2b", width=2.2),
                             marker=dict(size=6, symbol="square"),
                             name=f"{t2}, {n2} ({_zer} {z2:%d/%m/%Y})"), row=1, col=1)
    lim = [(cr1[0], "#e67e00", "dash", f"{n1}: {_ale} {cr1[0]:g} mm"),
           (cr1[1], "#c0140f", "dash", f"{n1}: {_ala} {cr1[1]:g} mm"),
           (cr2[0], "#e67e00", "dot", f"{n2}: {_ale} {cr2[0]:g} mm"),
           (cr2[1], "#c0140f", "dot", f"{n2}: {_ala} {cr2[1]:g} mm")]
    _ys = sorted(set(round(v, 3) for v in (cr1[0], cr1[1], cr2[0], cr2[1])))
    for y, cor, dash, txt in lim:
        _pos = "top right" if y == max(_ys) or (y > min(_ys) and _ys.index(round(y, 3)) % 2 == 1
                                                 and y != cr1[1]) else "bottom right"
        fig.add_hline(y=y, line=dict(color=cor, width=1.3, dash=dash), row=1, col=1,
                      annotation_text=txt, annotation_position=_pos,
                      annotation_font=dict(size=11, color=cor))

    def _primeira(g, est):
        x = g[g["_e"].isin(est)]
        return (x[COLS["data"]].iloc[0], x[COLS["desl_h"]].iloc[0]) if len(x) else (None, None)
    da1, ha1 = _primeira(g1, ["Alarme"]); dl2, hl2 = _primeira(g2, ["Alerta", "Alarme"])
    for dd, hh, cor, txt in ((da1, ha1, "#c0140f", (f"{t1} alarm" if en else f"alarme do {t1}")), (dl2, hl2, "#e67e00", (f"{t2} alert" if en else f"alerta do {t2}"))):
        if dd is not None and p0 <= dd <= p1:
            fig.add_vline(x=dd, line=dict(color="#7a7a7a", width=1.2))
            fig.add_annotation(x=dd, y=1.0, xref="x", yref="y domain", text=f"{dd:%d/%m/%Y}<br>{txt}<br>({_num_pt(hh)} mm)",
                               showarrow=False, xanchor="left", yanchor="top", xshift=4,
                               font=dict(size=11, color=cor), bgcolor="rgba(255,255,255,0.8)")
    # faixa de estados
    CE = {"Regular": "#1f9e55", "Alerta": "#e67e00", "Alarme": "#c0140f"}
    for yy, g, nome in ((1, w1, t1), (0, w2, t2)):
        g = g.reset_index(drop=True)
        for k in range(len(g)):
            xa = g[COLS["data"]].iloc[k]
            xb = g[COLS["data"]].iloc[k + 1] if k + 1 < len(g) else p1
            fig.add_shape(type="rect", x0=xa, x1=xb, y0=yy - 0.38, y1=yy + 0.38,
                          fillcolor=CE.get(g["_e"].iloc[k], "#cccccc"), line=dict(width=0),
                          row=2, col=1)
    # traço invisível com datas no eixo inferior, para as datas em português
    fig.add_trace(go.Scatter(x=[p0, p1], y=[-5, -5], mode="markers", marker=dict(opacity=0),
                             hoverinfo="skip", showlegend=False), row=2, col=1)
    for est, cor in CE.items():
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
                                 marker=dict(size=10, symbol="square", color=cor),
                                 name=est, showlegend=True), row=2, col=1)
    fig.update_yaxes(title_text="Deslocamento horizontal (mm)", rangemode="tozero", row=1, col=1)
    fig.update_yaxes(tickvals=[1, 0], ticktext=[f"{t1} ({cr1[0]:g}/{cr1[1]:g} mm)",
                                                f"{t2} ({cr2[0]:g}/{cr2[1]:g} mm)"],
                     range=[-0.6, 1.6], row=2, col=1)
    fig.update_xaxes(type="date", range=[p0 - pd.Timedelta(days=2), p1 + pd.Timedelta(days=2)])
    fig.update_layout(height=600, margin=dict(l=10, r=10, t=30, b=10),
                      legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0))
    st.plotly_chart(fig, use_container_width=True)

    # frase de leitura (texto em inglês escrito à parte, por ter números e nomes variáveis)
    partes = []
    if da1 is not None:
        x2 = g2[g2[COLS["data"]] <= da1]
        if len(x2):
            h2 = x2[COLS["desl_h"]].iloc[-1]; e2 = x2["_e"].iloc[-1]
            partes.append(
                f"When {t1} entered alarm ({da1:%d/%m/%Y}, {_num_pt(ha1)} mm), {t2} read "
                f"{_num_pt(h2)} mm, in {_EST_EN.get(e2, e2.lower())} state." if en else
                f"Quando o {t1} entrou em alarme ({da1:%d/%m/%Y}, {_num_pt(ha1)} mm), "
                f"o {t2} registava {_num_pt(h2)} mm, em estado {e2.lower()}.")
        if dl2 is not None and dl2 > da1:
            partes.append(
                f"{t2} reached the alert on {dl2:%d/%m/%Y} ({_num_pt(hl2)} mm), "
                f"{(dl2 - da1).days // 7} weeks later." if en else
                f"O {t2} atingiu o alerta a {dl2:%d/%m/%Y} ({_num_pt(hl2)} mm), "
                f"{(dl2 - da1).days // 7} semanas depois.")
    for (na, ca), (nb, cb) in (((n2, cr2), (n1, cr1)), ((n1, cr1), (n2, cr2))):
        if ca[0] > cb[1]:
            partes.append(
                f"The alert threshold of {na} ({ca[0]:g} mm) is higher than the alarm "
                f"threshold of {nb} ({cb[1]:g} mm)." if en else
                f"O limiar de alerta do {na} ({ca[0]:g} mm) é superior ao limiar "
                f"de alarme do {nb} ({cb[1]:g} mm).")
            break
    if partes:
        st.caption(" ".join(partes))
    # (traduções das frases: ver _I18N_EN)


def zona_escav_de_edificio(edif):
    """Frente de escavação (poente_sul / norte / nascente_sul) do edifício ou
    elemento de um alvo; None se não mapeavel. Edifícios vizinhos por nome
    (confirmado com a foto aerea); alçados de contenção pela alinhamento das
    vigas de distribuição do plano (ALCADO_ZONA)."""
    if not isinstance(edif, str):
        return None
    e = edif.lower()
    if "edifício 1" in e:
        return "poente_sul"
    if "edifício 2" in e:
        return "nascente_sul"
    if "edifício 3" in e:
        return "norte"
    alc = _extrair_alcado(edif)
    if alc:
        return ALCADO_ZONA.get(alc)
    return None


def _marcos_escavacao_zona(fig, zona, dt_min, dt_max, so_escavacao=True):
    """Desenha, na figura de eixo temporal, os marcos reais de escavação da
    ZONA dada (ESCAVACAO_ZONAS). Devolve a lista (k, rotulo, data, cota, tipo)
    dos eventos dentro da janela. Se zona for None, usa a lista global
    (ESCAVACAO_COTAS) como recurso."""
    vis = []
    if zona and zona in ESCAVACAO_ZONAS:
        eventos = [(pd.to_datetime(d), tp, cota, rot)
                   for d, tp, cota, rot in ESCAVACAO_ZONAS[zona]
                   if (not so_escavacao or tp == "escavacao")
                   and dt_min <= pd.to_datetime(d) <= dt_max]
    else:
        eventos = [(pd.to_datetime(fim), "escavacao", cota, rot)
                   for rot, cota, ini, fim in ESCAVACAO_COTAS
                   if dt_min <= pd.to_datetime(fim) <= dt_max]
    eventos.sort()
    for k, (t, tp, cota, rot) in enumerate(eventos, start=1):
        cor = "#8B4513" if tp == "escavacao" else (
            "#7a1fa0" if tp == "ancoragem" else (
                "#1f6fb4" if tp == "escora" else "#1f9e55"))
        fig.add_shape(type="line", xref="x", yref="paper", x0=t, x1=t,
                      y0=0, y1=1, line=dict(color=cor, width=1, dash="dash"),
                      layer="below")
        fig.add_annotation(
            x=t, y=-0.02, yref="paper", text=f"E{k}", showarrow=False,
            xanchor="center", yanchor="top", font=dict(size=9, color="white"),
            bgcolor=cor, borderpad=2,
            hovertext=f"{rot} — {t.strftime('%d/%m/%Y')}")
        vis.append((k, rot, t, cota, tp))
    return vis


def separador_correlacoes(dados):
    """
    CORRELAÇÕES — gráficos de interação entre instrumentos. Cruzam, num eixo de
    tempo comum (ou num diagrama XY), várias séries da instrumentação para
    ILUSTRAR a correlação entre os dados introduzidos (água, escavação,
    faseamento, carga nas ancoragens) e as movimentações medidas. O objetivo e
    DEMONSTRAR a correlação — não explicar a causa, que e do autor no texto.
    """
    import numpy as np
    st.subheader("Correlações — interação entre instrumentos")
    st.caption("Gráficos que cruzam várias séries da instrumentação para "
               "mostrar como os movimentos acentuados acompanham os outros "
               "dados (água, escavação, faseamento, carga nas ancoragens). "
               "Demonstram a correlação entre os dados; a interpretação das "
               "causas e do autor.")

    # =====================================================================
    # 1. DEFORMACAO x AGUA x ESCAVACAO x FASEAMENTO  (eixo de tempo comum)
    # =====================================================================
    st.markdown("#### 1. Deformação × Água × Escavação × Faseamento")
    st.caption("Num só eixo de tempo: a deformação medida (esquerda), a cota da "
               "água subterrânea (direita), os marcos de escavação (plano impactado) DA "
               "ZONA do alvo/inclinómetro escolhido (E1-E..., no fundo) e as "
               "fases da obra (faixas de cor). Cada instrumento e cruzado com a "
               "escavação da frente que lhe fica em frente — não com a do lado "
               "oposto.")

    c1, c2 = st.columns(2)
    with c1:
        tipo = st.radio("Série de deformação", ["Inclinómetro", "Alvo"],
                        horizontal=True, key="corr_tipo")
        st.checkbox("Marcos de escavação da zona (datas do planeamento)", value=True,
                    key="corr_escav",
                    help="Linhas verticais nas datas do plano de trabalhos impactado em que a escavação "
                         "atingiu cada cota NA FRENTE do instrumento escolhido "
                         "(plano de trabalhos impactado, por zona de estacas).")
    # --- serie de deformacao escolhida (e a sua zona de escavacao) ---
    df_def = None
    lbl_def = ""
    zona_sel = None
    with c2:
        if tipo == "Inclinómetro":
            res = dados.get("resumo")
            if res is not None and COLS["data"] in res.columns:
                res = res.copy()
                res[COLS["data"]] = pd.to_datetime(res[COLS["data"]], errors="coerce")
                col_inc = "Inclinómetro"
                incs = sorted(res[col_inc].dropna().unique())
                sel = st.selectbox("Inclinómetro", incs, key="corr_inc")
                s = res[res[col_inc] == sel].sort_values(COLS["data"])
                df_def = s[[COLS["data"], "Máx. desloc. acumulado total (mm)"]].rename(
                    columns={"Máx. desloc. acumulado total (mm)": "def"})
                lbl_def = f"Desl. máx. {sel} (mm)"
                zona_sel = INC_ZONA.get(str(sel))
        else:
            alv = dados["alvos"].copy()
            alv[COLS["data"]] = pd.to_datetime(alv[COLS["data"]], errors="coerce")
            alvos = sorted(alv[COLS["alvo"]].dropna().unique())
            idx = alvos.index("A3") if "A3" in alvos else 0
            sel = st.selectbox("Alvo", alvos, index=idx, key="corr_alvo")
            s = alv[alv[COLS["alvo"]] == sel].sort_values(COLS["data"])
            df_def = s[[COLS["data"], COLS["desl_h"]]].rename(
                columns={COLS["desl_h"]: "def"})
            lbl_def = f"Desl. horizontal {sel} (mm)"
            if not s.empty:
                zona_sel = zona_escav_de_edificio(s[COLS["edificio"]].iloc[0])
    if zona_sel:
        st.caption(f"↳ Frente de escavação associada a **{sel}**: "
                   f"**{ZONA_ESCAV_NOME[zona_sel]}**. Os marcos abaixo são as "
                   f"cotas atingidas nessa frente.")
    else:
        st.caption("↳ Sem zona de escavação associada a esta série — mostram-se "
                   "os marcos globais de escavação.")

    # --- serie da agua (piezometro) ---
    pz = dados["piezo"].copy()
    pz[COLS["data"]] = pd.to_datetime(pz[COLS["data"]], errors="coerce")
    pzs = sorted(pz[COLS["piezometro"]].dropna().unique())
    pz_sel = pzs[0] if pzs else None
    df_agua = None
    if pz_sel:
        sa = pz[pz[COLS["piezometro"]] == pz_sel].sort_values(COLS["data"])
        df_agua = sa[[COLS["data"], COLS["cota_agua"]]].rename(
            columns={COLS["cota_agua"]: "agua"})

    if df_def is None or df_def.empty:
        st.info("Sem dados de deformação para a série escolhida.")
    else:
        fig = go.Figure()
        dt_min = df_def[COLS["data"]].min()
        dt_max = df_def[COLS["data"]].max()
        if df_agua is not None and not df_agua.empty:
            dt_min = min(dt_min, df_agua[COLS["data"]].min())
            dt_max = max(dt_max, df_agua[COLS["data"]].max())
        fases_vis = adicionar_fases_obra(fig, dt_min, dt_max, barra_topo=False)

        # marcos de escavacao da ZONA do instrumento (datas reais) — E1,E2...
        mostrar_escav = st.session_state.get("corr_escav", True)
        escav_visiveis = []
        if mostrar_escav:
            vis = _marcos_escavacao_zona(fig, zona_sel, dt_min, dt_max,
                                         so_escavacao=True)
            escav_visiveis = [(k, rot, t, cota) for k, rot, t, cota, tp in vis]

        # deformacao (eixo Y esquerdo) — listas Python puras (ver nota abaixo)
        # NOTA: converter para listas puras evita a serializacao base64 ("bdata")
        # do Plotly que, com o eixo Y duplo (overlaying), fazia as series NAO
        # renderizarem no Streamlit Cloud (grafico em branco).
        fig.add_trace(go.Scatter(
            x=[pd.to_datetime(v) for v in df_def[COLS["data"]].tolist()],
            y=[float(v) for v in df_def["def"].tolist()],
            mode="lines+markers",
            name=lbl_def, line=dict(color="#c0140f", width=2)))
        if df_agua is not None and not df_agua.empty:
            fig.add_trace(go.Scatter(
                x=[pd.to_datetime(v) for v in df_agua[COLS["data"]].tolist()],
                y=[float(v) for v in df_agua["agua"].tolist()],
                mode="lines+markers",
                name=f"Cota da água {pz_sel} (m)", yaxis="y2",
                line=dict(color="#2563eb", width=2, dash="dot")))

        fig.update_layout(
            height=560,
            margin=dict(t=60, b=130),
            # type="date" EXPLICITO: o 1o trace e um trace fantasma das fases
            # (x=[None]); sem o tipo forcado o eixo X cairia para numerico e as
            # series de data nao renderizavam.
            xaxis=dict(title="Data", type="date"),
            yaxis=dict(title=dict(text=lbl_def, font=dict(color="#c0140f")),
                       tickfont=dict(color="#c0140f")),
            yaxis2=dict(title=dict(text="Cota da água (m)",
                                   font=dict(color="#2563eb")),
                        tickfont=dict(color="#2563eb"),
                        overlaying="y", side="right", anchor="x"),
            legend=dict(orientation="h", yanchor="top", y=-0.38,
                        xanchor="center", x=0.5))
        st.plotly_chart(fig, use_container_width=True)
        legenda_fases(fases_vis)
        if escav_visiveis:
            itens = "  ·  ".join(
                (f"**E{k}** cota {cota:.2f} m ({t.strftime('%d/%m')})"
                 if cota is not None else
                 f"**E{k}** {rotulo} ({t.strftime('%d/%m')})")
                for k, rotulo, t, cota in escav_visiveis)
            st.caption("⛏ Escavação na frente (cota atingida, data do planeamento): " + itens)

        if df_agua is not None and len(df_agua) >= 2 and len(df_def) >= 2:
            d_ini = df_def["def"].iloc[0]
            d_fim = df_def["def"].iloc[-1]
            a_ini = df_agua["agua"].iloc[0]
            a_fim = df_agua["agua"].iloc[-1]
            st.markdown(
                f"**Leitura:** no período, a deformação evoluiu de {d_ini:.1f} "
                f"para **{d_fim:.1f} mm** enquanto a cota da água desceu de "
                f"{a_ini:.2f} para **{a_fim:.2f} m** ({a_fim - a_ini:+.2f} m) — "
                f"as duas séries movem-se em sentido oposto ao longo do tempo.")

    # =====================================================================
    # 2. DEFORMACAO x CARGA NA ANCORAGEM  (mesma zona)
    # =====================================================================
    st.divider()
    st.markdown("#### 2. Deformação × Carga na ancoragem")
    st.caption("Cruza a CARGA medida numa célula de carga (ancoragem) com o "
               "MOVIMENTO de um alvo, no mesmo eixo de tempo. Só faz sentido "
               "cruzar célula e alvo do MESMO alçado/zona. Critério das "
               "células: +15% (alerta), +25% (alarme) sobre a blocagem.")

    alvos_c = dados.get("alvos")
    alvos_c = alvos_c.copy()
    alvos_c[COLS["data"]] = pd.to_datetime(alvos_c[COLS["data"]], errors="coerce")
    ult_data_c = alvos_c[COLS["data"]].max()
    ult_c = alvos_c[alvos_c[COLS["data"]] == ult_data_c]
    alvos_disp2 = sorted(ult_c.nlargest(15, COLS["desl_h"])[COLS["alvo"]]
                         .astype(str).tolist())

    cc = dados.get("celulas")
    if cc is None or cc.empty:
        st.info("Sem dados de células de carga.")
    else:
        cc = cc.copy()
        cc[COLS["data"]] = pd.to_datetime(cc[COLS["data"]], errors="coerce")
        cel_sel = st.selectbox(
            "Célula de carga", sorted(cc[COLS["celula"]].dropna().unique()),
            key="corr_celula")
        loc_cel = LOCALIZACAO_CELULAS.get(cel_sel)
        rotulo_zona = loc_cel[0] if loc_cel else None
        filtro_grupo = loc_cel[1] if loc_cel else None

        sc = cc[cc[COLS["celula"]] == cel_sel].sort_values(COLS["data"])
        bloc = float(sc[COLS["blocagem"]].iloc[0])
        anc = sc[COLS["ancoragem"]].iloc[0]

        if filtro_grupo:
            alvos_validos = sorted(
                ult_c[ult_c[COLS["edificio"]].astype(str)
                      .str.contains(filtro_grupo, case=False, na=False)]
                [COLS["alvo"]].astype(str).unique())
            if not alvos_validos:
                alvos_validos = alvos_disp2
            st.success(
                f"✓ Localização confirmada: a célula {cel_sel} (ancoragem "
                f"{anc}) está na **{rotulo_zona}**. O seletor de alvos abaixo "
                f"está restrito a essa zona — o cruzamento é fisicamente válido.")
        else:
            alvos_validos = alvos_disp2
            st.warning(
                f"⚠ A localização da célula {cel_sel} (ancoragem {anc}) não "
                f"está confirmada. O cruzamento só tem significado físico se o "
                f"alvo pertencer ao mesmo alçado da ancoragem.")

        alvo7 = st.selectbox("Alvo a sobrepor (mesma zona da célula)",
                             alvos_validos, key="corr_alvo_celula")
        _mostrar_imagens_celula(cel_sel, key_suffix="_corr")

        alvos_edi = alvos_c[(alvos_c[COLS["data"]] == ult_data_c) &
                            (alvos_c[COLS["edificio"]].astype(str)
                             .str.contains("Edifício 1", case=False, na=False))]
        if not alvos_edi.empty:
            edi_nome = alvos_edi[COLS["edificio"]].iloc[0]
            with st.expander(f"📍 Ver localização do alvo {alvo7} "
                             f"(no Edifício 1)"):
                COR_ESTADO_LOC = {"Alarme": "#c0140f", "Alerta": "#e67e00",
                                  "Regular": "#1f9e55"}
                mostrar_localizacao_alvo(alvos_edi, edi_nome, [alvo7],
                                         COR_ESTADO_LOC)

        sa2 = alvos_c[alvos_c[COLS["alvo"]].astype(str) == alvo7].sort_values(
            COLS["data"])[[COLS["data"], COLS["desl_h"]]].dropna()

        fig7 = go.Figure()
        fig7.add_trace(go.Scatter(
            x=[pd.to_datetime(v) for v in sc[COLS["data"]].tolist()],
            y=[float(v) for v in sc[COLS["carga_atual"]].tolist()],
            mode="lines+markers", name=f"Carga {cel_sel} (kN)",
            line=dict(color="#7a1fa0", width=2)))
        fig7.add_trace(go.Scatter(
            x=[pd.to_datetime(v) for v in sa2[COLS["data"]].tolist()],
            y=[float(v) for v in sa2[COLS["desl_h"]].tolist()],
            mode="lines+markers", name=f"Desl. H {alvo7} (mm)", yaxis="y2",
            line=dict(color="#c0140f", width=2, dash="dot")))
        for mult, txt, cor in [(1.0, "Blocagem", "#999"),
                               (1.15, "+15% alerta", "#e67e00"),
                               (1.25, "+25% alarme", "#c0140f")]:
            fig7.add_hline(y=bloc * mult, line=dict(color=cor, width=1, dash="dash"),
                           annotation_text=txt, annotation_position="right")
        fig7.update_layout(
            height=460,
            xaxis=dict(title="Data", type="date"),
            yaxis=dict(title=dict(text=f"Carga (kN) — bloc. {bloc:.0f}",
                                  font=dict(color="#7a1fa0")),
                       tickfont=dict(color="#7a1fa0")),
            yaxis2=dict(title=dict(text="Desl. horizontal (mm)",
                                   font=dict(color="#c0140f")),
                        tickfont=dict(color="#c0140f"),
                        overlaying="y", side="right", anchor="x"),
            margin=dict(t=40, b=90),
            legend=dict(orientation="h", yanchor="top", y=-0.2,
                        xanchor="center", x=0.5))
        st.plotly_chart(fig7, use_container_width=True)

        c0 = float(sc[COLS["carga_atual"]].iloc[0])
        c1v = float(sc[COLS["carga_atual"]].iloc[-1])
        pct_fim = (c1v / bloc - 1) * 100
        estado_fim = sc["Estado"].iloc[-1] if "Estado" in sc.columns else "?"
        cresceu = c1v > c0
        st.caption(
            f"A ancoragem {anc} (célula {cel_sel}) "
            f"{'ganhou' if cresceu else 'perdeu'} carga: {c0:.0f} → {c1v:.0f} kN "
            f"({(c1v/c0-1)*100:+.0f}%), terminando a {pct_fim:+.0f}% da blocagem "
            f"({estado_fim}).")

        st.markdown("**Comparação entre células (ancoragens)**")
        linhas_cmp = []
        for cl in sorted(cc[COLS["celula"]].dropna().unique()):
            s = cc[cc[COLS["celula"]] == cl].sort_values(COLS["data"])
            bl = float(s[COLS["blocagem"]].iloc[0])
            v_ini = float(s[COLS["carga_atual"]].iloc[0])
            v_fim = float(s[COLS["carga_atual"]].iloc[-1])
            v_max = float(s[COLS["carga_atual"]].max())
            loc = LOCALIZACAO_CELULAS.get(cl)
            linhas_cmp.append({
                "Célula": cl,
                "Ancoragem": s[COLS["ancoragem"]].iloc[0],
                "Localização": (loc[0].split("—")[-1].strip() if loc else "?"),
                "Período": (f"{s[COLS['data']].min().strftime('%d/%m/%y')}–"
                            f"{s[COLS['data']].max().strftime('%d/%m/%y')}"),
                "Blocagem (kN)": round(bl),
                "Carga fim (kN)": round(v_fim),
                "Variação (%)": round((v_fim / v_ini - 1) * 100, 1),
                "Máx vs bloc. (%)": round((v_max / bl - 1) * 100, 1),
                "Estado": (s["Estado"].iloc[-1] if "Estado" in s.columns else "?"),
            })
        df_cmp = pd.DataFrame(linhas_cmp)
        st.dataframe(df_cmp, use_container_width=True, hide_index=True)

        periodos = cc.groupby(COLS["celula"])[COLS["data"]].min()
        if periodos.nunique() > 1:
            mais_cedo = periodos.idxmin()
            mais_tarde = periodos.idxmax()
            dias_dif = (periodos.max() - periodos.min()).days
            st.warning(
                f"⚠ As células NÃO começaram a medir na mesma data: a "
                f"{mais_tarde} começou ~{dias_dif} dias depois da {mais_cedo}, "
                f"por isso perdeu a fase inicial de carregamento. A comparação "
                f"das variações (%) entre períodos diferentes deve ser lida com "
                f"esta ressalva.")

    # =====================================================================
    # 3. CORRELACAO DIRETA (XY) — deslocamento vs. cota da agua
    # =====================================================================
    st.divider()
    st.markdown("#### 3. Correlação direta — deslocamento vs. cota da água")
    st.caption("Cada ponto é uma campanha: no eixo X a cota da água (interpolada "
               "à data de cada leitura do alvo), no eixo Y o deslocamento "
               "horizontal do alvo. Se os pontos alinham numa tendência, há "
               "correlação. O coeficiente r quantifica a forca da associação "
               "estatística — não prova a causa.")
    alv3 = dados["alvos"].copy()
    alv3[COLS["data"]] = pd.to_datetime(alv3[COLS["data"]], errors="coerce")
    alvos_xy = sorted(alv3[COLS["alvo"]].dropna().unique())
    idx_a3 = alvos_xy.index("A3") if "A3" in alvos_xy else 0
    pz3 = dados["piezo"].copy()
    pz3[COLS["data"]] = pd.to_datetime(pz3[COLS["data"]], errors="coerce")
    pzs3 = sorted(pz3[COLS["piezometro"]].dropna().unique())
    cxy1, cxy2 = st.columns(2)
    with cxy1:
        alvo_xy = st.selectbox("Alvo", alvos_xy, index=idx_a3, key="corr_xy_alvo")
    with cxy2:
        pz_xy = st.selectbox("Piezómetro", pzs3, key="corr_xy_pz") if pzs3 else None

    if pz_xy is not None:
        sa3 = alv3[alv3[COLS["alvo"]] == alvo_xy].sort_values(COLS["data"])[
            [COLS["data"], COLS["desl_h"]]].dropna()
        sp3 = pz3[pz3[COLS["piezometro"]] == pz_xy].sort_values(COLS["data"])[
            [COLS["data"], COLS["cota_agua"]]].dropna()
        if len(sa3) >= 3 and len(sp3) >= 2:
            xp = sp3[COLS["data"]].values.astype("datetime64[ns]").astype("int64")
            yp = sp3[COLS["cota_agua"]].to_numpy(dtype=float)
            xa = sa3[COLS["data"]].values.astype("datetime64[ns]").astype("int64")
            desl = sa3[COLS["desl_h"]].to_numpy(dtype=float)
            # so datas do alvo dentro da janela do piezometro (nao extrapolar)
            mask = (xa >= xp.min()) & (xa <= xp.max())
            xa_in = xa[mask]
            desl_in = desl[mask]
            if len(xa_in) >= 3:
                agua_in = np.interp(xa_in, xp, yp)
                r = float(np.corrcoef(agua_in, desl_in)[0, 1])
                fig_xy = go.Figure()
                fig_xy.add_trace(go.Scatter(
                    x=[float(v) for v in agua_in.tolist()],
                    y=[float(v) for v in desl_in.tolist()],
                    mode="markers", name="Campanhas",
                    marker=dict(color="#c0140f", size=9)))
                b1, b0 = np.polyfit(agua_in, desl_in, 1)
                xs = np.linspace(float(agua_in.min()), float(agua_in.max()), 20)
                fig_xy.add_trace(go.Scatter(
                    x=[float(v) for v in xs.tolist()],
                    y=[float(b1 * v + b0) for v in xs.tolist()],
                    mode="lines", name="Tendência",
                    line=dict(color="#2563eb", width=1, dash="dash")))
                fig_xy.update_layout(
                    height=460,
                    xaxis=dict(title=f"Cota da água {pz_xy} (m)"),
                    yaxis=dict(title=f"Desl. horizontal {alvo_xy} (mm)"),
                    margin=dict(t=30, b=40),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02))
                st.plotly_chart(fig_xy, use_container_width=True)
                forca = ("forte" if abs(r) >= 0.8 else
                         "moderada" if abs(r) >= 0.5 else "fraca")
                sentido = "negativa" if r < 0 else "positiva"
                st.markdown(
                    f"**Correlação:** r = **{r:.2f}** ({forca}, {sentido}), "
                    f"em {len(xa_in)} campanhas. Uma correlação negativa "
                    f"significa que o deslocamento sobe a medida que a cota da "
                    f"água desce. r mede a associação estatística, não a causa.")
            else:
                st.info("Poucos pontos na janela comum ao piezómetro para "
                        "calcular a correlação.")
        else:
            st.info("Série insuficiente (alvo ou piezómetro) para a correlação.")

    # =====================================================================
    # 4. DEFORMACAO x ESCAVACAO DA FRENTE LOCAL  (mesma zona de estacas)
    # =====================================================================
    st.divider()
    st.markdown("#### 4. Deformação × Escavação da frente local")
    st.caption("Cruza o movimento de um alvo com a COTA a que a escavação chegou "
               "NA FRENTE desse alvo (plano de trabalhos impactado, por zona de "
               "estacas). A cota (castanho) desce em degraus a medida que a "
               "frente aprofunda; ve-se a deformação (vermelho) crescer com o "
               "aprofundamento local — e não com a escavação do lado oposto.")

    alv4 = dados["alvos"].copy()
    alv4[COLS["data"]] = pd.to_datetime(alv4[COLS["data"]], errors="coerce")
    alvos4 = sorted(alv4[COLS["alvo"]].dropna().unique())
    idxA = alvos4.index("A3") if "A3" in alvos4 else 0
    alvo4 = st.selectbox("Alvo", alvos4, index=idxA, key="corr_esc_alvo")
    s4 = alv4[alv4[COLS["alvo"]] == alvo4].sort_values(COLS["data"])
    zona4 = (zona_escav_de_edificio(s4[COLS["edificio"]].iloc[0])
             if not s4.empty else None)

    if not zona4 or zona4 not in ESCAVACAO_ZONAS:
        st.info(f"O alvo {alvo4} não tem frente de escavação mapeada. Escolhe um "
                f"alvo de um edifício/alçado com zona associada (Edifícios 1, "
                f"2 e 3, ou um alçado de contenção).")
    else:
        st.caption(f"↳ Frente de escavação: **{ZONA_ESCAV_NOME[zona4]}**")
        # eventos de escavacao (cota) da zona -> frente monotona descendente
        esc_ev = sorted((pd.to_datetime(d), float(cota))
                        for d, tp, cota, rot in ESCAVACAO_ZONAS[zona4]
                        if tp == "escavacao" and cota is not None)
        datas_e = [d for d, c in esc_ev]
        cotas_e, m = [], None
        for _, c in esc_ev:
            m = c if m is None else min(m, c)
            cotas_e.append(m)
        # serie de deformacao do alvo
        sd = s4[[COLS["data"], COLS["desl_h"]]].dropna()
        d_def = [pd.to_datetime(v) for v in sd[COLS["data"]].tolist()]
        y_def = [float(v) for v in sd[COLS["desl_h"]].tolist()]

        # --- grafico de eixo duplo: deformacao (esq) + cota escavada (dir) ---
        fig4 = go.Figure()
        fig4.add_trace(go.Scatter(
            x=d_def, y=y_def, mode="lines+markers",
            name=f"Desl. H {alvo4} (mm)", line=dict(color="#c0140f", width=2)))
        fig4.add_trace(go.Scatter(
            x=datas_e, y=cotas_e, mode="lines+markers",
            name="Cota escavada na frente (m)", yaxis="y2",
            line=dict(color="#8B4513", width=2, shape="hv")))
        # marcas de ancoragem / banda de laje da zona (linhas verticais)
        dmin4 = min([d for d in d_def] + datas_e)
        dmax4 = max([d for d in d_def] + datas_e)
        for d, tp, cota, rot in ESCAVACAO_ZONAS[zona4]:
            if tp == "escavacao":
                continue
            t = pd.to_datetime(d)
            if not (dmin4 <= t <= dmax4):
                continue
            cor = {"ancoragem": "#7a1fa0", "escora": "#1f6fb4"}.get(tp, "#1f9e55")
            fig4.add_vline(x=t, line=dict(color=cor, width=1, dash="dot"),
                           annotation_text={"ancoragem": "⚓", "escora": "⟷"}.get(tp, "▭"),
                           annotation_position="top",
                           annotation_font=dict(size=11, color=cor))
        fig4.update_layout(
            height=460, xaxis=dict(title="Data", type="date"),
            yaxis=dict(title=dict(text=f"Desl. horizontal {alvo4} (mm)",
                                  font=dict(color="#c0140f")),
                       tickfont=dict(color="#c0140f")),
            yaxis2=dict(title=dict(text="Cota escavada na frente (m)",
                                   font=dict(color="#8B4513")),
                        tickfont=dict(color="#8B4513"),
                        overlaying="y", side="right", anchor="x"),
            margin=dict(t=40, b=80),
            legend=dict(orientation="h", yanchor="top", y=-0.18,
                        xanchor="center", x=0.5))
        st.plotly_chart(fig4, use_container_width=True)
        st.caption("⚓ = ancoragens · ⟷ = escoras metálicas · ▭ = banda de laje "
                   "(betonagem), nas datas do plano impactado para a frente. A cota desce em degraus (cada degrau = uma "
                   "cota atingida).")

        # leitura factual
        if len(cotas_e) >= 2 and len(y_def) >= 2:
            st.markdown(
                f"**Leitura:** na frente da {ZONA_ESCAV_NOME[zona4].split('—')[0].strip()}, "
                f"a escavação desceu de cota **{cotas_e[0]:.2f}** para "
                f"**{cotas_e[-1]:.2f} m** enquanto o {alvo4} evoluiu de "
                f"{y_def[0]:.1f} para **{y_def[-1]:.1f} mm**.")

        # --- XY: deslocamento vs. cota escavada local (degrau) + coef. r ---
        st.markdown("**Correlação direta — deslocamento vs. cota escavada local**")
        xe = np.array([pd.Timestamp(d).value for d in datas_e])
        pts_c, pts_d = [], []
        for d, y in zip(d_def, y_def):
            idxs = np.where(xe <= pd.Timestamp(d).value)[0]
            if len(idxs):
                pts_c.append(cotas_e[idxs[-1]])
                pts_d.append(y)
        if len(pts_c) >= 3 and len(set(pts_c)) >= 2:
            r4 = float(np.corrcoef(pts_c, pts_d)[0, 1])
            figx = go.Figure()
            figx.add_trace(go.Scatter(
                x=[float(v) for v in pts_c], y=[float(v) for v in pts_d],
                mode="markers", name="Campanhas",
                marker=dict(color="#c0140f", size=9)))
            b1, b0 = np.polyfit(pts_c, pts_d, 1)
            xs = np.linspace(min(pts_c), max(pts_c), 20)
            figx.add_trace(go.Scatter(
                x=[float(v) for v in xs], y=[float(b1 * v + b0) for v in xs],
                mode="lines", name="Tendência",
                line=dict(color="#8B4513", width=1, dash="dash")))
            figx.update_layout(
                height=430,
                xaxis=dict(title="Cota escavada na frente (m)"),
                yaxis=dict(title=f"Desl. horizontal {alvo4} (mm)"),
                margin=dict(t=30, b=40),
                legend=dict(orientation="h", yanchor="bottom", y=1.02))
            st.plotly_chart(figx, use_container_width=True)
            forca = ("forte" if abs(r4) >= 0.8 else
                     "moderada" if abs(r4) >= 0.5 else "fraca")
            st.markdown(
                f"**Correlação:** r = **{r4:.2f}** ({forca}), em {len(pts_c)} "
                f"campanhas. Como a cota DESCE quando a escavação aprofunda, um "
                f"r negativo significa que o deslocamento sobe a medida que a "
                f"frente local desce. r mede a associação, não a causa.")
        else:
            st.info("Pontos insuficientes na janela da escavação para o XY.")

    # =====================================================================
    # 5. MATRIZ DE CORRELAÇÕES — visão sistemática (reproduz a Tabela 5.3)
    # =====================================================================
    st.divider()
    st.markdown("#### 5. Matriz de correlações entre instrumentos")
    st.caption(
        "Coeficiente de Pearson r entre as séries temporais da instrumentação, "
        "calculado sobre o período comum a cada par (interpolação linear nas "
        "datas da primeira série). À esquerda, os valores ACUMULADOS; à direita, "
        "os INCREMENTOS entre campanhas (variação período-a-período), que medem "
        "se os instrumentos se MOVEM ao mesmo tempo, não apenas se cresceram "
        "ambos. Reproduz a Tabela 5.3 da dissertação.")

    _ser = {}

    def _add_serie(label, df, col, datecol=COLS["data"]):
        if df is None or df.empty or col not in df.columns:
            return
        d = df[[datecol, col]].copy()
        d[datecol] = pd.to_datetime(d[datecol], errors="coerce")
        d = d.dropna(subset=[datecol, col]).sort_values(datecol)
        if d.empty:
            return
        s = d.groupby(datecol)[col].mean()
        if len(s) >= 3:
            _ser[label] = s

    _resm = dados.get("resumo")
    if _resm is not None and not _resm.empty:
        for _inc in sorted(_resm[COLS["inclinometro"]].dropna().unique()):
            _add_serie(str(_inc), _resm[_resm[COLS["inclinometro"]] == _inc],
                       COLS["desl_max_global"])
    _pzm = dados.get("piezo")
    if _pzm is not None and not _pzm.empty:
        for _p in sorted(_pzm[COLS["piezometro"]].dropna().unique()):
            _add_serie(f"Água {_p}", _pzm[_pzm[COLS["piezometro"]] == _p],
                       COLS["cota_agua"])
    _ccm = dados.get("celulas")
    if _ccm is not None and not _ccm.empty:
        for _c in sorted(_ccm[COLS["celula"]].dropna().unique()):
            _add_serie(f"Carga {_c}", _ccm[_ccm[COLS["celula"]] == _c],
                       COLS["carga_atual"])

    _alvm = dados["alvos"].copy()
    _alvos_disp = sorted(_alvm[COLS["alvo"]].dropna().unique())
    _def_alvos = [a for a in ["CD1", "A3", "GH1"] if a in _alvos_disp][:3]
    _sel_alvos = st.multiselect(
        "Alvos a incluir na matriz", _alvos_disp, default=_def_alvos,
        key="corr_matriz_alvos")
    for _a in _sel_alvos:
        _add_serie(_a, _alvm[_alvm[COLS["alvo"]] == _a], COLS["desl_h"])

    if len(_ser) < 2:
        st.info("São precisas pelo menos duas séries para construir a matriz.")
    else:
        _labels = list(_ser.keys())

        def _r_par(sa, sb, incr=False):
            ta = sa.index.astype("int64").to_numpy().astype(float)
            tb = sb.index.astype("int64").to_numpy().astype(float)
            lo, hi = max(ta.min(), tb.min()), min(ta.max(), tb.max())
            m = (ta >= lo) & (ta <= hi)
            if m.sum() < 3:
                return np.nan
            base = ta[m]
            ya = sa.to_numpy().astype(float)[m]
            yb = np.interp(base, tb, sb.to_numpy().astype(float))
            if incr:
                ya, yb = np.diff(ya), np.diff(yb)
            if len(ya) < 2 or np.std(ya) == 0 or np.std(yb) == 0:
                return np.nan
            return float(np.corrcoef(ya, yb)[0, 1])

        def _matriz(incr):
            n = len(_labels)
            M = np.full((n, n), np.nan)
            for i in range(n):
                M[i, i] = 1.0
                for j in range(i + 1, n):
                    r = _r_par(_ser[_labels[i]], _ser[_labels[j]], incr)
                    M[i, j] = M[j, i] = r
            return M

        for _incr, _tit in [(False, "Acumulados"), (True, "Incrementos")]:
            M = _matriz(_incr)
            txt = [[("" if np.isnan(v) else f"{v:.2f}") for v in row]
                   for row in M]
            figM = go.Figure(go.Heatmap(
                z=M, x=_labels, y=_labels, zmid=0, zmin=-1, zmax=1,
                colorscale="RdBu_r", text=txt, texttemplate="%{text}",
                textfont=dict(size=10), colorbar=dict(title="r"),
                hovertemplate="%{y} × %{x}<br>r = %{z:.2f}<extra></extra>"))
            figM.update_layout(
                title=f"Correlação — {_tit}",
                height=140 + 44 * len(_labels),
                xaxis=dict(side="top", tickangle=-45),
                yaxis=dict(autorange="reversed"),
                margin=dict(l=10, r=10, t=90, b=10))
            st.plotly_chart(figM, use_container_width=True)
        st.caption(
            "r > 0: as séries sobem juntas; r < 0: uma sobe enquanto a outra "
            "desce (típico da água, cuja cota desce quando o rebaixamento "
            "acompanha a escavação). A matriz de incrementos é a mais exigente: "
            "só fica alta se os movimentos coincidirem no tempo. Correlação não "
            "é causa — a leitura das causas é do autor no texto.")


def _velocidades_serie(datas, valores, janela7=False):
    """Velocidade (mm/dia) por campanha. Se janela7, cada leitura é comparada
    com a última feita há pelo menos 7 dias; caso contrário, com a anterior.
    Devolve (datas_ordenadas, lista de velocidades alinhada)."""
    d = pd.DataFrame({"data": pd.to_datetime(list(datas)),
                      "val": [float(v) if pd.notna(v) else np.nan for v in valores]})
    d = d.sort_values("data").reset_index(drop=True)
    n = len(d)
    vel = [np.nan] * n
    for i in range(1, n):
        if janela7:
            j = i - 1
            while j >= 0 and (d["data"][i] - d["data"][j]).days < 7:
                j -= 1
            if j < 0:
                continue
        else:
            j = i - 1
        dias = (d["data"][i] - d["data"][j]).days
        if dias > 0 and pd.notna(d["val"][i]) and pd.notna(d["val"][j]):
            vel[i] = (d["val"][i] - d["val"][j]) / dias
    return d["data"].tolist(), vel


def _sinais_serie(vel, limv, usar_vel, usar_acel, fator, dois_seguidos):
    """Lista de bool (sinal precursor por campanha) dada a série de velocidades.
    Regra: vᵢ ≥ limv (velocidade) ou vᵢ₋₁>0 e vᵢ ≥ fator·vᵢ₋₁ (aceleração)."""
    n = len(vel)
    base = [False] * n
    for i in range(1, n):
        v, va = vel[i], vel[i - 1]
        sv = usar_vel and pd.notna(v) and v >= limv
        sa = (usar_acel and pd.notna(v) and pd.notna(va) and va > 0
              and v >= fator * va)
        base[i] = bool(sv or sa)
    if dois_seguidos:
        return [bool(i >= 1 and base[i] and base[i - 1]) for i in range(n)]
    return base


def _nome_elemento(el):
    """Nome curto do elemento (sem o prefixo «Escavação e Contenção Periférica - »)."""
    return str(el).replace("Escavação e Contenção Periférica - ", "")


def sinal_generalizado(alvos_est, limv, fator, frac=2 / 3, n_min=3):
    """Sinal generalizado por elemento (Secção 5.4.1): em cada campanha, fração
    dos alvos ou prismas lidos de um elemento que dão sinal pela regra
    (velocidade ou aceleração), com a velocidade sobre >= 7 dias.
    Cada alvo conta uma vez por data (as leituras de zeragem repetidas na
    mesma data não contam a dobrar). Devolve um DataFrame com el, data, n, k,
    frac e gen (bool)."""
    linhas = []
    for a, g in alvos_est.groupby(COLS["alvo"]):
        g = g[g[COLS["desl_h"]].notna()].sort_values(COLS["data"])
        g = g.drop_duplicates(subset=[COLS["data"]], keep="last")
        if g.empty:
            continue
        ds, vel = _velocidades_serie(g[COLS["data"]], g[COLS["desl_h"]], True)
        sig = _sinais_serie(vel, limv, True, True, fator, False)
        el = _nome_elemento(g[COLS["edificio"]].iloc[0])
        linhas += [(el, a, d, bool(x)) for d, x in zip(ds, sig)]
    if not linhas:
        return pd.DataFrame(columns=["el", "data", "n", "k", "frac", "gen"])
    df = pd.DataFrame(linhas, columns=["el", "alvo", "data", "sig"])
    agg = (df.groupby(["el", "data"]).agg(n=("sig", "size"), k=("sig", "sum"))
           .reset_index())
    agg = agg[agg["n"] >= n_min].copy()
    agg["k"] = agg["k"].astype(int)
    agg["frac"] = agg["k"] / agg["n"]
    agg["gen"] = agg["k"] >= frac * agg["n"] - 1e-9
    return agg.sort_values(["el", "data"]).reset_index(drop=True)


def separador_precursores(dados, limiar_vel, fator_acel, usar_janela7=False):
    """
    PRECURSORES — leitura em tempo real simulada (antecedência face aos
    critérios) e análise de sensibilidade/seletividade da regra nos alvos.
    Reproduz as Secções 5.4 e 5.4.1 e a Tabela 5.5 da dissertação.
    """
    st.subheader("Sinais precursores — antecedência e seletividade")
    st.caption(
        "Um sinal precursor ocorre quando a velocidade do deslocamento "
        "horizontal entre leituras atinge o limiar (velocidade) ou acelera "
        "para o fator definido (aceleração). Os limiares da barra lateral "
        f"(v ≥ {limiar_vel:g} mm/dia; aceleração ≥ {fator_acel:g}×) aplicam-se "
        "à réplica em tempo real; a tabela de sensibilidade usa os valores da "
        "dissertação.")

    alvos_est = anexar_estado_calculado(dados["alvos"])
    dc = None   # data da análise do bloco A (usada também no bloco C)

    # =====================================================================
    # A. RÉPLICA EM TEMPO REAL — antecedência por alvo (Secção 5.1.5)
    # =====================================================================
    st.markdown("#### A. Leitura em tempo real simulada")
    st.caption(
        "Para um alvo, percorre-se o tempo usando apenas as leituras "
        "disponíveis até cada data. Marca-se o primeiro sinal, a entrada em "
        "alerta/alarme (critérios dos relatórios, ΔH ou ΔV) e a antecedência "
        "entre ambos.")

    alvos_disp = sorted(alvos_est[COLS["alvo"]].dropna().unique())
    _idx = alvos_disp.index("A3") if "A3" in alvos_disp else 0
    a_sel = st.selectbox("Alvo", alvos_disp, index=_idx, key="prec_alvo")
    g = alvos_est[alvos_est[COLS["alvo"]] == a_sel].sort_values(COLS["data"])
    g = g[g[COLS["desl_h"]].notna()]
    if len(g) < 3:
        st.info("Série curta de mais para a réplica.")
    else:
        datas = pd.to_datetime(g[COLS["data"]]).tolist()
        if usar_janela7:
            _dd, _vel = _velocidades_serie(g[COLS["data"]], g[COLS["desl_h"]], True)
            _sig = _sinais_serie(_vel, limiar_vel, True, True, fator_acel, False)
            vdf = pd.DataFrame({"data": _dd, "velocidade": _vel, "precursor": _sig})
        else:
            vdf = calcular_velocidade(g[COLS["data"]], g[COLS["desl_h"]],
                                      limiar_vel, fator_acel)
        st.caption("Janela de velocidade: " + ("**≥ 7 dias** (regra melhorada)."
                   if usar_janela7 else "**entre leituras consecutivas**. "
                   "Ativa a opção ≥ 7 dias na barra lateral para reduzir o ruído."))
        dmin, dmax = datas[0].date(), datas[-1].date()
        d_corte = st.slider("Data da análise (replay)", min_value=dmin,
                            max_value=dmax, value=dmax, format="DD/MM/YY",
                            key="prec_corte")
        dc = pd.Timestamp(d_corte)

        est = g["Estado calculado"].tolist()
        def _primeira(cond_datas):
            return cond_datas[0] if cond_datas else None
        d_alerta = _primeira([d for d, e in zip(datas, est)
                              if e in ("Alerta", "Alarme")])
        d_alarme = _primeira([d for d, e in zip(datas, est) if e == "Alarme"])
        datas_sinal = [pd.Timestamp(r["data"]) for _, r in vdf.iterrows()
                       if r["precursor"]]
        d_sinal = _primeira(datas_sinal)
        # até à data de corte (o que se saberia nessa data)
        d_sinal_ate = _primeira([d for d in datas_sinal if d <= dc])
        d_alerta_ate = _primeira([d for d, e in zip(datas, est)
                                  if e in ("Alerta", "Alarme") and d <= dc])

        m1, m2, m3 = st.columns(3)
        m1.metric("1.º sinal", d_sinal.strftime("%d/%m/%Y") if d_sinal else "—")
        m2.metric("Entrada em alerta",
                  d_alerta.strftime("%d/%m/%Y") if d_alerta is not None else "—")
        if d_sinal is not None and d_alerta is not None:
            ant = (d_alerta - d_sinal).days
            m3.metric("Antecedência", f"{ant} dias" if ant >= 0 else "sinal após alerta",
                      delta=f"≈ {ant/7:.0f} semanas" if ant >= 0 else None)
        else:
            m3.metric("Antecedência", "—")

        fig = go.Figure()
        xx = [pd.Timestamp(d) for d in datas]
        yy = [float(v) for v in g[COLS["desl_h"]].tolist()]
        fig.add_trace(go.Scatter(x=xx, y=yy, mode="lines+markers",
                                 name="δh acumulado", line=dict(color="#1f3a5f")))
        sig_x = [pd.Timestamp(r["data"]) for _, r in vdf.iterrows() if r["precursor"]]
        sig_y = []
        for sx in sig_x:
            k = datas.index(sx) if sx in datas else None
            sig_y.append(yy[k] if k is not None else np.nan)
        if sig_x:
            fig.add_trace(go.Scatter(x=sig_x, y=sig_y, mode="markers",
                          name="Sinal precursor",
                          marker=dict(color="#c0140f", size=11, symbol="triangle-up")))
        if d_alerta is not None:
            fig.add_vline(x=d_alerta, line=dict(color="#e67e00", width=1.5, dash="dash"),
                          annotation_text="Alerta", annotation_position="top")
        if d_alarme is not None:
            fig.add_vline(x=d_alarme, line=dict(color="#c0140f", width=1.5, dash="dash"),
                          annotation_text="Alarme", annotation_position="top")
        fig.add_vrect(x0=xx[0], x1=dc, fillcolor="#1f3a5f", opacity=0.05, line_width=0)
        fig.add_vline(x=dc, line=dict(color="#1f3a5f", width=1, dash="dot"),
                      annotation_text="Análise", annotation_position="bottom right")
        fig.update_layout(height=430, xaxis=dict(title="Data"),
                          yaxis=dict(title="δh acumulado (mm)"),
                          margin=dict(l=10, r=10, t=30, b=40),
                          legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(fig, use_container_width=True)

        if d_sinal_ate is not None:
            txt_ate = (f"Até **{d_corte:%d/%m/%Y}**, o primeiro sinal já teria sido "
                       f"lido a **{d_sinal_ate:%d/%m/%Y}**")
            if d_alerta_ate is not None:
                txt_ate += f" e o alvo já teria entrado em alerta ({d_alerta_ate:%d/%m/%Y})."
            else:
                txt_ate += ", antes de qualquer alerta."
            st.caption(txt_ate)
        else:
            st.caption(f"Até {d_corte:%d/%m/%Y}, nenhum sinal precursor teria sido lido.")

    # =====================================================================
    # B. SENSIBILIDADE E SELETIVIDADE — Tabela 5.5 (Secção 5.4.1)
    # =====================================================================
    st.divider()
    st.markdown("#### B. Sensibilidade e seletividade da regra (todos os alvos)")
    st.caption(
        "A regra e as suas variantes aplicadas ao deslocamento horizontal de "
        "todos os alvos e prismas. Contam-se os sinais anteriores ao primeiro "
        "alerta de cada alvo (todos, nos alvos sem alerta). Um alerta diz-se "
        "antecipado se teve um sinal nas 4 semanas anteriores; um sinal é "
        "acerto se o alerta do mesmo alvo se seguiu nesse prazo. Reproduz a "
        "Tabela 5.5.")

    # pré-computar, por alvo: datas, δh e data do 1.º alerta (ΔH ou ΔV)
    por_alvo = {}
    for a, gg in alvos_est.groupby(COLS["alvo"]):
        gg = gg.sort_values(COLS["data"])
        gg = gg[gg[COLS["desl_h"]].notna()]
        if len(gg) < 2:
            continue
        ds = pd.to_datetime(gg[COLS["data"]]).tolist()
        dh = gg[COLS["desl_h"]].tolist()
        est = gg["Estado calculado"].tolist()
        d_al = None
        for dt, e in zip(ds, est):
            if e in ("Alerta", "Alarme"):
                d_al = dt
                break
        por_alvo[a] = (ds, dh, d_al)

    FATOR_TAB = 1.8
    VARIANTES = [
        ("Velocidade entre leituras consecutivas", None, None),
        ("  Regra da aplicação (v ≥ 0,5 ou acel. ≥ 1,8)",
         dict(janela7=False, limv=0.5, usar_vel=True, usar_acel=True, dois=False), None),
        ("  Só velocidade, v ≥ 0,3",
         dict(janela7=False, limv=0.3, usar_vel=True, usar_acel=False, dois=False), None),
        ("  Só velocidade, v ≥ 0,5",
         dict(janela7=False, limv=0.5, usar_vel=True, usar_acel=False, dois=False), None),
        ("  Só velocidade, v ≥ 1,0",
         dict(janela7=False, limv=1.0, usar_vel=True, usar_acel=False, dois=False), None),
        ("  Só aceleração, ≥ 1,8×",
         dict(janela7=False, limv=0.5, usar_vel=False, usar_acel=True, dois=False), None),
        ("Velocidade sobre intervalo ≥ 7 dias", None, None),
        ("  Regra da aplicação (v ≥ 0,5 ou acel. ≥ 1,8)",
         dict(janela7=True, limv=0.5, usar_vel=True, usar_acel=True, dois=False), None),
        ("  Só velocidade, v ≥ 0,5",
         dict(janela7=True, limv=0.5, usar_vel=True, usar_acel=False, dois=False), None),
        ("  Só velocidade, v ≥ 0,3, dois sinais seguidos",
         dict(janela7=True, limv=0.3, usar_vel=True, usar_acel=False, dois=True), None),
    ]

    def _computar(cfg):
        sinais_tot = acertos = 0
        al_4s = al_antes = sem_al_sinal = 0
        alvos_sinal = alvos_sinal_alerta = 0
        antec = []
        n_al = n_sem = 0
        for ds, dh, d_al in por_alvo.values():
            _, vel = _velocidades_serie(ds, dh, cfg["janela7"])
            sig = _sinais_serie(vel, cfg["limv"], cfg["usar_vel"],
                                cfg["usar_acel"], FATOR_TAB, cfg["dois"])
            ds_sinal = [ds[i] for i in range(len(ds)) if sig[i]]
            if d_al is not None:
                n_al += 1
                antes = [d for d in ds_sinal if d < d_al]
                sinais_tot += len(antes)
                acertos += sum(1 for d in antes if 0 < (d_al - d).days <= 28)
                if antes:
                    al_antes += 1
                    alvos_sinal += 1
                    alvos_sinal_alerta += 1
                    antec.append((d_al - antes[0]).days)
                if any(0 < (d_al - d).days <= 28 for d in antes):
                    al_4s += 1
            else:
                n_sem += 1
                sinais_tot += len(ds_sinal)
                if ds_sinal:
                    sem_al_sinal += 1
                    alvos_sinal += 1
        return dict(
            sinais=sinais_tot,
            pct_seg=round(100 * acertos / sinais_tot) if sinais_tot else 0,
            al_4s=al_4s, al_antes=al_antes, sem=sem_al_sinal,
            pct_alvos=round(100 * alvos_sinal_alerta / alvos_sinal) if alvos_sinal else 0,
            antec=int(np.median(antec)) if antec else None,
            n_al=n_al, n_sem=n_sem)

    linhas = []
    base_ref = None
    for rotulo, cfg, _ in VARIANTES:
        if cfg is None:
            linhas.append({"Regra": f"▸ {rotulo.strip()}", "Sinais": "",
                           "Sinais seguidos de alerta (%)": "",
                           "Alertas c/ sinal ≤4 sem.": "",
                           "Alertas c/ sinal antes": "",
                           "Alvos s/ alerta c/ sinal": "",
                           "Alvos c/ sinal → alerta (%)": ""})
            continue
        r = _computar(cfg)
        if base_ref is None:
            base_ref = r
        linhas.append({
            "Regra": rotulo,
            "Sinais": r["sinais"],
            "Sinais seguidos de alerta (%)": r["pct_seg"],
            "Alertas c/ sinal ≤4 sem.": r["al_4s"],
            "Alertas c/ sinal antes": r["al_antes"],
            "Alvos s/ alerta c/ sinal": r["sem"],
            "Alvos c/ sinal → alerta (%)": r["pct_alvos"]})
    st.table(pd.DataFrame(linhas).set_index("Regra"))
    if base_ref is not None:
        st.caption(
            f"Dos {base_ref['n_al'] + base_ref['n_sem']} alvos, "
            f"**{base_ref['n_al']} atingiram o alerta** e {base_ref['n_sem']} "
            f"nunca o atingiram (contagens «de 38» e «de 66» na dissertação). "
            f"Com a regra da aplicação, a antecedência mediana entre o primeiro "
            f"sinal e o alerta é de **{base_ref['antec']} dias**. A regra é "
            f"sensível mas pouco seletiva: a maioria dos alvos que nunca "
            f"alertaram também gerou sinais, e calcular a velocidade sobre ≥7 "
            f"dias melhora a seletividade à custa de antecipar menos alertas.")

    # =====================================================================
    # C. SINAL GENERALIZADO POR ELEMENTO (Secção 5.4.1)
    # =====================================================================
    st.divider()
    st.markdown("#### C. Sinal generalizado por elemento")
    st.caption(
        "Persistência no espaço: numa campanha, um elemento (edifício vizinho ou "
        "alçado) dá sinal generalizado quando pelo menos a fração escolhida dos "
        "seus alvos ou prismas lidos nessa campanha (mínimo de 3) dá sinal pela "
        "regra da barra lateral, com a velocidade calculada sobre pelo menos 7 "
        "dias. Reduz o número de sinais a analisar face aos sinais por alvo.")
    frac_txt = st.radio("Fração mínima de alvos com sinal", ["2/3", "1/2"],
                        horizontal=True, key="gen_frac")
    frac = 2 / 3 if frac_txt == "2/3" else 0.5
    df_g = sinal_generalizado(alvos_est, limiar_vel, fator_acel, frac)
    if df_g.empty:
        st.info("Sem elementos com pelo menos 3 leituras na mesma campanha.")
        return
    gen = df_g[df_g["gen"]]
    # primeiro alerta (ΔH ou ΔV) de cada elemento
    _al = alvos_est[alvos_est["Estado calculado"].isin(["Alerta", "Alarme"])]
    prim_alerta = (_al.groupby(COLS["edificio"])[COLS["data"]].min()
                   .rename(index=_nome_elemento))
    ordem = sorted(df_g["el"].unique(),
                   key=lambda e: (not e.startswith("Edifício"), e))

    figG = go.Figure()
    figG.add_trace(go.Scatter(
        x=df_g["data"], y=df_g["el"], mode="markers",
        marker=dict(size=9, color=df_g["frac"], cmin=0, cmax=1,
                    colorscale=[[0, "#f2f2f2"], [0.5, "#f4b183"], [1, "#c0140f"]],
                    colorbar=dict(title="Fração com sinal", thickness=10, len=0.6),
                    line=dict(width=0.5, color="#9a9a9a")),
        customdata=np.stack([df_g["k"], df_g["n"]], axis=-1),
        hovertemplate="%{y}<br>%{x|%d/%m/%Y}: %{customdata[0]} de "
                      "%{customdata[1]} alvos com sinal<extra></extra>",
        name="Campanha (fração com sinal)"))
    figG.add_trace(go.Scatter(
        x=gen["data"], y=gen["el"], mode="markers",
        marker=dict(symbol="star", size=15, color="#7f0000",
                    line=dict(width=1, color="white")),
        hovertemplate="%{y}<br>%{x|%d/%m/%Y}: sinal generalizado<extra></extra>",
        name="Sinal generalizado"))
    pa = prim_alerta[prim_alerta.index.isin(ordem)]
    figG.add_trace(go.Scatter(
        x=pa.values, y=pa.index, mode="markers",
        marker=dict(symbol="diamond-open", size=13, color="#e67e00",
                    line=dict(width=2.5)),
        hovertemplate="%{y}<br>Primeiro alerta: %{x|%d/%m/%Y}<extra></extra>",
        name="Primeiro alerta"))
    if dc is not None:
        figG.add_vline(x=dc, line=dict(color="#1f3a5f", width=1, dash="dot"),
                       annotation_text="Análise", annotation_position="bottom right")
    figG.update_layout(
        height=110 + 30 * len(ordem), margin=dict(l=10, r=10, t=30, b=40),
        yaxis=dict(categoryorder="array", categoryarray=ordem[::-1], title=""),
        xaxis=dict(title="Data"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(figG, use_container_width=True)

    # quadro por elemento
    lin = []
    for el in ordem:
        g_el = gen[gen["el"] == el]
        d1 = g_el["data"].min() if len(g_el) else None
        da = prim_alerta.get(el)
        ant = (da - d1).days if (d1 is not None and da is not None and pd.notna(da)) else None
        lin.append({"Elemento": el, "Sinais generalizados": len(g_el),
                    "Primeiro sinal generalizado": d1.strftime("%d/%m/%Y") if d1 is not None else "—",
                    "Primeiro alerta": da.strftime("%d/%m/%Y") if da is not None and pd.notna(da) else "—",
                    "Antecedência (dias)": str(ant) if ant is not None and ant > 0 else "—"})
    st.table(pd.DataFrame(lin).set_index("Elemento"))
    n_el_al = sum(1 for r in lin if r["Primeiro alerta"] != "—")
    ants = [int(r["Antecedência (dias)"]) for r in lin if r["Antecedência (dias)"] != "—"]
    if _lang_en():
        st.caption(
            f"**{len(gen)} generalised signals** in {len(df_g)} element–campaign "
            f"pairs with at least 3 readings, in {gen['el'].nunique()} of "
            f"{len(ordem)} elements. In **{len(ants)} of the {n_el_al} elements** "
            f"that reached the alert, the first generalised signal preceded it"
            + (f" ({min(ants)} to {max(ants)} days)." if ants else ".")
            + " As almost all elements reached the alert, the selectivity of this "
            "indicator cannot be assessed in this case.")
    else:
        st.caption(
            f"**{len(gen)} sinais generalizados** em {len(df_g)} pares de elemento e "
            f"campanha com pelo menos 3 leituras, em {gen['el'].nunique()} de "
            f"{len(ordem)} elementos. Em **{len(ants)} dos {n_el_al} elementos** que "
            f"atingiram o alerta, o primeiro sinal generalizado precedeu-o"
            + (f" ({min(ants)} a {max(ants)} dias)." if ants else ".")
            + " Como quase todos os elementos atingiram o alerta, a seletividade "
            "deste indicador não pode ser avaliada neste caso.")
    if dc is not None:
        ate = gen[gen["data"] <= dc]
        if len(ate):
            partes = [f"{el} ({', '.join(d.strftime('%d/%m/%Y') for d in g['data'])})"
                      for el, g in ate.groupby("el", sort=False)]
            st.caption(f"Até **{dc:%d/%m/%Y}** (data da análise do bloco A), sinal "
                       f"generalizado em: " + "; ".join(partes) + ".")
        else:
            st.caption(f"Até {dc:%d/%m/%Y} (data da análise do bloco A), nenhum "
                       f"elemento tinha dado sinal generalizado.")


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
        "max-width:560px;'>Análise e visualização de instrumentação geotécnica "
        "— deslocamentos, velocidades, sinais precursores, geologia e "
        "sequência de obra.</div>"
        "</div></div>"
    )
    st.markdown(banner, unsafe_allow_html=True)

    # ---- enquadramento do caso (o problema, para o juri entrar na narrativa)
    st.markdown(
        "Esta aplicação apoia a **análise inversa** (*back-analysis*) da "
        "contenção periférica da reformulação do **Hotel Eden**, no Monte "
        "Estoril. A obra compreende uma escavação profunda — fundo à cota "
        "4,55 m, cerca de 16 m abaixo do piso 1 (cota 20,95 m) e mais nas zonas "
        "em que o terreno natural é mais alto — executada ao abrigo de uma "
        "cortina de estacas ancorada e travada por bandas de laje, num contexto "
        "exigente: confina com **edifícios sensíveis**, em particular três "
        "edifícios vizinhos instrumentados (Edifícios 1, 2 e 3), "
        "e desenvolve-se **abaixo do nível freático** de repouso. A "
        "instrumentação instalada — inclinómetros, alvos topográficos, células "
        "de carga e piezómetros — permite acompanhar, ao longo do tempo, os "
        "deslocamentos induzidos pela escavação, tanto na própria contenção "
        "como nos edifícios vizinhos.")
    st.markdown(
        "O propósito da ferramenta não se esgota na visualização das leituras: "
        "procura sobretudo **relacionar a deformação medida com os "
        "acontecimentos da obra** — o avanço da escavação, o rebaixamento do nível freático, "
        "a execução das ancoragens e a natureza do maciço. O separador "
        "**Correlações** cruza estas séries num eixo temporal comum; a "
        "interpretação das causas cabe ao engenheiro, e uma correlação "
        "temporal não demonstra, por si só, causalidade.")

    # ---- identificacao da obra + numeros-chave --------------------------
    col_id, col_num = st.columns([1.3, 2])
    with col_id:
        st.markdown("#### Obra")
        st.markdown(
            "**Hotel Eden, Estoril**  \n"
            "Reformulação — escavação e contenção periférica  \n"
            "Monte Estoril, Cascais  \n"
            "_Back-analysis de instrumentação_")

    with col_num:
        st.markdown("#### Instrumentação monitorizada")
        # calcular numeros reais a partir dos dados
        n_inc = dados["perfis"][COLS["inclinometro"]].nunique() if not dados["perfis"].empty else 0
        n_alvos = dados["alvos"][COLS["alvo"]].nunique() if not dados["alvos"].empty else 0
        n_camp = dados["alvos"][COLS["data"]].nunique() if not dados["alvos"].empty else 0
        n_cel = dados["celulas"][COLS["celula"]].nunique() if not dados["celulas"].empty else 0
        n_pz = dados["piezo"][COLS["piezometro"]].nunique() if not dados["piezo"].empty else 0
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Inclinómetros", n_inc)
        c2.metric("Alvos topogr.", n_alvos)
        c3.metric("Campanhas", n_camp)
        c4.metric("Células carga", n_cel)
        c5.metric("Piezómetros", n_pz)

    st.divider()

    # ---- cobertura das leituras / avisos de atraso (Tabela 6.1) ---------
    st.markdown("#### Cobertura das leituras — avisos de atraso")
    _fontes = [("alvos", COLS["alvo"], "Alvo"),
               ("resumo", COLS["inclinometro"], "Inclinómetro"),
               ("celulas", COLS["celula"], "Célula"),
               ("piezo", COLS["piezometro"], "Piezómetro")]
    _ultimas, _regs = [], []
    for _chave, _colid, _sis in _fontes:
        _df = dados.get(_chave)
        if _df is None or _df.empty or COLS["data"] not in _df.columns:
            continue
        _dd = _df.copy()
        _dd[COLS["data"]] = pd.to_datetime(_dd[COLS["data"]], errors="coerce")
        _ultimas.append(_dd[COLS["data"]].max())
    _data_estado = max([d for d in _ultimas if pd.notna(d)]) if _ultimas else None
    if _data_estado is not None:
        for _chave, _colid, _sis in _fontes:
            _df = dados.get(_chave)
            if _df is None or _df.empty or _colid not in _df.columns:
                continue
            _dd = _df.copy()
            _dd[COLS["data"]] = pd.to_datetime(_dd[COLS["data"]], errors="coerce")
            for _inst, _gi in _dd.groupby(_colid):
                _gi = _gi.sort_values(COLS["data"])
                _ult = _gi[COLS["data"]].max()
                _difs = _gi[COLS["data"]].diff().dt.days.dropna()
                _cad = float(_difs.median()) if len(_difs) else np.nan
                _atraso = int((_data_estado - _ult).days)
                _flag = bool(pd.notna(_cad) and _cad > 0
                             and _atraso > 2 * _cad and _atraso >= 10)
                _regs.append({"Instrumento": str(_inst), "Sistema": _sis,
                              "Última leitura": _ult.strftime("%d/%m/%Y"),
                              "Cadência (d)": (round(_cad) if pd.notna(_cad) else "—"),
                              "Dias desde a última": _atraso, "_f": _flag})
        _reg = pd.DataFrame(_regs)
        _n_atraso = int(_reg["_f"].sum()) if not _reg.empty else 0
        st.caption(
            f"Data de estado (última leitura global): **{_data_estado:%d/%m/%Y}**. "
            "Um instrumento é assinalado quando o intervalo desde a sua última "
            "leitura excede o dobro da sua cadência habitual.")
        if _n_atraso:
            _atr = _reg[_reg["_f"]].sort_values("Dias desde a última", ascending=False)
            st.warning(f"⚠ {_n_atraso} instrumento(s) com leitura em atraso face à "
                       f"data de estado.")
            st.dataframe(_atr.drop(columns="_f"), use_container_width=True,
                         hide_index=True)
            st.caption("As células de carga terminaram a 06/07/2026 (fim das leituras "
                       "indicado pela obra, a confirmar o motivo); os restantes casos "
                       "são instrumentos cuja série termina antes da data de estado.")
        else:
            st.success("Sem instrumentos em atraso face à data de estado.")

    # ---- cartoes das areas de analise -----------------------------------
    st.markdown("#### O que podes explorar")
    st.caption("Há duas componentes, selecionáveis na barra lateral: **Dados** "
               "(ilustração dos resultados de cada instrumento, geologia e "
               "planeamento, sem análise) e **Análise** (leitura dos resultados "
               "e correlação entre instrumentos). Usa os separadores no topo "
               "para navegar dentro de cada componente.")

    cartoes = [
        ("Dados · Inclinómetros", "Perfil deformado em profundidade, evolução "
         "no tempo, velocidade e sinais precursores. Sobreposição da geologia."),
        ("Dados · Alvos (2D)", "Séries temporais de deslocamento horizontal e "
         "assentamento vertical, por edifício e por alvo."),
        ("Dados · Células de carga", "Carga nas ancoragens vs. blocagem, com "
         "limiares de alerta (15%) e alarme (25%)."),
        ("Dados · Piezómetros", "Evolução da cota da água subterrânea."),
        ("Dados · Geologia", "Colunas litológicas das sondagens, ensaios SPT em "
         "profundidade e zonamento geotécnico (ZG1-ZG6)."),
        ("Dados · Obra", "Plano de trabalhos contratual e impactado, as suspensões "
         "de frentes (impactos) e o desfasamento de cada frente."),
        ("Dados · Planta (DXF)", "Leitura de plantas de escavação em DXF, com "
         "opção de alinhamento aos alvos por pontos de referência."),
        ("Análise · Terreno + Alvos 3D", "Alvos no espaço com a geometria da "
         "obra: contorno do recinto, edifícios vizinhos e vetores de "
         "deslocamento amplificados."),
        ("Análise · Resultados", "Magnitude, distribuição espacial, ritmo e "
         "profundidade das movimentações, e observado vs. previsto. Ilustração "
         "factual dos resultados, para a escrita da tese."),
        ("Análise · Correlações", "Gráficos de interação entre instrumentos: "
         "deformação vs. água vs. escavação vs. faseamento; deformação vs. "
         "carga na ancoragem; e a correlação direta (XY) com coeficiente r."),
        ("Análise · Pressupostos", "O que é medido vs. o que é assumido: "
         "inferências, limitações e qualidade dos dados."),
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
    st.caption("Nota metodológica: a integração dos perfis inclinométricos "
               "assume a base fixa. As datas da sequência de obra provem do "
               "plano de trabalhos impactado (o contratual com três suspensões "
               "de frentes), não de um registo diário de obra. A posição dos alvos "
               "no 3D é aproximada — apoio visual, não modelo do maciço.")


# =========================================================================
# PRINCIPAL
# =========================================================================
def main():
    st.sidebar.title("The Eden - Instrumentação")
    st.sidebar.radio("Idioma / Language", ["Português", "English"],
                     horizontal=True, key="ims_lang")
    fonte_op = st.sidebar.radio("Ficheiro de dados",
                                ["Usar o ficheiro ao lado do script", "Carregar manualmente"])
    if fonte_op == "Carregar manualmente":
        up = st.sidebar.file_uploader("Excel de instrumentação", type=["xlsx"])
        if up is None:
            st.info("Carrega o Excel na barra lateral para começar.")
            st.stop()
        fonte = up
    else:
        # caminho do Excel ancorado ao diretorio do script (robusto ao CWD)
        base_dir = Path(__file__).resolve().parent
        fonte = base_dir / FICHEIRO_EXCEL
        if not fonte.exists():
            # tentar tambem o CWD, por compatibilidade
            if Path(FICHEIRO_EXCEL).exists():
                fonte = FICHEIRO_EXCEL
            else:
                st.error(f"Não encontrei '{FICHEIRO_EXCEL}'. Poe o Excel na pasta "
                         f"do script ou usa 'Carregar manualmente'.")
                st.stop()
    try:
        dados = carregar_dados(fonte)
    except Exception as e:
        st.error(f"Não consegui ler o Excel. Detalhe: {e}")
        st.stop()
    aplicar_registos_ancoragens(dados)

    st.sidebar.divider()
    st.sidebar.subheader("Deteção de precursores")
    limiar_vel = st.sidebar.slider("Limiar de velocidade (mm/dia)", 0.1, 3.0,
                                   LIMIAR_VEL_DEFEITO, 0.05)
    fator_acel = st.sidebar.slider("Fator de aceleração (x)", 1.2, 3.0,
                                   FATOR_ACEL_DEFEITO, 0.1)
    usar_janela7 = st.sidebar.checkbox(
        "Velocidade sobre ≥ 7 dias (reduz ruído)", value=False,
        help="Regra melhorada (Secção 5.3.5): compara cada leitura com a última "
             "feita há pelo menos 7 dias, em vez da leitura imediatamente "
             "anterior. Aplica-se à réplica em tempo real dos precursores.")

    if not TEM_SCIPY:
        st.sidebar.info("Instala 'scipy' para ativar a superfície 3D interpolada.")

    if not TEM_EZDXF:
        st.sidebar.info("Instala 'ezdxf' para ativar a leitura de plantas DXF.")

    # -------------------------------------------------------------------
    # ORGANIZACAO EM DUAS COMPONENTES (escolhidas na barra lateral)
    #   Dados:   ilustracao dos resultados — instrumentacao, geologia e
    #            planeamento (real e teorico). SEM analise.
    #   Analise: leitura de resultados e correlacao entre instrumentos.
    #            «Resultados» = magnitude/onde/ritmo/profundidade/observado-vs-
    #            previsto (factual, para a escrita da tese). «Correlacoes» =
    #            graficos de interacao entre instrumentos (deformacao x agua x
    #            escavacao x faseamento; deformacao x carga; correlacao direta).
    # O separador Inicio esta sempre acessivel (primeiro de cada componente).
    # -------------------------------------------------------------------
    st.sidebar.divider()
    st.sidebar.subheader("Navegação")
    frente = st.sidebar.radio(
        "Componente",
        ["Dados — instrumentação, geologia e planeamento",
         "Análise — análise de resultados"],
        help="Dados: ilustração dos resultados (cada instrumento, a geologia e "
             "o planeamento), sem análise. Análise: leitura dos resultados e "
             "correlação entre os dados dos vários instrumentos.")

    if frente.startswith("Dados"):
        thome, tinc, talv, tcc, tpz, tgeo, tobra, tplan = st.tabs(
            ["Início", "Inclinómetros", "Alvos (2D)", "Células de carga",
             "Piezómetros", "Geologia", "Obra", "Planta (DXF)"])
        with thome:
            separador_home(dados)
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
    else:
        thome, t3d, tres, tcorr, tprec, tpress = st.tabs(
            ["Início", "Terreno + Alvos 3D", "Resultados", "Correlações",
             "Precursores", "Pressupostos"])
        with thome:
            separador_home(dados)
        with t3d:
            separador_terreno_alvos_3d(dados)
        with tres:
            separador_resultados(dados)
        with tcorr:
            separador_correlacoes(dados)
        with tprec:
            separador_precursores(dados, limiar_vel, fator_acel, usar_janela7)
        with tpress:
            separador_pressupostos(dados)


if __name__ == "__main__":
    main()
