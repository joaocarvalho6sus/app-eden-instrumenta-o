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

# Mapeamento CONFIRMADO (nao inferido) da zona de projeto para os edificios
# vizinhos, com base na orientacao real verificada em planta/mapa:
#   - Santa Casa: noroeste (poente + norte)  -> poente/norte (20 mm)
#   - Cimas:      nordeste (nascente + norte) -> poente/norte (20 mm), pela
#                 regra do projeto "poente OU norte = 20 mm"
#   - Clinica:    nascente + norte            -> poente/norte (20 mm)
# O referencial (M,P) dos alvos esta rodado face ao norte real, por isso a
# inferencia geometrica NAO e fiavel; estes tres sao fixados a mao.
ZONA_FIXA_GRUPO = {
    "Santa Casa": "poente_norte",
    "Cimas": "poente_norte",
    "Clínica": "poente_norte",
    "Clinica": "poente_norte",
}

# Localizacao das CELULAS DE CARGA (alcado/zona onde esta a ancoragem).
# So faz sentido cruzar a carga de uma celula com o movimento de alvos do
# MESMO alcado/zona. CONFIRMADO em planta e modelo 3D do projeto: ambas as
# celulas estao na cortina sob o edificio da Santa Casa (lado poente, ZG2,
# junto a ancoragem A26). Por isso cruzam validamente com os alvos da Santa
# Casa (A1-A8...). Cada entrada: (rotulo da zona, filtro de grupo de alvos).
#   - CC 2501796 / A26: cortina da Santa Casa (Piso -1, cota 15,90)
#   - CC 200792  / DE : cortina da Santa Casa (proxima da A26)
LOCALIZACAO_CELULAS = {
    "CC 2501796": ("Cortina sob a Santa Casa — ZG2, Piso -1 (cota 15,90), junto a A26", "Santa Casa"),
    "CC 200792":  ("Cortina sob a Santa Casa — ZG2, Piso -2 (cota 12,45), perto da A26", "Santa Casa"),
}

# Imagens de localizacao de cada celula (planta/alcado do projeto + contexto).
# Ficheiros em fotos_celulas/ (ao lado do script). Cada entrada: lista de
# (ficheiro, legenda). A "planta_geral" e partilhada como contexto do recinto.
IMAGENS_CELULAS = {
    "CC 2501796": [
        ("cc_2501796_planta.jpg",
         "Planta de localizacao das celulas — A26, Piso -1 (cota 15,90), ZG2"),
        ("planta_geral.jpg",
         "Planta de localizacao geral (recinto) — celula assinalada a poente"),
    ],
    "CC 200792": [
        ("cc_200792_alcado.jpg",
         "Alcado D-E — ancoragem DE, Piso -2 (cota 12,45), ZG2"),
        ("cc_200792_3d.jpg",
         "Modelo 3D do projeto — localizacao da celula (lado da Santa Casa)"),
        ("planta_geral.jpg",
         "Planta de localizacao geral (recinto) — celula assinalada a poente"),
    ],
}

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
    "Aterro": "#c0641e",          # mesmo laranja do ZG6 (aterro)
    "Gres (C1As)": "#a9c47f",     # verde-base do gres (alinhado ao ZG4)
    "Calcario (C1A)": "#5c8a45",  # verde mais escuro (calcario de fundo)
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
COTA_MURO_SCML = 22.50               # muro tradicional na fronteira com a Santa Casa

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
    "I1": {"sondagem": "SC8/Pz", "confianca": "media-alta",
           "azimute": 330, "posicao": "canto SO (poente)"},
    "I2": {"sondagem": "SC9/Pz", "confianca": "alta",
           "azimute": 225, "posicao": "topo N (bolbo curvo)"},
    "I3": {"sondagem": "SC6/Pz", "confianca": "media",
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
# FASEAMENTO DA OBRA — datas REAIS de execucao (plano de trabalhos impactado,
# Aquatecnica, 22/04/2026). As macro-fases usam as datas impactadas (o que foi
# de facto executado). Fonte: "Plano de Trabalhos - Impactado ECP".
# =========================================================================
FASES_OBRA = [
    ("Contencao periferica",                    "2025-05-13", "2026-04-15"),
    ("Estacas Poente e Norte",                  "2025-05-13", "2025-08-12"),
    ("Estacas Central e Nascente",              "2025-08-13", "2025-10-07"),
    ("Estacas Cimas",                           "2025-05-13", "2025-09-25"),
    ("Escavacao + bandas de laje + ancoragens", "2025-07-08", "2026-04-15"),
    ("Viga coroamento + muros contencao",       "2025-07-08", "2025-11-13"),
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

CORES_FASES = ["#8dd3c7", "#ffffb3", "#bebada", "#fb8072", "#80b1d3",
               "#fdb462", "#b3de69", "#fccde5", "#d9d9d9"]

# FASEAMENTO: real (impactado) vs contratual (previsto), para comparacao de
# desempenho face ao prazo. Fonte: plano de trabalhos impactado (Aquatecnica).
# (nome, ini_real, fim_real, dur_real_dias, dur_contratual_dias)
FASES_COMPARACAO = [
    ("Contencao periferica",        "2025-05-13", "2026-04-15", 242, 210),
    ("Execucao de estacas",         "2025-05-13", "2025-10-07", 106,  70),
    ("Estacas Poente e Norte",      "2025-05-13", "2025-08-12",  66,  30),
    ("Estacas Central e Nascente",  "2025-08-13", "2025-10-07", 106,  40),
    ("Estacas Cimas",               "2025-05-13", "2025-09-25",  98,  98),
    ("Escavacao + bandas + ancoragens", "2025-07-08", "2026-04-15", 202, 170),
    ("Viga coroamento + muros",     "2025-07-08", "2025-11-13",  93,  57),
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
    Sobrepoe as fases da obra a um grafico com o tempo no eixo X.

    barra_topo=True (por omissao): BARRA DE FASEAMENTO (tipo Gantt) numa banda
    no topo do proprio grafico — cada fase e um segmento de COR SOLIDA (sem
    faixas de fundo translucidas, que se misturavam onde as fases se
    sobrepunham), com o seu NUMERO fixo. Fases sobrepostas sao empilhadas em
    sub-linhas para nao se misturarem. O eixo Y dos dados e encolhido (domain)
    para abrir espaco; a Gantt partilha o eixo X e alinha automaticamente.
    NAO usar com eixos Y duplos (overlaying) — o domain so afeta um eixo.

    barra_topo=False: modo legado para graficos de eixo Y duplo (ex. Sintese)
    — faixas de fundo translucidas + numero no inicio de cada fase, SEM tocar
    no domain do eixo Y.

    O nome completo esta na legenda a direita (trace fantasma) e no hover.
    Devolve a lista de fases visiveis (para a legenda).
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
    """Escreve, por baixo do grafico, a legenda numero -> nome das fases.
    Usa o numero FIXO global que vem no tuplo (nao re-enumera)."""
    if not visiveis:
        return
    itens = "  ·  ".join(f"**{n}**. {nome}"
                         for _, _, nome, _, n in visiveis)
    st.caption("Fases da obra (planeamento real): " + itens)


def barra_faseamento(dt_min, dt_max, altura=34):
    """
    Desenha uma mini-barra de faseamento (tipo Gantt) para o periodo visivel:
    cada fase e uma barra horizontal na sua propria linha, com o nome legivel,
    sem sobreposicoes. Devolve uma figura plotly compacta para colocar POR CIMA
    do grafico principal — assim as etiquetas das fases saem de dentro do
    grafico e deixam de colidir.
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


def configurar_eixo_tempo(fig, granularidade="Automatico"):
    """
    Define a granularidade das marcas do eixo temporal (X).
    'Mensal' -> 1 marca/mes; 'Quinzenal' -> de 15 em 15 dias;
    'Semanal' -> de 7 em 7 dias; 'Automatico' -> deixa o plotly decidir.
    Marcas mais finas ajudam a ler o faseamento da obra ao nivel a que as
    campanhas existem (~8 em 8 dias).
    """
    if granularidade == "Mensal":
        fig.update_xaxes(dtick="M1", tickformat="%b %Y", tickangle=-30)
    elif granularidade == "Quinzenal":
        fig.update_xaxes(dtick=14 * 24 * 3600 * 1000, tickformat="%d %b",
                         tickangle=-45)
    elif granularidade == "Semanal":
        fig.update_xaxes(dtick=7 * 24 * 3600 * 1000, tickformat="%d %b",
                         tickangle=-45)
    # 'Automatico' -> nao mexe


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


def _casa_edificio(fig, grp, cor, nome):
    """
    Desenha uma casa simples e MERAMENTE ILUSTRATIVA na posicao de um edificio
    vizinho: um bloco (paredes) rematado por um telhado de duas aguas. A posicao
    e a extensao assentam nas coordenadas reais dos alvos do edificio; a forma e
    esquematica e nao representa a geometria real do edificio.
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
        mostrar_caixa = st.checkbox(
            "Caixa de escavacao", value=True,
            help="Desenha o volume escavado abaixo do contorno, com a "
                 "PROFUNDIDADE real de escavacao do projeto (coroamento-fundo "
                 "= 16,3 m). E uma distancia, nao uma cota absoluta — os alvos "
                 "e o projeto usam referenciais de cota diferentes.")
        mostrar_fases = st.checkbox(
            "Fases de escavacao (pisos)", value=False,
            help="Marca dentro da caixa os niveis dos pisos (-1 a -4) como "
                 "planos, a partir das cotas de projeto. Sao distancias abaixo "
                 "do coroamento, invariantes ao referencial.")
        destacar_alarmes = st.checkbox(
            "Destacar alarmes/alertas", value=True,
            help="Marca a vermelho os alvos e setas em alarme e a laranja os "
                 "em alerta, segundo os criterios oficiais recalculados.")
        destacar_sc = st.checkbox(
            "Realcar fachadas da Santa Casa", value=True,
            help="Distingue a fachada frontal (A1-A4, exposta a escavacao) da "
                 "lateral (A5-A8, ao mar).")
        identificar_edif = st.checkbox(
            "Identificar edificios", value=True,
            help="Etiqueta com o nome de cada edificio, flutuando sobre os "
                 "seus alvos.")
        mostrar_casas = st.checkbox(
            "Casas dos edificios (ilustrativo)", value=True,
            help="Desenha cada edificio vizinho como uma casa simples "
                 "(bloco + telhado). Posicao real dos alvos; forma esquematica "
                 "— nao representa a geometria real do edificio.")

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
                name=f"Fundo de escavacao (−{prof:.1f} m do coroamento)",
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
        fachadas = grp["Fachada SC"].to_numpy()

        e_santa_casa = isinstance(chave, str) and "Santa Casa" in chave
        if tipo == "edificio":
            cor = CORES_EDIFICIO.get(chave, "#7f7f7f")
            nome_leg = chave
        else:
            cor = "#ff7f0e"
            nome_leg = f"Contencao — Alcado {etiqueta}"

        # simbolo por fachada da Santa Casa (frontal vs lateral)
        if destacar_sc and e_santa_casa:
            simbolos = ["diamond" if f == "Frente escavacao" else "circle"
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
    c2.metric("Desl. horizontal max. (mm)", f"{np.nanmax(desl_h_all):.1f}")
    c3.metric("Em alarme", n_alarme)
    c4.metric("Em alerta", n_alerta)
    idx = int(np.nanargmax(desl_h_all))
    # leitura frente vs lateral da Santa Casa, se houver dados
    sc = campanha[campanha["Fachada SC"] != ""]
    linha_sc = ""
    if len(sc):
        frente = sc[sc["Fachada SC"] == "Frente escavacao"][COLS["desl_h"]]
        lateral = sc[sc["Fachada SC"] == "Lateral (mar)"][COLS["desl_h"]]
        if len(frente) and len(lateral):
            linha_sc = (f" Na Santa Casa, a fachada frontal (losangos, media "
                        f"{frente.mean():.0f} mm) move-se mais que a lateral "
                        f"(circulos, {lateral.mean():.0f} mm) — coerente com a "
                        f"exposicao direta a escavacao.")
    st.caption(
        f"O alvo mais afetado ({nomes_all[idx]}, {np.nanmax(desl_h_all):.1f} mm) "
        f"pertence a: {edif_all[idx]}. Setas e contornos: vermelho = alarme, "
        f"laranja = alerta, verde = regular (criterios oficiais recalculados). "
        f"A caixa mostra a profundidade real de escavacao (16,3 m, do projeto) "
        f"como distancia abaixo da cortina — os alvos e o projeto usam "
        f"referenciais de cota diferentes, por isso e profundidade, nao cota "
        f"absoluta.{linha_sc}")


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

    tab_perfil, tab_evol = st.tabs(
        ["📐 Perfil deformado", "📈 Evolucao do deslocamento"])
    with tab_perfil:
        st.subheader("Perfil deformado")
        st.caption("Deslocamento acumulado ao longo da profundidade. Base fixa.")
        idx = sorted(set([0, len(datas_inc) // 2, len(datas_inc) - 1]))
        sel = st.multiselect("Leituras", datas_inc,
                             default=[datas_inc[i] for i in idx],
                             format_func=lambda d: pd.to_datetime(d).strftime("%d/%m/%Y"))

        # opcao de sobrepor a geologia de uma sondagem (peca-chave do back-analysis)
        geo_on = st.checkbox("Sobrepor geologia + SPT da sondagem", value=False,
                             help="Mostra a litologia, o nivel freatico e o "
                                  "perfil SPT de uma sondagem no mesmo eixo de "
                                  "profundidade, para relacionar a deformacao "
                                  "com a resistencia do terreno.")
        sond_sel = None
        if geo_on:
            # sugestao por proximidade (inferida das plantas)
            meta = INC_META.get(inc)
            sonds = list(GEO_LITOLOGIA.keys())
            default_idx = 0
            if meta and meta["sondagem"] in sonds:
                default_idx = sonds.index(meta["sondagem"])
            sond_sel = st.selectbox(
                "Sondagem de referencia", sonds, index=default_idx)
            if meta:
                rumo = _rumo_cardeal(meta["azimute"])
                sugerida = meta["sondagem"]
                nota = (f"Sugerida por proximidade: **{sugerida}** "
                        f"(confianca {meta['confianca']}; {inc} fica em "
                        f"{meta['posicao']}). Eixo A+ orientado a "
                        f"{meta['azimute']}° ({rumo}). ")
                if sond_sel != sugerida:
                    nota += f"Estas a ver **{sond_sel}**, diferente da sugerida."
                st.caption(nota)
                st.caption("Associacao inclinometro-sondagem inferida da "
                           "sobreposicao das plantas — a confirmar com a "
                           "instrumentacao.")

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
                       f"back-analysis: ve se o 'joelho' de maior deformacao do "
                       f"perfil coincide com uma subida do SPT (grés a "
                       f"consolidar) ou com o nivel freatico. Onde o "
                       f"inclinometro passa da base da sondagem, nao ha dado "
                       f"geologico.")

    with tab_evol:
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
        fases_vis_inc = None
        if st.session_state.get("mostrar_obra") and len(s_max):
            fases_vis_inc = adicionar_fases_obra(fig2, s_max[COLS["data"]].min(),
                                                 s_max[COLS["data"]].max())
        fig2.update_xaxes(title="Data")
        fig2.update_yaxes(title="Deslocamento (mm)")
        fig2.update_layout(
            height=780,
            legend=dict(title="Serie / fases", orientation="v",
                        yanchor="top", y=1, xanchor="left", x=1.02,
                        font=dict(size=11)),
            margin=dict(r=60, t=60, b=40))
        st.plotly_chart(fig2, use_container_width=True)
        if fases_vis_inc:
            st.caption("A barra no topo mostra o faseamento da obra (cada fase "
                       "com o seu numero e cor, ver legenda a direita). Repara "
                       "se a aceleracao do deslocamento coincide com uma fase.")

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
def _plotar_alcado_esquematico(sub_edi, edi, selecionados, cores_estado):
    """
    Desenha um alcado ESQUEMATICO de um edificio a partir das coordenadas
    reais dos alvos. Usa a coordenada ao longo da fachada (M ou P, conforme
    a orientacao dominante) no eixo horizontal e o Z no eixo vertical. As
    POSICOES RELATIVAS sao fieis aos dados; a ESCALA e esquematica (nao se
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
    horiz_lbl = "Posicao ao longo da fachada (m, relativo)"

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
    """Nome de ficheiro seguro (sem acentos, minusculas) para a foto do edificio."""
    import unicodedata
    txt = unicodedata.normalize("NFKD", str(edi)).encode("ascii", "ignore").decode()
    slug = "".join(c if c.isalnum() else "_" for c in txt).strip("_").lower()
    while "__" in slug:
        slug = slug.replace("__", "_")
    return slug[:60]


def mostrar_localizacao_alvo(sub_edi, edi, selecionados, cores_estado):
    """
    Mostra onde estao os alvos selecionados no edificio. Prioridade:
      1) foto real do relatorio, se existir em fotos_alvos/<slug>.{png,jpg}
      2) alcado esquematico das coordenadas reais (fallback)
    A foto e propriedade do relatorio de instrumentacao (33GRADOS) — creditar.
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
                 caption=f"Localizacao dos alvos — {edi}. "
                         f"Fonte: relatorio de instrumentacao (33GRADOS).")
        return
    # fallback: esquema das coordenadas
    fig = _plotar_alcado_esquematico(sub_edi, edi, selecionados, cores_estado)
    if fig is not None:
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Esquema a partir das coordenadas reais dos alvos (posicoes "
                   "relativas fieis; escala esquematica). O alvo selecionado "
                   "aparece em estrela. Para uma foto real, coloca a imagem em "
                   f"'fotos_alvos/{slug}.png'.")
    else:
        st.caption("Sem coordenadas suficientes para esquematizar este edificio.")


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

    # localizacao fisica dos alvos (foto real ou esquema das coordenadas)
    if st.checkbox("Mostrar localizacao dos alvos no edificio", value=False,
                   help="Foto real do relatorio, se disponivel; caso contrario "
                        "um alcado esquematico a partir das coordenadas."):
        COR_ESTADO = {"Alarme": "#c0140f", "Alerta": "#e67e00",
                      "Regular": "#1f9e55"}
        mostrar_localizacao_alvo(sub, edi, sel, COR_ESTADO)

    # criterio aplicavel a este edificio (para desenhar as linhas de limiar)
    crit_edi, rotulo_edi, ac_edi = criterios_do_alvo(edi)
    Ha, Hm, Va, Vm = crit_edi
    st.caption(f"Criterio aplicado: **{rotulo_edi}**. As linhas tracejadas nos "
               f"graficos marcam os limiares de alerta e alarme.")
    if ac_edi:
        st.warning("Este alcado esta com criterio ASSUMIDO (a confirmar com o "
                   "projeto de contencao).")
    data_rezerag = pd.to_datetime("2025-10-20")

    # --- controlo de zoom temporal (util quando as fases da obra estao ligadas)
    cz1, cz2 = st.columns([1, 2])
    with cz1:
        granul = st.selectbox("Detalhe do eixo temporal",
                              ["Automatico", "Mensal", "Quinzenal", "Semanal"],
                              index=0,
                              help="Marcas mais finas ajudam a ler o faseamento "
                                   "da obra (as campanhas sao ~8 em 8 dias).")
    with cz2:
        # janela de datas para focar um periodo (ex. onde os deslocamentos disparam)
        d_min = pd.to_datetime(sub[COLS["data"]].min()).date()
        d_max = pd.to_datetime(sub[COLS["data"]].max()).date()
        janela = st.slider("Janela temporal", min_value=d_min, max_value=d_max,
                           value=(d_min, d_max), format="DD/MM/YY")
    j0 = pd.to_datetime(janela[0])
    j1 = pd.to_datetime(janela[1])
    sub = sub[(sub[COLS["data"]] >= j0) & (sub[COLS["data"]] <= j1)]

    tab_h, tab_v = st.tabs(
        ["↔ Deslocamento horizontal (H)", "↕ Assentamento vertical (ΔZ)"])
    with tab_h:
        st.markdown(
            "<h4 style='text-align:center; margin-bottom:0; color:#1f2a44;'>"
            "Deslocamento horizontal acumulado (mm)</h4>",
            unsafe_allow_html=True)
        fig = go.Figure()
        for a in sel:
            s = sub[sub[COLS["alvo"]] == a].sort_values(COLS["data"])
            fig.add_trace(go.Scatter(x=s[COLS["data"]], y=s[COLS["desl_h"]],
                                     mode="lines+markers", name=a))
        fases_vis_alv = None
        if st.session_state.get("mostrar_obra") and len(sub):
            fases_vis_alv = adicionar_fases_obra(fig, sub[COLS["data"]].min(),
                                                 sub[COLS["data"]].max())
        fig.add_hline(y=Ha, line_dash="dash", line_color="orange",
                      annotation_text=f"Alerta {Ha}", annotation_position="right")
        fig.add_hline(y=Hm, line_dash="dash", line_color="red",
                      annotation_text=f"Alarme {Hm}", annotation_position="right")
        if e_santa_casa:
            fig.add_vline(x=data_rezerag, line=dict(color="gray", width=1.5, dash="dot"),
                          annotation_text="Re-zeragem A5b–A8b", annotation_position="top")
        fig.update_xaxes(title="Data")
        fig.update_yaxes(title="Desl. horizontal (mm)")
        configurar_eixo_tempo(fig, granul)
        obra_on = st.session_state.get("mostrar_obra")
        fig.update_layout(
            height=780,
            margin=dict(r=60, t=60, b=40),
            legend=dict(orientation="v", yanchor="top", y=1,
                        xanchor="left", x=1.02, font=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    with tab_v:
        st.markdown(
            "<h4 style='text-align:center; margin-bottom:0; color:#1f2a44;'>"
            "Assentamento vertical acumulado, ΔZ (mm)</h4>",
            unsafe_allow_html=True)
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
        configurar_eixo_tempo(fig2, granul)
        obra_on2 = st.session_state.get("mostrar_obra")
        fig2.update_layout(
            height=780,
            margin=dict(r=60, t=60, b=40),
            legend=dict(orientation="v", yanchor="top", y=1,
                        xanchor="left", x=1.02, font=dict(size=11)))
        st.plotly_chart(fig2, use_container_width=True)

    # nota: as fases aparecem agora na legenda a direita de cada grafico
    # (uma entrada por faixa), por isso a caption redundante foi removida.


# =========================================================================
# SEPARADOR 4 — CELULAS DE CARGA
# =========================================================================
def _mostrar_imagens_celula(cel_sel, key_suffix=""):
    """Mostra, num expander, as imagens de localizacao da celula (pecas do
    projeto). Reutilizada no separador Celulas (Inputs) e na Analise."""
    imgs = IMAGENS_CELULAS.get(cel_sel, [])
    if not imgs:
        return
    with st.expander(f"📍 Ver localizacao da celula {cel_sel} (pecas do projeto)"):
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
            st.info("Imagens de localizacao nao encontradas (pasta "
                    "'fotos_celulas/'). Verifica que foram publicadas com a app.")
        loc = LOCALIZACAO_CELULAS.get(cel_sel)
        if loc:
            st.caption(f"Localizacao: {loc[0]}. Imagens das pecas desenhadas do "
                       f"projeto de contencao (JETsj).")


def separador_celulas(dados):
    cc = dados["celulas"]
    if not validar_colunas(cc, [COLS["data"], COLS["celula"], COLS["carga_atual"],
                                COLS["variacao"]], "Celulas de carga"):
        return
    cc = cc.sort_values(COLS["data"])
    st.subheader("Celulas de carga")
    cel = st.selectbox("Celula", sorted(cc[COLS["celula"]].dropna().unique()))
    _mostrar_imagens_celula(cel, key_suffix="_inputs")
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
    st.subheader("Piezometros — cota da agua vs. escavacao")
    st.caption("O eixo vertical e a COTA (m), partilhada com as cotas de "
               "projeto da escavacao/contencao e com o nivel freatico de "
               "repouso. Assim ve-se a que profundidade anda a agua face a "
               "cada piso e ao fundo de escavacao.")

    c1, c2 = st.columns([1, 2])
    with c1:
        p = st.selectbox("Piezometro",
                         sorted(pz[COLS["piezometro"]].dropna().unique()))
    with c2:
        mostrar_pisos = st.checkbox("Cotas dos pisos e fundo de escavacao",
                                    value=True)
        mostrar_nf = st.checkbox("Nivel freatico de repouso (2022)", value=True)

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
                      annotation_text=f"Fundo de escavacao ({COTA_FUNDO_ESCAVACAO:.2f})",
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
                             mode="lines+markers", name="Cota da agua (PZ)",
                             line=dict(color="#2563eb", width=2)))
    fases_vis_pz = []
    if st.session_state.get("mostrar_obra") and len(sub):
        fases_vis_pz = adicionar_fases_obra(fig, sub[COLS["data"]].min(),
                                            sub[COLS["data"]].max())
    fig.update_xaxes(title="Data")
    fig.update_yaxes(title="Cota (m)")
    obra_on_pz = st.session_state.get("mostrar_obra") and fases_vis_pz
    fig.update_layout(
        height=760 if obra_on_pz else 560,
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
            f"**Leitura:** a agua no {p} desceu de **{c_ini:.2f}** para "
            f"**{c_fim:.2f} m** ({desc:+.2f} m) no periodo monitorizado. "
            f"O fundo de escavacao (**{COTA_FUNDO_ESCAVACAO:.2f} m**) fica "
            f"{c_fim - COTA_FUNDO_ESCAVACAO:.1f} m abaixo da agua atual e "
            f"~{nf_med - COTA_FUNDO_ESCAVACAO:.0f} m abaixo do nivel freatico "
            f"de repouso de 2022 (~{nf_med:.0f} m). A descida acompanha o "
            f"avanco da escavacao — coerente com rebaixamento induzido.")
    st.caption("Cotas de projeto: escavacao e contencao periferica (JETsj, "
               "PRO/2023/368). NF de repouso: piezometros das sondagens "
               "(ENGGEO, Quadro III, 24/11/2022). Possivel melhoria: sobrepor "
               "precipitacao diaria para separar a resposta a pluviosidade do "
               "efeito da escavacao.")


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
    "Gres (C1As)": "#a9c47f",
    "Calcario (C1A)": "#5c8a45",
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
    st.caption("gama: peso volumico | c': coesao | fi': angulo de atrito | "
               "E': modulo de deformabilidade. Zonas ZG3-ZG1 (rocha) com c' e E' "
               "em MPa/GPa; ZG6-ZG4 (solo/grés brando) em kPa/MPa. Nota: as tres "
               "zonas de nega (ZG3-ZG1) distinguem-se pelo RQD, que o SPT sozinho "
               "nao mede — por isso o perfil agrupa-as como 'nega'.")



# =========================================================================
# SEPARADOR 8 — CRONOGRAMA DA OBRA
# =========================================================================
def separador_obra(dados):
    st.subheader("Cronograma da obra — planeado vs. executado")
    st.caption("Comparacao entre o planeamento CONTRATUAL (previsto) e o "
               "EXECUTADO (real, do plano de trabalhos impactado da "
               "Aquatecnica, 22/04/2026). A janela de instrumentacao esta "
               "assinalada, para relacionar as fases com a monitorizacao.")

    vista = st.radio("Cronograma a mostrar",
                     ["Executado (real)", "Comparacao real vs. previsto"],
                     horizontal=True)

    fig = go.Figure()
    nomes = [f[0] for f in FASES_COMPARACAO]

    if vista == "Executado (real)":
        for i, (nome, ini, fim, dur_r, dur_c) in enumerate(FASES_COMPARACAO):
            cor = CORES_FASES[i % len(CORES_FASES)]
            fig.add_trace(go.Scatter(
                x=[pd.to_datetime(ini), pd.to_datetime(fim)],
                y=[nome, nome], mode="lines", line=dict(color=cor, width=16),
                hovertemplate=f"{nome}<br>{ini} a {fim}<br>{dur_r} dias "
                              f"(previsto {dur_c})<extra></extra>",
                showlegend=False))
    else:
        # duas barras por fase: real (cor) e contratual (cinza), deslocadas
        for i, (nome, ini, fim, dur_r, dur_c) in enumerate(FASES_COMPARACAO):
            cor = CORES_FASES[i % len(CORES_FASES)]
            t0 = pd.to_datetime(ini)
            t1_real = pd.to_datetime(fim)
            t1_prev = t0 + pd.Timedelta(days=dur_c)   # fim previsto ancorado ao inicio real
            # barra real (em cima)
            fig.add_trace(go.Scatter(
                x=[t0, t1_real], y=[f"{nome} ", f"{nome} "], mode="lines",
                line=dict(color=cor, width=11),
                hovertemplate=f"REAL: {dur_r} dias<extra></extra>",
                showlegend=False))
            # barra prevista (em baixo, cinza)
            fig.add_trace(go.Scatter(
                x=[t0, t1_prev], y=[f" {nome}", f" {nome}"], mode="lines",
                line=dict(color="rgba(120,120,120,0.6)", width=11),
                hovertemplate=f"PREVISTO: {dur_c} dias<extra></extra>",
                showlegend=False))

    # faixa da janela de instrumentacao
    alvos = dados.get("alvos")
    if alvos is not None and not alvos.empty and COLS["data"] in alvos.columns:
        d0 = alvos[COLS["data"]].min()
        d1 = alvos[COLS["data"]].max()
        fig.add_vrect(x0=d0, x1=d1, fillcolor="crimson", opacity=0.10,
                      line_width=0, annotation_text="Instrumentacao",
                      annotation_position="top left")

    fig.update_xaxes(title="Data")
    fig.update_layout(height=480, margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig, use_container_width=True)

    # ---- analise de desempenho de prazo -------------------------------
    st.markdown("#### Desempenho face ao prazo")
    linhas = []
    for nome, ini, fim, dur_r, dur_c in FASES_COMPARACAO:
        desvio = dur_r - dur_c
        pct = (desvio / dur_c * 100) if dur_c else 0
        linhas.append({
            "Fase": nome,
            "Previsto (dias)": dur_c,
            "Executado (dias)": dur_r,
            "Desvio (dias)": f"{desvio:+d}",
            "Desvio (%)": f"{pct:+.0f}%",
        })
    df_des = pd.DataFrame(linhas)
    st.dataframe(df_des, use_container_width=True, hide_index=True)

    # sintese do desvio global das fases estruturais
    tot_r = sum(f[3] for f in FASES_COMPARACAO)
    tot_c = sum(f[4] for f in FASES_COMPARACAO)
    pior = max(FASES_COMPARACAO, key=lambda f: f[3] - f[4])
    st.markdown(
        f"**Leitura:** somando as fases estruturais, foram executadas em "
        f"**{tot_r} dias** contra **{tot_c} previstos** "
        f"({(tot_r-tot_c)/tot_c*100:+.0f}%). O maior desvio foi em "
        f"**{pior[0]}** (+{pior[3]-pior[4]} dias). Os desvios de prazo sao "
        f"relevantes para o back-analysis: fases que se prolongaram "
        f"mantiveram o macico desconfinado mais tempo, o que ajuda a "
        f"interpretar a evolucao da deformacao.")
    st.caption("Nota: a duracao prevista e a contratual (sem impacto); a "
               "executada e a impactada. As datas de inicio reais e "
               "contratuais coincidem na maioria das macro-fases — o desvio "
               "manifesta-se na duracao, nao no arranque.")


# =========================================================================
# SEPARADOR 0 — PAGINA INICIAL (HOME)
# =========================================================================
def separador_pressupostos(dados):
    """
    Reune num so sitio TODOS os pressupostos, inferencias e limitacoes da
    ferramenta e dos dados — o que e medido vs. o que e assumido, e o estado de
    confirmacao de cada um. Assumir as limitacoes explicitamente e uma forca:
    antecipa as perguntas dificeis da defesa.
    """
    st.subheader("Pressupostos, inferencias e qualidade dos dados")
    st.caption("Transparencia sobre o que e MEDIDO e o que e ASSUMIDO. Cada "
               "pressuposto tem a sua origem e o estado de confirmacao. Esta "
               "seccao reune, num so sitio, as ressalvas assinaladas ao longo "
               "da aplicacao.")

    st.markdown("#### Pressupostos e inferencias da ferramenta")
    press = pd.DataFrame([
        {"Item": "Referencial de cota (3D)",
         "O que se assume": "A caixa de escavacao usa a PROFUNDIDADE real "
                            "(16,3 m), nao a cota absoluta — os sistemas de cota "
                            "dos alvos (Z 42-63) e do projeto (4,5-24,6) nao "
                            "estao relacionados.",
         "Estado": "A confirmar (falta 1 par de cotas nos 2 sistemas)"},
        {"Item": "Associacao inclinometro-sondagem",
         "O que se assume": "I1<->SC8, I2<->SC9, I3<->SC6, inferido por "
                            "sobreposicao das plantas de prospecao e de "
                            "instrumentacao.",
         "Estado": "A confirmar com a equipa de instrumentacao"},
        {"Item": "Criterio do alcado DE",
         "O que se assume": "Classificado como contencao 17 m; os deslocamentos "
                            "sao ~0, pelo que os dados nao distinguem 17 de 24 m.",
         "Estado": "A confirmar com o projeto de contencao"},
        {"Item": "Faseamento da obra",
         "O que se assume": "Datas do cronograma PREVISTO, nao do executado.",
         "Estado": "Substituivel pelas datas reais de obra"},
        {"Item": "Zona geotecnica no perfil SPT",
         "O que se assume": "A cor da zona (ZG) e a provavel pelo valor de N; o "
                            "relatorio nao define fronteiras de zona por "
                            "profundidade.",
         "Estado": "Leitura qualitativa (nao fronteiras reais)"},
    ])
    st.dataframe(press, use_container_width=True, hide_index=True)

    st.markdown("#### Tratamento de dados (criterios aplicados)")
    trat = pd.DataFrame([
        {"Item": "Estado dos alvos",
         "Tratamento": "Recalculado de ΔH/ΔV com criterios oficiais do "
                       "relatorio; auditado contra a coluna do Excel (467/467, "
                       "100%)."},
        {"Item": "Alvos A5-A8 -> A5b-A8b",
         "Tratamento": "Tratados como series separadas; os 'b' foram re-zerados "
                       "em 20/10/2025 e os acumulados nao sao comparaveis "
                       "diretamente com A1-A4."},
        {"Item": "Fachadas da Santa Casa",
         "Tratamento": "Distinguidas frontal (A1-A4) e lateral (A5-A8) pela "
                       "posicao em planta."},
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
            st.markdown("#### Questoes de qualidade dos dados (do relatorio)")
            st.caption("Ressalvas identificadas no proprio modelo de dados, "
                       "transcritas da folha Qualidade_Dados.")
            st.dataframe(qd, use_container_width=True, hide_index=True)
    except Exception:
        pass

    st.info("Nota metodologica: estas limitacoes nao invalidam a analise — "
            "delimitam o seu alcance. As conclusoes centrais (relacao entre "
            "escavacao, rebaixamento da agua e deformacao; consolidacao "
            "crescente do gres) assentam em dados medidos, nao nos "
            "pressupostos acima.")


@st.cache_data
def _carregar_terreno():
    """Le o modelo digital de terreno (MDT) extraido do levantamento topografico."""
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
    Coloracao ILUSTRATIVA e QUALITATIVA dos lados do recinto de escavacao
    conforme o nivel de movimento dos alvos, associado por ORIENTACAO.

    Como o terreno (coords nacionais) e os alvos (coords locais) estao em
    referenciais diferentes, nao ha correspondencia ponto-a-ponto. A associacao
    e feita por lado/orientacao: divide-se o contorno em 4 quadrantes (N, S, E,
    O) relativos ao centro, e a cada quadrante atribui-se a cor do pior estado
    dos alvos cuja orientacao (relativa ao centro dos alvos) corresponde a esse
    lado. E uma leitura de tendencia, nao uma medida exata.

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
    Calcula os parametros de alinhamento CAIXA-ENVOLVENTE entre o referencial
    dos alvos (local, M~5000/P~5050) e o do terreno (nacional, X~-110900/
    Y~-106300). Centra o conjunto dos alvos no centro do terreno e aplica uma
    escala UNICA proporcional (a menor razao de extensao, para nao distorcer).
    Devolve um dicionario com centros, escala e extensao — a transformacao em si
    e aplicada por _transformar_xy, que aceita ajustes finos do utilizador.

    NOTA IMPORTANTE (honestidade para a tese): este alinhamento e ILUSTRATIVO.
    Nao ha pontos de controlo comuns aos dois referenciais, por isso a posicao
    de cada alvo sobre o terreno e APROXIMADA (erro tipico de metros). Serve
    para ver a tendencia do movimento sobre o relevo, nao para medir. A precisao
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
    """Aplica a transformacao caixa-envolvente (par de _bbox_alinhamento) a
    coordenadas de alvos, com ajustes finos opcionais (rotacao em torno do
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
    """Transforma um VETOR de deslocamento (sem translacao nem centro): so
    aplica espelhamento, rotacao e escala. Para as setas de movimento assentarem
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
    """Interpola a cota da superficie do terreno nos pontos (X,Y) — para
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
    """Como _casa_edificio, mas recebe coordenadas JA TRANSFORMADAS (X,Y no
    referencial do terreno) e a cota de base do terreno — para desenhar a casa
    ilustrativa do edificio vizinho sobre a superficie fundida."""
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
    3D do terreno REAL, a partir do levantamento topografico (MDT triangulado),
    no referencial da obra e com cotas Z reais. Mostra a superficie do terreno,
    a linha de escavacao (cota 4,55) e o volume escavado. Ao contrario do 3D dos
    alvos (referencial local, cotas relativas), este assenta em cotas absolutas.
    """
    st.subheader("Terreno 3D — modelo do levantamento topografico")
    st.caption("Superficie real do terreno (modelo digital triangulado) do "
               "levantamento topografico, no referencial da obra e com COTAS "
               "REAIS. Mostra a topografia original e o volume a escavar ate a "
               "cota de fundo (4,55 m). Complementa o 3D dos alvos — que usa um "
               "referencial local e cotas relativas — com geometria absoluta.")

    terreno = _carregar_terreno()
    if terreno is None:
        st.info("Modelo de terreno nao disponivel (falta terreno_mdt.json).")
        return

    import numpy as np
    faces = np.array(terreno["faces"])           # (N,3,3)
    z_fundo = terreno.get("cota_fundo", 4.55)
    z_max_terreno = float(faces.reshape(-1, 3)[:, 2].max())  # cota mais alta

    c1, c2 = st.columns(2)
    with c1:
        mostrar_escav = st.checkbox("Mostrar volume de escavacao", value=True)
        mostrar_curvas = st.checkbox("Curvas de nivel", value=False,
                                     help="Curvas de nivel do levantamento, que "
                                          "dao a leitura do relevo (alcados).")
        colorir_lados = st.checkbox(
            "Colorir alcados por movimento (ilustrativo)", value=False,
            help="Colore os lados do recinto conforme o nivel de movimento dos "
                 "alvos proximos, por orientacao. E QUALITATIVO e ilustrativo: "
                 "o terreno e os alvos estao em referenciais diferentes, pelo "
                 "que a associacao e por lado, nao por ponto exato.")
    with c2:
        exagero = st.slider("Exagero vertical", 1.0, 4.0, 1.5, 0.5,
                            help="Amplia a escala vertical para realcar o "
                                 "relevo. 1.0 = escala real.")
        # filtro por fase de escavacao: ate que piso ja se escavou
        fases_esc = ["Terreno original (sem escavar)"] + \
            [f"Ate {nome} (cota {cota:.2f})" for nome, cota in COTAS_PISOS
             if cota < z_max_terreno] + \
            [f"Escavacao completa (fundo {z_fundo})"]
        fase_esc = st.selectbox(
            "Fase de escavacao a representar", fases_esc,
            index=len(fases_esc) - 1,
            help="Mostra a escavacao ate a cota do piso correspondente. As "
                 "cotas sao as do projeto; a escavacao real progride por fases.")

    # cota de escavacao correspondente a fase escolhida
    if fase_esc.startswith("Terreno original"):
        cota_escav_fase = None                 # nao escavado
    elif fase_esc.startswith("Escavacao completa"):
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
                name=f"Escavacao ate cota {cota_escav_fase:.2f}"))
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
                name="Curvas de nivel", hoverinfo="skip"))

    fig.update_layout(
        height=640,
        scene=dict(
            xaxis_title="M (m)", yaxis_title="P (m)", zaxis_title="Cota (m)",
            aspectmode="data"),
        margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig, use_container_width=True)

    st.caption(f"Superficie do terreno entre as cotas {z.min():.1f} e "
               f"{z.max():.1f} m; fundo de escavacao a {z_fundo} m — "
               f"aproximadamente {z.max()-z_fundo:.0f} m de altura escavada no "
               f"ponto mais alto. Fonte: levantamento topografico (DXF). Notas: "
               f"o exagero vertical e apenas visual; o filtro de fase mostra a "
               f"escavacao ate a cota do piso do projeto (a progressao real e "
               f"por fases). A coloracao dos alcados por movimento, quando "
               f"ativa, e ILUSTRATIVA e qualitativa — associa o estado dos "
               f"alvos ao lado do recinto por orientacao, nao por coordenada "
               f"exata, porque terreno e alvos usam referenciais distintos.")


def _detetar_cantos_contorno(pts, n_cantos=4):
    """Deteta os cantos de um contorno fechado (angulo interno mais fechado),
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
    Desenha, de forma ESQUEMATICA mas INFORMADA PELO PROJETO, as solucoes de
    contencao periferica: viga de coroamento/distribuicao, bandas de laje e
    escoramentos. Baseado nas pecas desenhadas da JETsj (EDN-JET-ZZ-ZZ-DR-U-
    0021..0027): cotas dos pisos confirmadas com o projeto, cores das lajes
    conforme a legenda oficial, e escoras HORIZONTAIS entre lados opostos as
    cotas dos pisos (como nos cortes tipo), nao diagonais de canto.

    ATENCAO — HONESTIDADE: a FORMA em planta (o contorno curvo real, a geometria
    exata de cada banda) nao foi extraida do PDF — a sua reproducao vetorial nao
    era fiavel. Por isso as bandas sao aneis perimetrais aproximados sobre o
    contorno disponivel, e as escoras sao troços representativos. As COTAS e as
    CORES sao reais; a forma e esquematica. A representacao rigorosa esta nas
    pecas desenhadas do projeto e nas VISTAS 3D do projeto (mostradas no fim
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
        name="Viga de coroamento/distribuicao (projeto: cota real)",
        hovertemplate="Viga de coroamento/distribuicao<br>"
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
        name="Escoras metalicas provisorias (esquematico)",
        hovertemplate="Escora metalica provisoria (horizontal)<br>"
                      "posicao esquematica<extra></extra>"))


def separador_terreno_alvos_3d(dados):
    """
    SEPARADOR FUNDIDO: terreno real (MDT) + alvos com movimento drapejados
    sobre a superficie. Junta o melhor dos dois 3D antigos — o relevo real com
    curvas de nivel e volume de escavacao, e os alvos coloridos por estado com
    setas de deslocamento, casas dos edificios e destaque da Santa Casa.

    Referenciais diferentes: os alvos sao alinhados ao terreno por CAIXA-
    ENVOLVENTE (ilustrativo, ver _bbox_alinhamento) e assentes na cota do
    terreno (drapejados). A posicao e APROXIMADA; a precisao esta no Alvos 2D.
    """
    import numpy as np
    st.subheader("Terreno + Alvos 3D — movimento sobre o relevo real")
    st.caption("Funde o modelo real do terreno (levantamento topografico, com "
               "curvas de nivel e volume de escavacao) com os alvos e o seu "
               "movimento. Os alvos estao num referencial diferente do terreno, "
               "por isso sao POUSADOS sobre a superficie por alinhamento "
               "aproximado (caixa-envolvente) — a posicao e ILUSTRATIVA, para "
               "ler a tendencia do movimento sobre o relevo. A medicao precisa "
               "esta no separador Alvos (2D).")

    terreno = _carregar_terreno()
    if terreno is None:
        st.info("Modelo de terreno nao disponivel (falta terreno_mdt.json).")
        return
    alvos = dados.get("alvos")
    if alvos is None or alvos.empty:
        st.info("Sem dados de alvos.")
        return

    par = _bbox_alinhamento(dados, terreno)
    if par is None:
        st.info("Nao foi possivel alinhar alvos e terreno.")
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
        fator = st.slider("Amplificacao do deslocamento", 50, 2000, 500, 50,
                          key="fus_fator",
                          help="Deslocamentos milimetricos sobre coordenadas em "
                               "metros; amplia-se para se verem.")
        exagero = st.slider("Exagero vertical do terreno", 1.0, 4.0, 1.5, 0.5,
                            key="fus_exag",
                            help="Amplia so a escala vertical do relevo.")
    with col_b:
        mostrar_curvas = st.checkbox("Curvas de nivel", value=True,
                                     key="fus_curvas")
        mostrar_escav = st.checkbox("Volume de escavacao", value=True,
                                    key="fus_escav")
        mostrar_casas = st.checkbox("Casas dos edificios", value=True,
                                    key="fus_casas")
        identificar_edif = st.checkbox("Identificar edificios", value=True,
                                       key="fus_idedif")
        mostrar_contencao = st.checkbox(
            "Elementos de contencao (esquematico)", value=False,
            key="fus_contencao",
            help="Desenha, de forma ESQUEMATICA e ILUSTRATIVA, a viga de "
                 "coroamento, as bandas de laje (as cotas reais dos pisos) e "
                 "os escoramentos de canto. A geometria de projeto destes "
                 "elementos nao consta dos dados — a forma e assumida, so as "
                 "cotas das lajes/viga sao reais. Os escoramentos sao "
                 "totalmente ilustrativos.")
    with col_c:
        destacar_alarmes = st.checkbox("Destacar alarmes/alertas", value=True,
                                       key="fus_alarmes")
        destacar_sc = st.checkbox("Realcar Santa Casa", value=True,
                                  key="fus_sc")
        # ajustes finos do alinhamento (o utilizador valida a olho)
        with st.expander("Ajuste fino do alinhamento"):
            rot = st.select_slider("Rotacao", [0, 90, 180, 270], value=0,
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
            name="Paredes de escavacao", hoverinfo="skip"))
        # contorno do fundo (linha fechada, cota de fundo)
        fig.add_trace(go.Scatter3d(
            x=list(ex) + [ex[0]], y=list(ey) + [ey[0]], z=[zf] * (n + 1),
            mode="lines", line=dict(color="#8B4513", width=4),
            name=f"Fundo de escavacao ({cota_escav_fase:.2f} m)"))
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
                name="Curvas de nivel", hoverinfo="skip"))

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
        fachadas = grp["Fachada SC"].to_numpy()
        e_santa_casa = isinstance(chave, str) and "Santa Casa" in chave

        if tipo == "edificio":
            cor = CORES_EDIFICIO.get(chave, "#7f7f7f"); nome_leg = chave
        else:
            cor = "#ff7f0e"; nome_leg = f"Contencao — Alcado {etiqueta}"

        if destacar_sc and e_santa_casa:
            simbolos = ["diamond" if f == "Frente escavacao" else "circle"
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
    c2.metric("Desl. horizontal max. (mm)", f"{np.nanmax(desl_h_all):.1f}")
    c3.metric("Em alarme", n_alarme)
    c4.metric("Em alerta", n_alerta)
    idx = int(np.nanargmax(desl_h_all))
    st.caption(
        f"O alvo mais afetado ({nomes_all[idx]}, {np.nanmax(desl_h_all):.1f} mm) "
        f"esta drapejado sobre o relevo real. Vermelho = alarme, laranja = "
        f"alerta, verde = regular. IMPORTANTE: a posicao dos alvos sobre o "
        f"terreno e aproximada (alinhamento por caixa-envolvente entre "
        f"referenciais diferentes) — para leitura tendencial do movimento sobre "
        f"o relevo, nao metrica. A medicao precisa esta no separador Alvos (2D). "
        f"Se algum alvo cair no lado errado, usa o 'Ajuste fino do alinhamento'.")
    if mostrar_contencao:
        st.warning(
            "⚠ Os elementos de contencao (viga de coroamento/distribuicao, "
            "bandas de laje, escoras metalicas) sao ESQUEMATICOS mas INFORMADOS "
            "PELO PROJETO: as cotas dos pisos e as cores das lajes seguem as "
            "pecas desenhadas da JETsj (EDN-JET-ZZ-ZZ-DR-U-0021..0027), e as "
            "escoras sao horizontais entre lados (como nos cortes tipo). A FORMA "
            "em planta (contorno curvo, geometria exata de cada banda) e "
            "aproximada — nao foi extraida do projeto. Para a representacao "
            "rigorosa, ver as Vistas 3D do projeto no fim deste separador.")

    # ---- Vistas 3D do proprio projeto (referencia fiel, JETsj) ----
    with st.expander("📐 Vistas 3D do projeto de contencao (JETsj) — referencia"):
        st.caption("Imagens das pecas desenhadas do projeto de contencao "
                   "periferica (JETsj, pranchas EDN-JET-ZZ-ZZ-DR-U-0002 a 0005). "
                   "Sao a representacao AUTORITATIVA e rigorosa da solucao — "
                   "cortina de estacas, bandas de laje por cota, ancoragens, "
                   "escoras metalicas e faseamento. O 3D interativo acima e uma "
                   "leitura do movimento dos alvos; estas vistas sao a geometria "
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
            st.info("Imagens das vistas 3D do projeto nao encontradas "
                    "(pasta 'projeto_vistas/'). Verifica que foi publicada "
                    "junto com a app.")


def separador_analise(dados):
    """
    ANALISE DESCRITIVA das principais movimentacoes observadas. Quantifica o
    QUE aconteceu (quanto, onde, quando/a que ritmo, a que profundidade), de
    forma a servir de base defensavel para a interpretacao do PORQUE (feita
    pelo autor no texto da tese, nao pela app).

    Principio: a app organiza e quantifica a evidencia; NAO afirma causas. As
    notas orientadoras apontam o que cruzar (faseamento, geologia, agua), mas a
    conclusao causal e do autor.
    """
    import numpy as np
    st.subheader("Analise — principais movimentacoes observadas")
    st.caption("Quantificacao das movimentacoes medidas pela instrumentacao, "
               "para fundamentar a interpretacao. Esta seccao descreve O QUE "
               "aconteceu (quanto, onde, a que ritmo, a que profundidade); a "
               "interpretacao do PORQUE (geologia, faseamento, dimensionamento) "
               "e do autor, apoiada nesta evidencia e nos outros separadores.")

    alvos = dados.get("alvos")
    if alvos is None or alvos.empty:
        st.info("Sem dados de alvos.")
        return
    alvos = alvos.copy()
    alvos[COLS["data"]] = pd.to_datetime(alvos[COLS["data"]], errors="coerce")
    ult_data = alvos[COLS["data"]].max()
    ult = anexar_estado_calculado(alvos[alvos[COLS["data"]] == ult_data].copy())

    # ============ (a) QUANTO ============
    st.markdown("#### 1. Quanto — magnitude das movimentacoes")
    n_alarme = int((ult["Estado calculado"] == "Alarme").sum())
    n_alerta = int((ult["Estado calculado"] == "Alerta").sum())
    dh = ult[COLS["desl_h"]]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alvos monitorizados", ult[COLS["alvo"]].nunique())
    c2.metric("Desl. H maximo (mm)", f"{dh.max():.1f}")
    c3.metric("Em alarme", n_alarme)
    c4.metric("Em alerta", n_alerta)

    top = ult.nlargest(10, COLS["desl_h"])[
        [COLS["alvo"], COLS["edificio"], COLS["desl_h"], COLS["dZ"],
         "Estado calculado"]].copy()
    top.columns = ["Alvo", "Edificio / elemento", "Desl. H (mm)",
                   "Desl. V ΔZ (mm)", "Estado"]
    top["Desl. H (mm)"] = top["Desl. H (mm)"].round(1)
    top["Desl. V ΔZ (mm)"] = top["Desl. V ΔZ (mm)"].round(1)
    st.markdown("**Ranking dos 10 alvos mais deslocados** (ultima campanha, "
                f"{ult_data.strftime('%d/%m/%Y')}):")
    st.dataframe(top, use_container_width=True, hide_index=True)

    top1 = top.iloc[0]
    st.caption(
        f"O maior deslocamento horizontal e do alvo {top1['Alvo']} "
        f"({top1['Desl. H (mm)']:.1f} mm), no elemento «{top1['Edificio / elemento']}». "
        f"Nota para a interpretacao: verifique se os alvos no topo do ranking "
        f"pertencem ao mesmo elemento/zona — isso indica movimentacao "
        f"concentrada, nao dispersa.")

    # ============ (b) ONDE ============
    st.divider()
    st.markdown("#### 2. Onde — concentracao espacial")
    g = ult.groupby(COLS["edificio"])[COLS["desl_h"]].agg(
        media="mean", maximo="max", n="count").reset_index()
    g = g.sort_values("maximo", ascending=False)
    g_plot = g.head(12)
    fig_b = go.Figure()
    fig_b.add_trace(go.Bar(
        y=g_plot[COLS["edificio"]], x=g_plot["maximo"].round(1),
        orientation="h", name="Maximo",
        marker=dict(color="#c0140f"),
        text=g_plot["maximo"].round(1), textposition="auto"))
    fig_b.add_trace(go.Bar(
        y=g_plot[COLS["edificio"]], x=g_plot["media"].round(1),
        orientation="h", name="Media",
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
    txt_onde = (f"O elemento com maior movimentacao e «{lider[COLS['edificio']]}» "
                f"(max {lider['maximo']:.1f} mm, media {lider['media']:.1f} mm).")
    if seg is not None:
        racio = lider["maximo"] / seg["maximo"] if seg["maximo"] else float("nan")
        txt_onde += (f" O segundo é «{seg[COLS['edificio']]}» (max "
                     f"{seg['maximo']:.1f} mm) — o líder move-se {racio:.1f}x "
                     f"mais, o que indica concentracao espacial e nao "
                     f"movimento uniforme do recinto.")
    st.caption(txt_onde + " Nota para a interpretacao: cruze a zona critica "
               "com a geologia (separador Geologia) e com a orientacao face "
               "a escavacao (a fachada voltada para a escavacao move mais).")

    # ============ (c) QUANDO / RITMO ============
    st.divider()
    st.markdown("#### 3. Quando e a que ritmo — velocidade de movimentacao")
    st.caption("O deslocamento total nao distingue um movimento lento e "
               "continuo de um surto rapido. A velocidade (mm/dia entre "
               "campanhas) revela QUANDO o movimento acelerou — o momento a "
               "cruzar com o avanco da obra.")
    alvos_disp = sorted(ult.nlargest(15, COLS["desl_h"])[COLS["alvo"]]
                        .astype(str).tolist())
    alvo_sel = st.selectbox("Alvo a analisar (velocidade)", alvos_disp,
                            key="an_alvo_vel")
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
        st.caption(
            f"O alvo {alvo_sel} teve velocidade maxima de {pico['vel']:.2f} "
            f"mm/dia no intervalo terminado a "
            f"{pico[COLS['data']].strftime('%d/%m/%Y')}. Nota para a "
            f"interpretacao: veja no separador Sintese o que a obra estava a "
            f"fazer nessa data (fase, cota de escavacao atingida, nivel de "
            f"agua) — a coincidencia temporal e a base do argumento causal.")

    # ============ (d) PROFUNDIDADE ============
    st.divider()
    st.markdown("#### 4. A que profundidade — perfil inclinometrico")
    st.caption("A profundidade do deslocamento maximo indica o mecanismo: um "
               "maximo perto da superficie sugere um comportamento diferente "
               "de um maximo em profundidade (junto ao pe da cortina ou a uma "
               "camada especifica).")
    res = dados.get("resumo")
    if res is not None and not res.empty:
        res = res.copy()
        res[COLS["data"]] = pd.to_datetime(res[COLS["data"]], errors="coerce")
        linhas = []
        for inc in sorted(res[COLS["inclinometro"]].dropna().unique()):
            si = res[res[COLS["inclinometro"]] == inc].sort_values(
                COLS["data"]).iloc[-1]
            linhas.append({
                "Inclinometro": inc,
                "Desl. max (mm)": round(float(si[COLS["desl_max_global"]]), 1),
                "Profundidade do max (m)": round(float(si[COLS["prof_do_max"]]), 1),
            })
        df_d = pd.DataFrame(linhas)
        st.dataframe(df_d, use_container_width=True, hide_index=True)
        # comentario neutro sobre a dispersao das profundidades
        profs = df_d["Profundidade do max (m)"]
        if profs.max() - profs.min() > 5:
            st.caption(
                "Nota para a interpretacao: os inclinometros apresentam o "
                "maximo a profundidades diferentes — o mecanismo de deformacao "
                "nao e o mesmo em todo o recinto. Cruze cada profundidade com "
                "o perfil geologico local (separador Geologia) e com a cota de "
                "escavacao para distinguir os mecanismos.")
        else:
            st.caption(
                "Nota para a interpretacao: os maximos ocorrem a profundidades "
                "semelhantes — sugere um mecanismo de deformacao comum. Cruze "
                "com o perfil geologico e a cota de escavacao.")
    else:
        st.info("Sem dados de inclinometros para o perfil de profundidade.")

    # ============ (5) OBSERVADO vs. PREVISTO ============
    st.divider()
    st.markdown("#### 5. Observado vs. Previsto — comparacao com o projeto")
    st.caption(
        "Compara o deslocamento horizontal OBSERVADO com a ESTIMATIVA DE "
        "PROJETO (memoria descritiva JETsj: ~20 mm na cortina poente/norte, "
        "~10 mm na nascente/sul — valor de CALCULO para a ultima fase de "
        "escavacao, nao um limite) e com os criterios de alerta/alarme. "
        "A coluna «Origem zona» distingue: «confirmada» = orientacao verificada "
        "em planta (edificios vizinhos Santa Casa, Cimas e Clinica, todos com "
        "componente poente ou norte -> 20 mm); «inferida» = deduzida da "
        "geometria dos alvos, A CONFIRMAR, pois o referencial (M,P) esta rodado "
        "face ao norte real. Nota: os alvos de edificios vizinhos medem o "
        "movimento do EDIFICIO, relacionado mas nao identico a deformacao da "
        "propria cortina.")

    Mc = ult[COLS["M0"]].mean()
    Pc = ult[COLS["P0"]].mean()
    linhas = []
    for chave, grp in ult.groupby(COLS["edificio"]):
        # 1) tentar mapeamento CONFIRMADO (edificios vizinhos) por nome
        zona_key = None
        origem = "inferida"
        for alvo_nome, zk in ZONA_FIXA_GRUPO.items():
            if isinstance(chave, str) and alvo_nome.lower() in chave.lower():
                zona_key = zk
                origem = "confirmada"
                break
        # 2) se nao for conhecido, inferir pela geometria (a confirmar).
        #    NOTA: o referencial (M,P) esta rodado face ao norte real, por isso
        #    esta inferencia so vale como aproximacao inicial para os alcados.
        if zona_key is None:
            M = grp[COLS["M0"]].mean()
            P = grp[COLS["P0"]].mean()
            poente = M < Mc
            norte = P > Pc
            zona_key = "poente_norte" if (poente or norte) else "nascente_sul"
        prev = DEFORM_PROJETO[zona_key]
        zona = "poente/norte" if zona_key == "poente_norte" else "nascente/sul"
        obs = grp[COLS["desl_h"]].max()
        dif = obs - prev
        pct = (dif / prev * 100) if prev else float("nan")
        linhas.append({
            "Grupo": chave,
            "Zona": zona,
            "Origem zona": origem,
            "Previsto (mm)": prev,
            "Observado max (mm)": round(obs, 1),
            "Diferenca (mm)": round(dif, 1),
            "Excesso (%)": round(pct, 0),
        })
    df5 = pd.DataFrame(linhas).sort_values("Observado max (mm)", ascending=False)
    st.dataframe(df5, use_container_width=True, hide_index=True)

    # grafico previsto vs observado (top por observado)
    top5 = df5.head(12)
    fig5 = go.Figure()
    fig5.add_trace(go.Bar(
        y=top5["Grupo"], x=top5["Previsto (mm)"], orientation="h",
        name="Previsto (projeto)", marker=dict(color="#9aa7b5")))
    fig5.add_trace(go.Bar(
        y=top5["Grupo"], x=top5["Observado max (mm)"], orientation="h",
        name="Observado (max)", marker=dict(color="#c0140f")))
    # linhas de referencia dos criterios da contencao
    for xval, txt, cor in [(40, "Alarme contencao (40)", "#8B0000")]:
        fig5.add_vline(x=xval, line=dict(color=cor, width=1, dash="dash"),
                       annotation_text=txt, annotation_position="top")
    fig5.update_layout(
        height=460, barmode="group",
        xaxis=dict(title="Desl. horizontal (mm)"),
        yaxis=dict(autorange="reversed"),
        margin=dict(l=10, r=10, t=40, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig5, use_container_width=True)

    # destaque factual do caso critico, sem afirmar causa
    pior = df5.iloc[0]
    st.caption(
        f"O grupo com maior excedencia e «{pior['Grupo']}»: observado "
        f"{pior['Observado max (mm)']:.1f} mm vs. previsto "
        f"{pior['Previsto (mm)']:.0f} mm (zona {pior['Zona']}, "
        f"origem {pior['Origem zona']}) — "
        f"cerca de {pior['Excesso (%)']:.0f}% acima da estimativa de projeto. "
        f"Nota para a interpretacao: uma excedencia desta ordem face ao "
        f"CALCULO de projeto e o ponto de partida da discussao — se reflete a "
        f"geologia real, o faseamento, a execucao ou o dimensionamento e uma "
        f"interpretacao a fundamentar no texto (a memoria descritiva ressalva "
        f"que os proprios criterios ficam 'a confirmar em fase de obra').")

    # ============ (6) MOVIMENTO x FASEAMENTO ============
    st.divider()
    st.markdown("#### 6. Movimento × Faseamento — coincidencia temporal")
    st.caption(
        "Alinha a evolucao do movimento de um alvo com as datas REAIS em que a "
        "escavacao atingiu cada cota (marcos E1-E8). Para cada janela entre "
        "marcos, calcula quanto o alvo se moveu e a que velocidade. Se a "
        "velocidade aumenta a medida que a escavacao desce, ha uma coincidencia "
        "temporal entre o aprofundamento da escavacao e a aceleracao do "
        "movimento. IMPORTANTE: coincidencia temporal e EVIDENCIA para o "
        "argumento causal, nao a sua prova — a interpretacao do mecanismo "
        "(desconfinamento, geologia, faseamento) e do autor.")

    alvos_disp2 = sorted(ult.nlargest(15, COLS["desl_h"])[COLS["alvo"]]
                         .astype(str).tolist())
    alvo6 = st.selectbox("Alvo a analisar (movimento vs. escavacao)",
                         alvos_disp2, key="an_alvo_fase")
    s6 = alvos[alvos[COLS["alvo"]].astype(str) == alvo6].sort_values(
        COLS["data"])[[COLS["data"], COLS["desl_h"]]].dropna()
    if len(s6) >= 2:
        # interpolacao do desl acumulado numa data (escala de datas consistente)
        xd = s6[COLS["data"]].values.astype("datetime64[ns]").astype("int64")
        yv = s6[COLS["desl_h"]].to_numpy()

        def _desl_em(d):
            dv = pd.to_datetime(d).to_datetime64().astype(
                "datetime64[ns]").astype("int64")
            return float(np.interp(dv, xd, yv))

        dmin = s6[COLS["data"]].min()
        dmax = s6[COLS["data"]].max()
        # marcos de escavacao dentro da janela de leituras do alvo
        marcos = []
        for rot, cota, ini, fim in ESCAVACAO_COTAS:
            t = pd.to_datetime(fim)
            if dmin <= t <= dmax:
                marcos.append((t, rot, cota))
        marcos.sort()
        # pontos: inicio -> cada marco -> ultima leitura
        pts = [(dmin, "inicio", None)] + marcos + [(dmax, "ultima leitura", None)]
        linhas6 = []
        for i in range(1, len(pts)):
            d0, r0, c0 = pts[i - 1]
            d1, r1, c1 = pts[i]
            if d1 <= d0:
                continue
            v0 = _desl_em(d0); v1 = _desl_em(d1)
            dias = (d1 - d0).days
            dv = v1 - v0
            vel = dv / dias if dias else 0.0
            linhas6.append({
                "Ate": d1.strftime("%d/%m/%Y"),
                "Marco (cota escavada)": (f"{r1} (cota {c1:.1f})"
                                          if c1 is not None else r1),
                "Movimento na janela (mm)": round(dv, 1),
                "Dias": dias,
                "Velocidade (mm/dia)": round(vel, 3),
            })
        df6 = pd.DataFrame(linhas6)
        st.dataframe(df6, use_container_width=True, hide_index=True)

        # grafico: velocidade por janela (barras) alinhado ao avanco da cota
        fig6 = go.Figure()
        fig6.add_trace(go.Bar(
            x=df6["Ate"], y=df6["Velocidade (mm/dia)"],
            marker=dict(color="#1f78d1"), name="Velocidade (mm/dia)"))
        fig6.update_layout(
            height=340, xaxis=dict(title="Fim da janela de escavacao"),
            yaxis=dict(title="Velocidade (mm/dia)"),
            margin=dict(t=30, b=60),
            legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(fig6, use_container_width=True)

        # leitura factual da tendencia (sem afirmar causa)
        if len(df6) >= 3:
            v_ini = df6["Velocidade (mm/dia)"].iloc[0]
            v_fim = df6["Velocidade (mm/dia)"].iloc[-1]
            cresce = v_fim > v_ini
            racio = (v_fim / v_ini) if v_ini > 0 else float("nan")
            tend = ("AUMENTA" if cresce else "diminui")
            txt = (f"A velocidade do alvo {alvo6} {tend} da primeira janela "
                   f"({v_ini:.2f} mm/dia) para a ultima ({v_fim:.2f} mm/dia)")
            if cresce and v_ini > 0 and racio == racio:
                txt += f" — cerca de {racio:.0f}x mais rapido"
            txt += (". Nota para a interpretacao: um aumento de velocidade "
                    "conforme a escavacao aprofunda e coerente com um efeito de "
                    "desconfinamento progressivo, MAS a confirmacao do mecanismo "
                    "exige cruzar com a geologia local e a sequencia construtiva "
                    "— e argumento do autor, nao conclusao automatica da app.")
            st.caption(txt)
    else:
        st.info("Alvo sem leituras suficientes para a analise temporal.")

    # ============ (7) CELULAS DE CARGA x MOVIMENTO ============
    st.divider()
    st.markdown("#### 7. Cargas nas ancoragens × Movimento")
    st.caption(
        "Cruza a CARGA medida numa celula de carga (ancoragem) com o MOVIMENTO "
        "de um alvo, no mesmo eixo de tempo. Responde a: a ancoragem ganhou "
        "carga (o terreno continuou a empurrar) enquanto o alvo se movia? "
        "Carga acima da blocagem indica que a ancoragem esta a ser solicitada "
        "acima do valor de instalacao. Criterio das celulas: +15% (alerta), "
        "+25% (alarme) sobre a carga. IMPORTANTE: a leitura conjunta e "
        "evidencia sobre o comportamento da contencao; a conclusao sobre "
        "dimensionamento e do autor.")

    cc = dados.get("celulas")
    if cc is None or cc.empty:
        st.info("Sem dados de celulas de carga.")
    else:
        cc = cc.copy()
        cc[COLS["data"]] = pd.to_datetime(cc[COLS["data"]], errors="coerce")
        # escolher a celula primeiro; o filtro de alvos depende da sua zona
        cel_sel = st.selectbox(
            "Celula de carga", sorted(cc[COLS["celula"]].dropna().unique()),
            key="an_celula")
        loc_cel = LOCALIZACAO_CELULAS.get(cel_sel)
        rotulo_zona = loc_cel[0] if loc_cel else None
        filtro_grupo = loc_cel[1] if loc_cel else None

        sc = cc[cc[COLS["celula"]] == cel_sel].sort_values(COLS["data"])
        bloc = float(sc[COLS["blocagem"]].iloc[0])
        anc = sc[COLS["ancoragem"]].iloc[0]

        # RESTRICAO FISICA: se a zona da celula e conhecida, so oferecer alvos
        # do grupo correspondente (mesmo alcado). Caso contrario, todos, com
        # aviso de que o par pode nao ser valido.
        if filtro_grupo:
            alvos_validos = sorted(
                ult[ult[COLS["edificio"]].astype(str)
                    .str.contains(filtro_grupo, case=False, na=False)]
                [COLS["alvo"]].astype(str).unique())
            if not alvos_validos:
                alvos_validos = alvos_disp2   # salvaguarda
            st.success(
                f"✓ Localizacao confirmada: a celula {cel_sel} (ancoragem "
                f"{anc}) esta na **{rotulo_zona}**. O seletor de alvos abaixo "
                f"esta restrito aos alvos dessa zona — o cruzamento carga vs. "
                f"movimento e, assim, fisicamente valido.")
        else:
            alvos_validos = alvos_disp2
            st.warning(
                f"⚠ A localizacao da celula {cel_sel} (ancoragem {anc}) nao "
                f"esta confirmada. O cruzamento so tera significado fisico se o "
                f"alvo pertencer ao mesmo alcado da ancoragem — confirme na "
                f"planta antes de tirar conclusoes.")

        alvo7 = st.selectbox("Alvo a sobrepor (mesma zona da celula)",
                             alvos_validos, key="an_alvo_celula")

        # imagens de localizacao da celula (planta/alcado/3D do projeto)
        _mostrar_imagens_celula(cel_sel, key_suffix="_analise")
        sa = alvos[alvos[COLS["alvo"]].astype(str) == alvo7].sort_values(
            COLS["data"])[[COLS["data"], COLS["desl_h"]]].dropna()

        fig7 = go.Figure()
        # carga (eixo esquerdo)
        fig7.add_trace(go.Scatter(
            x=[pd.to_datetime(v) for v in sc[COLS["data"]].tolist()],
            y=[float(v) for v in sc[COLS["carga_atual"]].tolist()],
            mode="lines+markers", name=f"Carga {cel_sel} (kN)",
            line=dict(color="#7a1fa0", width=2)))
        # movimento do alvo (eixo direito)
        fig7.add_trace(go.Scatter(
            x=[pd.to_datetime(v) for v in sa[COLS["data"]].tolist()],
            y=[float(v) for v in sa[COLS["desl_h"]].tolist()],
            mode="lines+markers", name=f"Desl. H {alvo7} (mm)", yaxis="y2",
            line=dict(color="#c0140f", width=2, dash="dot")))
        # linhas de blocagem e limiares 15%/25% (em kN), yref ao eixo esquerdo
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

        # leitura factual: a carga cresceu? em que estado terminou?
        c0 = float(sc[COLS["carga_atual"]].iloc[0])
        c1 = float(sc[COLS["carga_atual"]].iloc[-1])
        pct_fim = (c1 / bloc - 1) * 100
        estado_fim = sc["Estado"].iloc[-1] if "Estado" in sc.columns else "?"
        cresceu = c1 > c0
        st.caption(
            f"A ancoragem {anc} (celula {cel_sel}) "
            f"{'ganhou' if cresceu else 'perdeu'} carga: {c0:.0f} → {c1:.0f} kN "
            f"({(c1/c0-1)*100:+.0f}%), terminando a {pct_fim:+.0f}% da blocagem "
            f"({estado_fim}). Nota para a interpretacao: se a carga na ancoragem "
            f"sobe ao mesmo tempo que o alvo se move, o terreno esta a solicitar "
            f"a contencao acima do previsto nessa zona — evidencia relevante "
            f"para discutir o dimensionamento e uma eventual solucao alternativa "
            f"(ex.: mais ancoragens ou ancoragens mais precoces na zona critica). "
            f"A conclusao e do autor.")

    st.divider()
    st.info(
        "Esta analise quantifica as movimentacoes e cruza-as com o faseamento e "
        "as cargas nas ancoragens. A interpretacao das causas (mecanismo "
        "geologico, sequencia construtiva, dimensionamento) e a eventual "
        "proposta de solucao alternativa sao desenvolvidas no texto da tese, "
        "usando esta evidencia. A app organiza e quantifica; nao substitui a "
        "interpretacao do autor.")


def separador_sintese(dados):
    """
    Sintese do back-analysis: num unico eixo temporal, cruza a DEFORMACAO
    (alvo OU inclinometro, a escolha) com a COTA DA AGUA (piezometro) e o
    FASEAMENTO da obra. E aqui que a relacao causa-efeito da tese se ve: a
    agua desce e a deformacao acelera nas datas em que a escavacao avanca.
    """
    st.subheader("Sintese — deformacao vs. agua vs. obra")
    st.caption("Os tres fatores do back-analysis num so eixo de tempo: a "
               "deformacao medida (esquerda), a cota da agua subterranea "
               "(direita) e as fases da obra (fundo). Permite ler a "
               "relacao causa-efeito: o rebaixamento da agua e a aceleracao "
               "da deformacao acompanham o avanco da escavacao.")

    c1, c2 = st.columns(2)
    with c1:
        tipo = st.radio("Serie de deformacao", ["Inclinometro", "Alvo"],
                        horizontal=True, key="sint_tipo")
        st.checkbox("Marcos de escavacao por cota (datas reais)", value=True,
                    key="sint_escav",
                    help="Linhas verticais nas datas reais em que a escavacao "
                         "atingiu cada cota (do plano de trabalhos impactado). "
                         "Permite ver a deformacao acelerar apos cada fase.")
    # --- serie de deformacao escolhida ---
    df_def = None
    lbl_def = ""
    with c2:
        if tipo == "Inclinometro":
            res = dados.get("resumo")
            if res is not None and COLS["data"] in res.columns:
                res = res.copy()
                res[COLS["data"]] = pd.to_datetime(res[COLS["data"]], errors="coerce")
                col_inc = "Inclinómetro"
                incs = sorted(res[col_inc].dropna().unique())
                sel = st.selectbox("Inclinometro", incs, key="sint_inc")
                s = res[res[col_inc] == sel].sort_values(COLS["data"])
                df_def = s[[COLS["data"], "Máx. desloc. acumulado total (mm)"]].rename(
                    columns={"Máx. desloc. acumulado total (mm)": "def"})
                lbl_def = f"Desl. max. {sel} (mm)"
        else:
            alv = dados["alvos"].copy()
            alv[COLS["data"]] = pd.to_datetime(alv[COLS["data"]], errors="coerce")
            alvos = sorted(alv[COLS["alvo"]].dropna().unique())
            # sugerir A3 (o alvo critico) se existir
            idx = alvos.index("A3") if "A3" in alvos else 0
            sel = st.selectbox("Alvo", alvos, index=idx, key="sint_alvo")
            s = alv[alv[COLS["alvo"]] == sel].sort_values(COLS["data"])
            df_def = s[[COLS["data"], COLS["desl_h"]]].rename(
                columns={COLS["desl_h"]: "def"})
            lbl_def = f"Desl. horizontal {sel} (mm)"

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
        st.info("Sem dados de deformacao para a serie escolhida.")
        return

    # --- grafico de eixo duplo ---
    fig = go.Figure()
    # fases da obra (faixas + etiquetas diagonais), na janela dos dados
    dt_min = df_def[COLS["data"]].min()
    dt_max = df_def[COLS["data"]].max()
    if df_agua is not None and not df_agua.empty:
        dt_min = min(dt_min, df_agua[COLS["data"]].min())
        dt_max = max(dt_max, df_agua[COLS["data"]].max())
    fases_vis = adicionar_fases_obra(fig, dt_min, dt_max, barra_topo=False)

    # marcos de escavacao por cota (datas reais) — linha vertical + etiqueta
    # CURTA (E1, E2...) no fundo. O rotulo completo (cota + data) vai numa
    # legenda por baixo do grafico. Etiquetas curtas nao colidem, ao
    # contrario dos rotulos longos com datas proximas.
    mostrar_escav = st.session_state.get("sint_escav", True)
    escav_visiveis = []
    if mostrar_escav:
        # so os que caem na janela, ordenados por data
        na_janela = [(pd.to_datetime(fim), rotulo, cota)
                     for rotulo, cota, ini, fim in ESCAVACAO_COTAS
                     if dt_min <= pd.to_datetime(fim) <= dt_max]
        na_janela.sort()
        for k, (t, rotulo, cota) in enumerate(na_janela, start=1):
            fig.add_shape(
                type="line", xref="x", yref="paper",
                x0=t, x1=t, y0=0, y1=1,
                line=dict(color="#8B4513", width=1, dash="dash"), layer="below")
            fig.add_annotation(
                x=t, y=-0.02, yref="paper", text=f"E{k}",
                showarrow=False, xanchor="center", yanchor="top",
                font=dict(size=9, color="white"),
                bgcolor="#8B4513", borderpad=2,
                hovertext=f"{rotulo} — {t.strftime('%d/%m/%Y')}")
            escav_visiveis.append((k, rotulo, t, cota))

    # deformacao (eixo Y esquerdo)
    # NOTA: converter para listas Python puras (list()) em vez de deixar passar
    # arrays numpy/pandas. Com numpy, o Plotly recente serializa os valores em
    # base64 ("bdata"), o que — em combinacao com o eixo Y duplo (overlaying)
    # desta figura — faz as series NAO renderizarem no Plotly.js do Streamlit
    # Cloud (grafico aparecia em branco). As listas puras evitam essa
    # codificacao. Isto e local a Sintese; os outros graficos nao sao afetados.
    fig.add_trace(go.Scatter(
        x=[pd.to_datetime(v) for v in df_def[COLS["data"]].tolist()],
        y=[float(v) for v in df_def["def"].tolist()],
        mode="lines+markers",
        name=lbl_def, line=dict(color="#c0140f", width=2)))
    # agua (eixo Y direito)
    if df_agua is not None and not df_agua.empty:
        fig.add_trace(go.Scatter(
            x=[pd.to_datetime(v) for v in df_agua[COLS["data"]].tolist()],
            y=[float(v) for v in df_agua["agua"].tolist()],
            mode="lines+markers",
            name=f"Cota da agua {pz_sel} (m)", yaxis="y2",
            line=dict(color="#2563eb", width=2, dash="dot")))

    fig.update_layout(
        height=560,
        margin=dict(t=60, b=130),
        # type="date" EXPLICITO: o primeiro trace da figura e um trace fantasma
        # das fases (x=[None], de adicionar_fases_obra em modo barra_topo=False).
        # Sem o tipo forcado, o Plotly.js infere o eixo X a partir desse primeiro
        # trace nulo, falha, e cai para eixo numerico (1,2,3...) — o que fazia as
        # series de data NAO renderizarem (grafico em branco). Forcar o tipo
        # resolve na raiz.
        xaxis=dict(title="Data", type="date"),
        yaxis=dict(title=dict(text=lbl_def, font=dict(color="#c0140f")),
                   tickfont=dict(color="#c0140f")),
        yaxis2=dict(title=dict(text="Cota da agua (m)", font=dict(color="#2563eb")),
                    tickfont=dict(color="#2563eb"),
                    overlaying="y", side="right", anchor="x"),
        # legenda mais abaixo (y=-0.38) para nao colidir com o titulo "Data"
        # do eixo X; a margem inferior (b=130) abre o espaco necessario.
        legend=dict(orientation="h", yanchor="top", y=-0.38,
                    xanchor="center", x=0.5))
    st.plotly_chart(fig, use_container_width=True)
    legenda_fases(fases_vis)
    # legenda dos marcos de escavacao (E1, E2...) -> cota + data
    if escav_visiveis:
        itens = "  ·  ".join(
            f"**E{k}** {rotulo} ({t.strftime('%d/%m')})"
            for k, rotulo, t, cota in escav_visiveis)
        st.caption("⛏ Escavacao (cota atingida, data real): " + itens)

    # leitura automatica da correlacao
    if df_agua is not None and len(df_agua) >= 2 and len(df_def) >= 2:
        d_ini = df_def["def"].iloc[0]
        d_fim = df_def["def"].iloc[-1]
        a_ini = df_agua["agua"].iloc[0]
        a_fim = df_agua["agua"].iloc[-1]
        st.markdown(
            f"**Leitura:** no periodo, a deformacao evoluiu de {d_ini:.1f} para "
            f"**{d_fim:.1f} mm** enquanto a cota da agua desceu de {a_ini:.2f} "
            f"para **{a_fim:.2f} m** ({a_fim - a_ini:+.2f} m). A deformacao "
            f"cresce a medida que a agua e rebaixada e a escavacao avanca — "
            f"consistente com o desconfinamento induzido pela escavacao.")
    st.caption("Deformacao e cota da agua tem escalas independentes (eixos Y "
               "esquerdo/direito). As fases da obra sao o cronograma previsto.")


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

    # ---- enquadramento do caso (o problema, para o juri entrar na narrativa)
    st.markdown(
        "Esta aplicação apoia a **análise inversa** (*back-analysis*) da "
        "contenção periférica da reformulação do **Hotel Eden**, no Monte "
        "Estoril. A obra compreende uma escavação profunda — cerca de 16 m — "
        "executada ao abrigo de uma cortina de contenção, num contexto "
        "exigente: confina com **edifícios sensíveis**, em particular a Santa "
        "Casa da Misericórdia, o Restaurante Cimas e a Clínica Abreu Loureiro, "
        "e desenvolve-se **abaixo do nível freático** de repouso. A "
        "instrumentação instalada — inclinómetros, alvos topográficos, células "
        "de carga e piezómetros — permite acompanhar, ao longo do tempo, os "
        "deslocamentos induzidos pela escavação, tanto na própria contenção "
        "como nos edifícios vizinhos.")
    st.markdown(
        "O propósito da ferramenta não se esgota na visualização das leituras: "
        "procura sobretudo **relacionar a deformação medida com as suas "
        "causas** — o avanço da escavação, o rebaixamento do nível freático e "
        "a natureza do maciço. A leitura conjunta destes fatores sustenta que "
        "a deformação da contenção é governada pelo **grau de consolidação "
        "crescente do grés** — e não por qualquer camada mole, que não existe "
        "— acentuando-se à medida que a escavação progride abaixo do nível "
        "freático. O separador **Síntese** reúne estes três fatores num único "
        "eixo temporal, tornando essa relação diretamente legível.")

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
        ("Sintese", "A relacao causa-efeito num so eixo de tempo: deformacao "
         "(alvo ou inclinometro) vs. cota da agua vs. fases da obra."),
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
        ("Pressupostos", "O que e medido vs. o que e assumido: inferencias, "
         "limitacoes e qualidade dos dados, num so sitio."),
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
        # caminho do Excel ancorado ao diretorio do script (robusto ao CWD)
        base_dir = Path(__file__).resolve().parent
        fonte = base_dir / FICHEIRO_EXCEL
        if not fonte.exists():
            # tentar tambem o CWD, por compatibilidade
            if Path(FICHEIRO_EXCEL).exists():
                fonte = FICHEIRO_EXCEL
            else:
                st.error(f"Nao encontrei '{FICHEIRO_EXCEL}'. Poe o Excel na pasta "
                         f"do script ou usa 'Carregar manualmente'.")
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

    # -------------------------------------------------------------------
    # ORGANIZACAO EM DUAS FRENTES (escolhidas na barra lateral)
    #   Inputs:  dados que alimentam a analise — instrumentacao, geologia
    #            e planeamento (obra).
    #   Outputs: analise e avaliacao de resultados, e comportamento no espaco.
    # O separador Inicio esta sempre acessivel (primeiro de cada frente).
    # -------------------------------------------------------------------
    st.sidebar.divider()
    st.sidebar.subheader("Navegacao")
    frente = st.sidebar.radio(
        "Componente",
        ["Inputs — instrumentacao, geologia e planeamento",
         "Outputs — analise de resultados"],
        help="Inputs: os dados que alimentam a analise (cada instrumento, a "
             "geologia e o planeamento da obra). Outputs: a analise e avaliacao "
             "de resultados e o comportamento no espaco.")

    if frente.startswith("Inputs"):
        thome, tinc, talv, tcc, tpz, tgeo, tobra, tplan = st.tabs(
            ["Inicio", "Inclinometros", "Alvos (2D)", "Celulas de carga",
             "Piezometros", "Geologia", "Obra", "Planta (DXF)"])
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
        thome, t3d, tanal, tsint, tpress = st.tabs(
            ["Inicio", "Terreno + Alvos 3D", "Analise", "Sintese",
             "Pressupostos"])
        with thome:
            separador_home(dados)
        with t3d:
            separador_terreno_alvos_3d(dados)
        with tanal:
            separador_analise(dados)
        with tsint:
            separador_sintese(dados)
        with tpress:
            separador_pressupostos(dados)


if __name__ == "__main__":
    main()
