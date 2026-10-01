"""
Desempenho Analítico - Grupo Sabin
Análise do desempenho analítico dos testes, em módulos: CIQ (exports do Infinity), EP (CAP e
ControlLab) e, em breve, Comparabilidade (harmonização entre equipamentos).

Como rodar:
    pip install streamlit pandas plotly openpyxl
    streamlit run app.py
"""
import io
import json
import os
import re
import zipfile
from collections import defaultdict, Counter
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import ep_base
import ep_calculo

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(page_title="Desempenho Analítico - Grupo Sabin", layout="wide")
# Tabelas coloridas (Styler) com vários meses de todos os testes passam do limite padrão do pandas
pd.set_option("styler.render.max_elements", 2_000_000)

REF_DIR = Path(__file__).parent / "reference_data"
MESES_NOME = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
              7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
MESES_ABREV = {v.lower(): k for k, v in MESES_NOME.items()}

MODULO_PLATAFORMA = {
    "Atellica": "Atellica", "BN": "BN", "Liaison": "Liaison", "Maglumi": "Maglumi",
    "Immulite": "Immulite", "Diestro": "Diestro", "D10": "D10", "C513": "Cobas",
}

STATUS_COLORS = {"VERDE": "#D9EAD3", "AMARELO": "#FFF2CC", "VERMELHO": "#F4CCCC"}


# ============================================================
# CARGA DOS DADOS DE REFERÊNCIA (embutidos no app)
# ============================================================
def _hash_arquivos_referencia():
    """Assinatura baseada na data de modificação dos arquivos de referência.
    Muda sempre que qualquer um deles é editado e salvo, fazendo o
    st.cache_data invalidar e recarregar a tabela_mestre.xlsx (e afins)
    automaticamente — sem precisar reiniciar o app ou limpar o cache."""
    arquivos = ["tabela_mestre.xlsx", "bd_fallback.xlsx", "equipamentos.xlsx", "testes_excluidos.xlsx",
                "niveis_controle.xlsx"]
    partes = []
    for nome in arquivos:
        caminho = REF_DIR / nome
        partes.append(str(caminho.stat().st_mtime_ns) if caminho.exists() else "ausente")
    return "|".join(partes)


# "NA" é o mnemônico do sódio: o pandas, por padrão, lê "NA" (e "N/A", "NULL"...) como vazio
SO_VAZIO_E_NA = {"keep_default_na": False, "na_values": [""]}


@st.cache_data
def load_reference(assinatura):   # sem "_" no nome: o cache só é refeito se o parâmetro for considerado
    df_mestre = pd.read_excel(REF_DIR / "tabela_mestre.xlsx", **SO_VAZIO_E_NA).astype(object)
    mestre = df_mestre.where(pd.notna(df_mestre), None).to_dict("records")

    df_bd = pd.read_excel(REF_DIR / "bd_fallback.xlsx", **SO_VAZIO_E_NA).astype(object)
    df_bd = df_bd.where(pd.notna(df_bd), None)
    bd_fallback = {
        str(row["Mneumonico Infinity"]).strip().upper(): {
            "Analito": row["Analito"], "ETM (%)": row["ETM (%)"],
            "ESM (%)": row["ESM (%)"], "CV Máx (%)": row["CV Máx (%)"],
        }
        for _, row in df_bd.iterrows() if row["Mneumonico Infinity"]
    }

    df_equip = pd.read_excel(REF_DIR / "equipamentos.xlsx", **SO_VAZIO_E_NA).astype(object)
    df_equip = df_equip.where(pd.notna(df_equip), None)
    equip_depara = {
        str(row["Equipamento"]): {
            "Tipo": row["Tipo"], "Módulo": row["Módulo"], "Célula/Rotor": row["Célula/Rotor"],
        }
        for _, row in df_equip.iterrows() if row["Equipamento"]
    }

    testes_excluidos_path = REF_DIR / "testes_excluidos.xlsx"
    if testes_excluidos_path.exists():
        df_excl = pd.read_excel(testes_excluidos_path, **SO_VAZIO_E_NA)
        testes_excluidos = {str(t).strip().upper() for t in df_excl["Teste"].dropna()}
    else:
        testes_excluidos = set()

    # De-para manual de nível do controle (opcional). Colunas: Controle, Lote, Nível e, opcional,
    # Teste. Lote em branco = qualquer lote; Teste em branco = todos os testes daquele controle/lote
    # (Teste preenchido serve pra quando só um analito tem os níveis trocados, ex.: bilirrubina no
    # Lyphochek 8978x, em que o frasco "nível 1" é o de concentração alta).
    niveis_depara = {}
    niveis_path = REF_DIR / "niveis_controle.xlsx"
    if niveis_path.exists():
        df_niv = pd.read_excel(niveis_path, dtype=str, **SO_VAZIO_E_NA)
        for _, row in df_niv.iterrows():
            controle, nivel = row.get("Controle"), row.get("Nível")
            if pd.isna(controle) or pd.isna(nivel):
                continue
            lote = "" if pd.isna(row.get("Lote")) else str(row.get("Lote")).strip().rstrip(".")
            teste = "" if pd.isna(row.get("Teste")) else str(row.get("Teste")).strip().upper()
            niveis_depara[(norm_nivel(str(controle)), lote, teste)] = int(float(nivel))

    mestre_by_mneu = defaultdict(list)
    for r in mestre:
        if r["Mneumonico Infinity"]:
            mestre_by_mneu[str(r["Mneumonico Infinity"]).strip().upper()].append(r)
    return mestre_by_mneu, bd_fallback, equip_depara, testes_excluidos, niveis_depara


def modulo_to_plataforma(modulo):
    if modulo in MODULO_PLATAFORMA:
        return MODULO_PLATAFORMA[modulo]
    if "Bioq" in modulo or "Imuno" in modulo:
        return "Cobas"
    return None


def get_spec(mneu, modulo, mestre_by_mneu, bd_fallback):
    mneu = mneu.strip().upper()
    candidatos = mestre_by_mneu.get(mneu, [])
    if not candidatos:
        bd = bd_fallback.get(mneu)
        if bd:
            return {
                "fonte": "BD_fallback", "Analito": bd["Analito"],
                "CV1": bd["CV Máx (%)"], "CV2": bd["CV Máx (%)"],
                "CV3": bd["CV Máx (%)"], "CV4": bd["CV Máx (%)"],
                "ETM (%)": bd["ETM (%)"],
                "ETM Absoluto - Cutoff": None, "ETM Absoluto - Valor": None,
                "ESM (%)": bd.get("ESM (%)"),
                "ESM Absoluto - Cutoff": None, "ESM Absoluto - Valor": None,
            }
        return None
    escolhido = candidatos[0]
    if len(candidatos) > 1:
        plat = modulo_to_plataforma(modulo)
        if plat:
            for c in candidatos:
                if c["Plataforma"] == plat:
                    escolhido = c
                    break
    return {
        "fonte": "Tabela Mestre", "Analito": escolhido["Analito"],
        "CV1": escolhido["CV Nível 1 (%)"], "CV2": escolhido["CV Nível 2 (%)"],
        "CV3": escolhido["CV Nível 3 (%)"], "CV4": escolhido["CV Nível 4 (%)"],
        "Níveis Avaliados": escolhido.get("Níveis Avaliados"),
        "Sigma Mínimo": escolhido.get("Sigma Mínimo"),
        "ETM (%)": escolhido["ETM (%)"],
        "ETM Absoluto - Cutoff": escolhido["ETM Absoluto - Cutoff"],
        "ETM Absoluto - Valor": escolhido["ETM Absoluto - Valor"],
        "ESM (%)": escolhido["ESM/ES (%)"],
        "ESM Absoluto - Cutoff": escolhido["ESM Absoluto - Cutoff"],
        "ESM Absoluto - Valor": escolhido["ESM Absoluto - Valor"],
    }


# ============================================================
# PARSING DOS CSVs DO INFINITY
# ============================================================
def parse_filename(fname):
    base = Path(fname).stem
    parts = [p.strip() for p in base.split(" - ")]
    mes_ano = parts[-1]
    m = re.search(r"([A-Za-z]{3})_?(\d{2})", mes_ano)
    if not m:
        return None, None, None
    mes = MESES_ABREV.get(m.group(1).lower())
    ano = 2000 + int(m.group(2))
    modulo = parts[1] if len(parts) >= 3 else (parts[0] if parts else fname)
    return modulo, mes, ano


def norm_nivel(s):
    if s is None:
        return s
    s = re.sub(r"\s+", " ", s.strip())
    s = re.sub(r"(?<=[A-Za-zÀ-ÿ])\s+(?=\d)", " ", s)
    s = re.sub(r"^([A-Za-zÀ-ÿ]+)(\d)", r"\1 \2", s)
    return s.upper()


def to_float(v):
    if v is None or v == "" or v == "-":
        return None
    if isinstance(v, str):
        # No export do Infinity a vírgula é separador de milhar ("1,905.0"), não decimal
        v = v.strip().replace(",", "")
    try:
        return float(v)
    except ValueError:
        return None


def extrai_nivel(nome):
    if not nome:
        return None
    m = re.search(r"(\d+)\s*$", nome.strip())
    if m:
        n = int(m.group(1))
        if 1 <= n <= 4:
            return n
    return None


@st.cache_data(show_spinner=False)
def parse_zip(zip_bytes):
    """Lê um .zip com os CSVs mensais do Infinity e retorna a lista de registros brutos."""
    records = []
    skipped = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        csv_names = [n for n in z.namelist() if n.lower().endswith(".csv")]
        for name in csv_names:
            fname = Path(name).name
            modulo, mes, ano = parse_filename(fname)
            if mes is None:
                skipped.append((fname, "mês/ano não identificado no nome do arquivo"))
                continue
            try:
                raw = z.read(name).decode("utf-8-sig")
            except UnicodeDecodeError:
                raw = z.read(name).decode("latin-1")
            reader = pd.read_csv(io.StringIO(raw), sep=";", dtype=str, **SO_VAZIO_E_NA)
            for _, row in reader.iterrows():
                teste = (row.get("Teste") or "").strip() if pd.notna(row.get("Teste")) else ""
                if not teste:
                    continue
                nivel_raw = row.get("Controlos de CQ")
                nivel_raw = nivel_raw if pd.notna(nivel_raw) else None
                if nivel_raw and str(nivel_raw).upper().startswith("COMP"):
                    continue  # harmonização - fora do módulo de CIQ
                records.append({
                    "Arquivo": fname,
                    "Módulo/Arquivo": modulo,
                    "Ano": ano, "Mês": mes,
                    "Equipamento": str(row.get("Equipamento") or "").strip(),
                    "Teste": teste,
                    "Nível": norm_nivel(nivel_raw),
                    "Número de lote": str(row.get("Número de lote") or "").strip().rstrip("."),
                    "N": to_float(row.get("N")),
                    "Config. valor alvo": to_float(row.get("Config. valor alvo")),
                    "Média": to_float(row.get("Média")),
                    "CV (%)": to_float(row.get("CV (%)")),
                })
    # Remove linhas exatamente duplicadas (ex: o mesmo arquivo entrando duas
    # vezes no ZIP), que inflariam indevidamente médias e contagens.
    vistos = set()
    unicos = []
    for r in records:
        chave = tuple((k, v) for k, v in r.items() if k != "Arquivo")
        if chave not in vistos:
            vistos.add(chave)
            unicos.append(r)
    records = unicos
    return records, skipped, len(csv_names)


# Controles Bio-Rad (e afins): o último dígito do lote é o nível (85771 → 1,
# 85783 → 3, 45992T2 → 2). O nome no Infinity muitas vezes termina no lote
# ("IMMUNOLOGY 85773"), então o nível não sai do nome.
FAMILIAS_NIVEL_NO_LOTE = ("LIQUICHECK", "IMMUNOLOGY", "IMMUNOASSAY", "IMUNOASSAY",
                          "LYPHOCHEK", "LIPHO", "INTELIQ", "BNP", "HCY")


def nivel_pelo_lote(nivel_nome, lote):
    if not nivel_nome or not lote:
        return None
    if not any(f in nivel_nome for f in FAMILIAS_NIVEL_NO_LOTE):
        return None
    m = re.fullmatch(r"\d{4,}(?:T\d)?", lote.strip().upper())
    if not m:
        return None
    n = int(lote.strip()[-1])
    return n if 1 <= n <= 4 else None


def nivel_pelo_sufixo_c(nivel_nome):
    """Diestro/Maglumi: "CQCAIN 1C", "CQICA 50N2C" → nível antes do C final."""
    if not nivel_nome:
        return None
    m = re.search(r"(\d)C$", nivel_nome.strip())
    if m and 1 <= int(m.group(1)) <= 4:
        return int(m.group(1))
    return None


def atribui_niveis(recs, niveis_depara=None):
    """Nível fixo de cada controle, nesta ordem de prioridade:
    1. de-para manual (reference_data/niveis_controle.xlsx), por controle + lote ou só controle;
    2. número 1-4 no fim do nome do controle;
    3. padrão "1C/2C/3C" no fim do nome;
    4. último dígito do lote, nas famílias Bio-Rad;
    5. fallback: ranking por concentração dentro do mesmo mês (sinalizado na tela)."""
    niveis_depara = niveis_depara or {}
    for r in recs:
        lote = r.get("Número de lote") or ""
        teste = r["Teste"].strip().upper()
        manual = next((niveis_depara[k] for k in ((r["Nível"], lote, teste), (r["Nível"], lote, ""),
                                                  (r["Nível"], "", teste), (r["Nível"], "", ""))
                       if k in niveis_depara), None)
        candidatos = [
            (manual, "de-para manual"),
            (extrai_nivel(r["Nível"]), "nome do controle"),
            (nivel_pelo_sufixo_c(r["Nível"]), "nome do controle (sufixo C)"),
            (nivel_pelo_lote(r["Nível"], lote), "lote do controle"),
        ]
        r["NívelNum"], r["NívelOrigem"] = next(((n, o) for n, o in candidatos if n is not None), (None, None))
        # nome e lote apontando níveis diferentes: vale o nome, mas sinaliza pra conferência
        n_lote = nivel_pelo_lote(r["Nível"], lote)
        r["Conflito Nível"] = (manual is None and n_lote is not None
                               and r["NívelNum"] is not None and n_lote != r["NívelNum"])

    grupos_mes = defaultdict(set)
    for r in recs:
        if r["NívelNum"] is None:
            grupos_mes[(r["Teste"], r["Equipamento"], r["Ano"], r["Mês"])].add(r["Nível"])

    media_local = {}
    for r in recs:
        if r["NívelNum"] is None and r["Média"] is not None:
            key = (r["Teste"], r["Equipamento"], r["Ano"], r["Mês"], r["Nível"])
            media_local[key] = r["Média"]

    rank = {}
    for (teste, equip, ano, mes), niveis in grupos_mes.items():
        # desempate pelo nome: sem ele, controles com a mesma Média trocavam de nível a cada execução
        ordenados = sorted(niveis, key=lambda n: (media_local.get((teste, equip, ano, mes, n), 0), n or ""))
        for i, n in enumerate(ordenados[:4], start=1):
            rank[(teste, equip, ano, mes, n)] = i

    for r in recs:
        if r["NívelNum"] is None:
            key = (r["Teste"], r["Equipamento"], r["Ano"], r["Mês"], r["Nível"])
            r["NívelNum"] = rank.get(key)
            r["NívelOrigem"] = "concentração (fallback mensal)"
    return recs


# ============================================================
# CÁLCULOS: Bias, Sigma, Status CV, Tendência
# ============================================================
def clip(v, lo, hi):
    return None if v is None else max(lo, min(hi, v))


def calc_bias(media, alvo):
    if media is None or alvo is None or alvo == 0:
        return None
    return (media - alvo) / alvo * 100


def calc_sigma(etm_pct, bias_pct, cv_pct):
    if etm_pct is None or bias_pct is None or cv_pct is None or cv_pct == 0:
        return None
    den = cv_pct * (1 + abs(bias_pct) / 100)
    if den == 0:
        return None
    return clip((etm_pct - abs(bias_pct)) / den, 0, 10)


def calc_sigma_abs(etm_abs, media, target_abs, cv_pct):
    if etm_abs is None or media is None or target_abs is None or cv_pct is None:
        return None
    sd = (cv_pct / 100) * media
    if sd == 0:
        return None
    return clip((etm_abs - abs(media - target_abs)) / sd, 0, 10)


def processa(recs, mestre_by_mneu, bd_fallback, equip_depara, margem=0.05, n_minimo=5):
    for r in recs:
        r["Equip Mapeado"] = r["Equipamento"] in equip_depara
        r["Equipamento (nome)"] = equip_depara.get(r["Equipamento"], {}).get("Célula/Rotor", r["Equipamento"])
        r["Tipo"] = equip_depara.get(r["Equipamento"], {}).get("Tipo", "?")

        # Volume insuficiente (N <= n_minimo): não há análise confiável possível
        # nessa linha específica. Zera CV(%), Bias(%) e toda a análise, mantendo
        # só os dados brutos (N, Média, lote etc.). Não afeta outras linhas do
        # mesmo teste com N maior (outro mês/equipamento).
        if r["N"] is None or r["N"] <= n_minimo:
            r["Spec"] = None
            r["Spec - Analito"] = None
            r["CV (%)"] = None
            r["CV Máximo"] = None
            r["Status CV"] = None
            r["Bias (%)"] = None
            r["Bias Máximo"] = None
            r["Status Bias"] = None
            r["Sigma Mensal"] = None
            r["Sigma Mínimo"] = None
            r["Status Sigma"] = None
            r["Classificação Sigma (Westgard)"] = None
            r["Erro Total Observado"] = None
            r["ETM (para comparação)"] = None
            r["Status Erro Total"] = None
            continue

        spec = get_spec(r["Teste"], r["Módulo/Arquivo"], mestre_by_mneu, bd_fallback)
        r["Spec"] = spec
        r["Bias (%)"] = calc_bias(r["Média"], r["Config. valor alvo"])

        if spec is None:
            r["CV Máximo"] = None
            r["Status CV"] = None
            r["Sigma Mensal"] = None
            r["Sigma Mínimo"] = None
            r["Status Sigma"] = None
            r["Classificação Sigma (Westgard)"] = None
            r["Spec - Analito"] = None
            r["Bias Máximo"] = None
            r["Status Bias"] = None
            r["Erro Total Observado"] = None
            r["ETM (para comparação)"] = None
            r["Status Erro Total"] = None
            continue

        r["Spec - Analito"] = spec["Analito"]

        # Nível não avaliado (fora da lista "Níveis Avaliados" da tabela_mestre):
        # mesmo tratamento do "sem especificação" — mantém CV(%)/Bias(%) brutos,
        # mas sem status/limites, já que não há critério de avaliação pra esse nível.
        niveis_raw = spec.get("Níveis Avaliados")
        if niveis_raw not in (None, ""):
            niveis_permitidos = {int(x.strip()) for x in str(niveis_raw).split(",") if x.strip().isdigit()}
            if niveis_permitidos and r["NívelNum"] not in niveis_permitidos:
                r["CV Máximo"] = None
                r["Status CV"] = None
                r["Sigma Mensal"] = None
                r["Sigma Mínimo"] = None
                r["Status Sigma"] = None
                r["Classificação Sigma (Westgard)"] = None
                r["Bias Máximo"] = None
                r["Status Bias"] = None
                r["Erro Total Observado"] = None
                r["ETM (para comparação)"] = None
                r["Status Erro Total"] = None
                continue

        r["Spec - Analito"] = spec["Analito"]
        cvmax = {1: spec["CV1"], 2: spec["CV2"], 3: spec["CV3"], 4: spec["CV4"]}.get(r["NívelNum"])
        r["CV Máximo"] = cvmax
        cv = r["CV (%)"]
        if cvmax is None or cv is None:
            r["Status CV"] = None
        elif cv > cvmax:
            r["Status CV"] = "VERMELHO"
        elif margem is not None and cv >= cvmax * (1 - margem):
            r["Status CV"] = "AMARELO"
        else:
            r["Status CV"] = "VERDE"

        cutoff = spec.get("ETM Absoluto - Cutoff")
        etm_abs = spec.get("ETM Absoluto - Valor")
        media = r["Média"]
        alvo = r["Config. valor alvo"]
        bias_pct = r["Bias (%)"]
        usa_abs = cutoff is not None and media is not None and media < cutoff and etm_abs is not None
        if usa_abs:
            r["Sigma Mensal"] = calc_sigma_abs(etm_abs, media, alvo, cv)
            r["Critério Sigma"] = "absoluto"
        else:
            r["Sigma Mensal"] = calc_sigma(spec.get("ETM (%)"), bias_pct, cv)
            r["Critério Sigma"] = "percentual"

        sigma = r["Sigma Mensal"]
        sigma_min = spec.get("Sigma Mínimo")
        r["Sigma Mínimo"] = sigma_min
        # Status Sigma: mesmo padrão Verde/Amarelo/Vermelho de CV/Bias/Erro Total,
        # só que invertido — aqui quanto MAIOR o Sigma, melhor.
        if sigma is None or sigma_min is None:
            r["Status Sigma"] = None
        elif sigma < sigma_min:
            r["Status Sigma"] = "VERMELHO"
        elif margem is not None and sigma <= sigma_min * (1 + margem):
            r["Status Sigma"] = "AMARELO"
        else:
            r["Status Sigma"] = "VERDE"

        # Classificação Sigma (Westgard): escala fixa e universal, independente
        # da meta cadastrada — funciona mesmo sem "Sigma Mínimo" preenchido.
        if sigma is None:
            r["Classificação Sigma (Westgard)"] = None
        elif sigma >= 6:
            r["Classificação Sigma (Westgard)"] = "Classe Mundial"
        elif sigma >= 4:
            r["Classificação Sigma (Westgard)"] = "Bom"
        elif sigma >= 3:
            r["Classificação Sigma (Westgard)"] = "Marginal"
        else:
            r["Classificação Sigma (Westgard)"] = "Inadequado"

        # --- Bias vs Bias Máximo (ESM) ---
        esm_cutoff = spec.get("ESM Absoluto - Cutoff")
        esm_abs = spec.get("ESM Absoluto - Valor")
        usa_abs_esm = esm_cutoff is not None and media is not None and media < esm_cutoff and esm_abs is not None
        if usa_abs_esm and alvo is not None and media is not None:
            bias_max = esm_abs
            bias_obs = abs(media - alvo)
            r["Critério Bias"] = "absoluto"
        else:
            bias_max = spec.get("ESM (%)")
            bias_obs = abs(bias_pct) if bias_pct is not None else None
            r["Critério Bias"] = "percentual"
        r["Bias Máximo"] = bias_max
        r["Bias Observado"] = bias_obs
        r["Bias Observado (sinal)"] = (media - alvo) if usa_abs_esm and media is not None and alvo is not None else bias_pct
        if bias_max is None or bias_obs is None:
            r["Status Bias"] = None
        elif bias_obs > bias_max:
            r["Status Bias"] = "VERMELHO"
        elif margem is not None and bias_obs >= bias_max * (1 - margem):
            r["Status Bias"] = "AMARELO"
        else:
            r["Status Bias"] = "VERDE"

        # --- Erro Total Observado (|Bias| + 1,65 x CV) vs ETM ---
        Z = 1.65
        if usa_abs and media is not None and alvo is not None and cv is not None:
            sd_abs = (cv / 100) * media
            et_obs = abs(media - alvo) + Z * sd_abs
            etm_ref = etm_abs
            r["Critério Erro Total"] = "absoluto"
        elif bias_pct is not None and cv is not None:
            et_obs = abs(bias_pct) + Z * cv
            etm_ref = spec.get("ETM (%)")
            r["Critério Erro Total"] = "percentual"
        else:
            et_obs = None
            etm_ref = None
            r["Critério Erro Total"] = None
        r["Erro Total Observado"] = et_obs
        r["ETM (para comparação)"] = etm_ref
        if etm_ref is None or et_obs is None:
            r["Status Erro Total"] = None
        elif et_obs > etm_ref:
            r["Status Erro Total"] = "VERMELHO"
        elif margem is not None and et_obs >= etm_ref * (1 - margem):
            r["Status Erro Total"] = "AMARELO"
        else:
            r["Status Erro Total"] = "VERDE"
    return recs


def calcula_tendencia(recs):
    series = defaultdict(list)
    for r in recs:
        if r["NívelNum"] is None:
            continue
        series[(r["Teste"], r["Equipamento (nome)"], r["NívelNum"])].append(r)
    for lst in series.values():
        lst.sort(key=lambda r: (r["Ano"], r["Mês"]))
        for r in lst:
            r["Tendência CV"] = "—"
        for i in range(2, len(lst)):
            cv0, cv1, cv2 = lst[i - 2]["CV (%)"], lst[i - 1]["CV (%)"], lst[i]["CV (%)"]
            if None not in (cv0, cv1, cv2) and cv1 > cv0 and cv2 > cv1:
                lst[i]["Tendência CV"] = "⚠ CV subindo 3+ meses seguidos"
    return recs, series


def periodo_trimestre(mes):
    return (mes - 1) // 3 + 1


def periodo_semestre(mes):
    return 1 if mes <= 6 else 2


def calcula_sigma_periodos(series):
    period_records = []
    for (teste, equip, nivel), lst in series.items():
        by_trim, by_sem, by_ano = defaultdict(list), defaultdict(list), defaultdict(list)
        for r in lst:
            if r["Sigma Mensal"] is None:
                continue
            by_trim[(r["Ano"], periodo_trimestre(r["Mês"]))].append(r["Sigma Mensal"])
            by_sem[(r["Ano"], periodo_semestre(r["Mês"]))].append(r["Sigma Mensal"])
            by_ano[r["Ano"]].append(r["Sigma Mensal"])
        for (ano, trim), vals in by_trim.items():
            period_records.append({"Teste": teste, "Equipamento": equip, "Nível": nivel,
                                    "Período": "Trimestral", "Ano": ano, "Sub-período": f"T{trim}",
                                    "Sigma (pior cenário)": min(vals), "N meses": len(vals)})
        for (ano, sem), vals in by_sem.items():
            period_records.append({"Teste": teste, "Equipamento": equip, "Nível": nivel,
                                    "Período": "Semestral", "Ano": ano, "Sub-período": f"S{sem}",
                                    "Sigma (pior cenário)": min(vals), "N meses": len(vals)})
        for ano, vals in by_ano.items():
            period_records.append({"Teste": teste, "Equipamento": equip, "Nível": nivel,
                                    "Período": "Anual", "Ano": ano, "Sub-período": str(ano),
                                    "Sigma (pior cenário)": min(vals), "N meses": len(vals)})
    return pd.DataFrame(period_records)


def calcula_sigma_periodos_df(df_filtrado):
    """Mesma lógica de calcula_sigma_periodos, mas operando direto sobre um DataFrame
    já filtrado (respeita os filtros globais de Período e Módulo/Plataforma). Também traz
    o CV e o Bias do mês que gerou o pior Sigma, pra dar pista da causa."""
    base = df_filtrado.dropna(subset=["Sigma Mensal"]).copy()
    if base.empty:
        return pd.DataFrame(columns=["Teste", "Equipamento", "Nível", "Ano", "Sub-período",
                                      "Sigma (pior cenário)", "Mês do pior cenário",
                                      "CV (%)", "Bias (%)", "N meses"])

    # quando há mais de um lote de controle no mesmo nível/mês, usa o que teve
    # mais pontos (N) em vez de contar as duas linhas como se fossem meses diferentes
    base["N"] = base["N"].fillna(0)
    idx_maior_n = base.groupby(["Teste", "Equipamento (nome)", "NívelNum", "_ordem_tempo"])["N"].idxmax()
    base = base.loc[idx_maior_n]

    base["Trimestre"] = base["Mês"].apply(periodo_trimestre)
    base["Semestre"] = base["Mês"].apply(periodo_semestre)

    period_records = []
    for (periodo_nome, cols, sub_fmt) in [
        ("Trimestral", ["Teste", "Equipamento (nome)", "NívelNum", "Ano", "Trimestre"], lambda a, s: f"T{s}"),
        ("Semestral", ["Teste", "Equipamento (nome)", "NívelNum", "Ano", "Semestre"], lambda a, s: f"S{s}"),
        ("Anual", ["Teste", "Equipamento (nome)", "NívelNum", "Ano"], lambda a, s: str(a)),
    ]:
        idx_pior = base.groupby(cols)["Sigma Mensal"].idxmin()
        contagem = base.groupby(cols).size()
        piores = base.loc[idx_pior.values].copy()
        piores["_chave"] = list(zip(*[piores[c] for c in cols])) if len(cols) > 1 else piores[cols[0]]
        for _, row in piores.iterrows():
            sub_valor = row[cols[-1]] if len(cols) > 4 else None
            period_records.append({
                "Teste": row["Teste"], "Equipamento": row["Equipamento (nome)"],
                "Nível": int(row["NívelNum"]), "Período": periodo_nome, "Ano": int(row["Ano"]),
                "Sub-período": sub_fmt(row["Ano"], sub_valor),
                "Sigma (pior cenário)": row["Sigma Mensal"], "Mês do pior cenário": row["Mês/Ano"],
                "CV (%)": row["CV (%)"], "Bias (%)": row["Bias (%)"],
                "N meses": int(contagem.loc[row["_chave"]]),
            })
    return pd.DataFrame(period_records)


# ============================================================
# UI
# ============================================================
st.title("Desempenho Analítico")
st.caption("Grupo Sabin · Módulos: CIQ (Infinity) · EP (CAP e ControlLab) · Comparabilidade (em breve)")

mestre_by_mneu, bd_fallback, equip_depara, testes_excluidos, niveis_depara = load_reference(
    _hash_arquivos_referencia())


def nome_amigavel_equip(raw):
    info = equip_depara.get(raw)
    return info["Célula/Rotor"] if info else raw

with st.sidebar:
    # Modo de análise: "Simples" mostra CV, Bias e Erro Total com semáforo, sem Sigma
    # (nem a aba Sigma por Período, nem as sugestões técnicas de CV/Bias).
    # Por enquanto qualquer usuário pode trocar; no futuro o padrão virá do perfil da regional.
    modo_analise = st.segmented_control(
        "Modo de análise", ["Simples", "Completa"], default="Completa",
        key="modo_analise", required=True,
        help="Simples: CV, Bias e Erro Total, sem Sigma. Completa: tudo, incluindo Sigma por período.",
    )
    completo = modo_analise == "Completa"

    st.header("Dados de entrada")
    uploaded_zip = st.file_uploader("ZIP com os CSVs mensais do Infinity", type="zip")
    st.caption(
        "A tabela de especificações (ETM, ESM, CV Máximo por nível) já vem "
        "embutida no app — só é preciso subir os dados novos do Infinity."
    )
    pacote_ep = st.file_uploader(
        "Base de EP (.zip) — opcional", type="zip", key="pacote_ep",
        help="O arquivo pacote_ep.zip (pasta dados_ep do app, ou \"Baixar pacote de EP\" na aba EP). "
             "Um arquivo só, com CAP e ControlLab juntos — não os zips originais dos provedores. No app "
             "online, é assim que a base de EP chega: ele não lê os PDFs do CAP nem guarda arquivos.")
    st.caption("Base de EP: envie só o **pacote_ep.zip** — ele já traz CAP e ControlLab juntos.")
    st.header("Margem de proximidade")
    margem_opcao = st.radio(
        "Sinalizar em amarelo quando estiver a quantos % do limite?",
        ["Desativado", "5%", "10%"], index=0, horizontal=True,
    )
    # margem None = sem faixa amarela: só verde (dentro) ou vermelho (fora do limite)
    margem_pct = None if margem_opcao == "Desativado" else int(margem_opcao.rstrip("%"))
    margem = None if margem_pct is None else margem_pct / 100
    if margem is None:
        st.caption("Vermelho = ultrapassou o limite. Verde = dentro do limite. Sem faixa amarela.")
    else:
        st.caption(
            f"Vermelho = ultrapassou o limite. Amarelo = está dentro de {margem_pct}% "
            "do limite (CV Máximo, Bias Máximo ou ETM, conforme a aba)."
        )

if uploaded_zip is None:
    st.info("Envie o arquivo .zip com os exports mensais do Infinity para começar.")
    st.stop()

with st.spinner("Lendo arquivos..."):
    recs, skipped, n_files = parse_zip(uploaded_zip.getvalue())
    if testes_excluidos:
        recs = [r for r in recs if r["Teste"].strip().upper() not in testes_excluidos]
    recs = atribui_niveis(recs, niveis_depara)
    recs = processa(recs, mestre_by_mneu, bd_fallback, equip_depara, margem=margem)
    recs, series = calcula_tendencia(recs)

st.success(
    f"{n_files} arquivos lidos · {len(recs)} registros de CIQ processados "
    f"(COMP e {len(testes_excluidos)} teste(s) reflexo excluídos)"
)
if skipped:
    with st.expander(f"⚠ {len(skipped)} arquivo(s) não processado(s)"):
        for fname, motivo in skipped:
            st.write(f"- **{fname}**: {motivo}")

if testes_excluidos:
    with st.expander(f"ℹ️ {len(testes_excluidos)} teste(s) reflexo excluído(s) da análise"):
        st.caption(
            "Lista vem de `reference_data/testes_excluidos.xlsx` — adicione uma linha nesse "
            "arquivo (coluna 'Teste', com o código do Infinity) pra excluir outros testes "
            "reflexo no futuro, sem precisar mexer no código."
        )
        for t in sorted(testes_excluidos):
            st.write(f"- `{t}`")

def _tabela_controles(filtro):
    linhas = {}
    for r in recs:
        if filtro(r):
            chave = (r["Equipamento (nome)"], r["Nível"], r.get("Número de lote") or "")
            linhas.setdefault(chave, {"Equipamento": chave[0], "Controle": chave[1], "Lote": chave[2],
                                      "Nível atribuído": r["NívelNum"], "Testes": set()})
            linhas[chave]["Testes"].add(r["Teste"])
    tab = pd.DataFrame(list(linhas.values()))
    if not tab.empty:
        tab["Testes"] = tab["Testes"].apply(lambda s: ", ".join(sorted(s)))
        tab = tab.sort_values(["Equipamento", "Controle", "Lote"])
    return tab


sem_nivel_fixo = _tabela_controles(lambda r: r["NívelOrigem"] == "concentração (fallback mensal)")
if not sem_nivel_fixo.empty:
    with st.expander(
        f"⚠ {len(sem_nivel_fixo)} controle(s) sem nível fixo — nível estimado pela concentração "
        "dentro de cada mês (pode mudar de um mês para outro)"
    ):
        st.caption(
            "O nível desses controles não aparece no nome nem no lote. Pra fixar, cadastre em "
            "`reference_data/niveis_controle.xlsx` (colunas Controle, Lote, Nível e, opcional, Teste — "
            "Lote em branco vale pra qualquer lote; Teste em branco vale pra todos os testes). O de-para "
            "manual tem prioridade sobre qualquer outra regra."
        )
        st.dataframe(sem_nivel_fixo, hide_index=True, use_container_width=True)
        modelo = io.BytesIO()
        (sem_nivel_fixo[["Controle", "Lote"]].drop_duplicates()
         .assign(**{"Nível": None}).to_excel(modelo, index=False))
        st.download_button("Baixar modelo do niveis_controle.xlsx", modelo.getvalue(),
                           "niveis_controle.xlsx")

conflitos_nivel = _tabela_controles(lambda r: r.get("Conflito Nível"))
if not conflitos_nivel.empty:
    with st.expander(f"⚠ {len(conflitos_nivel)} controle(s) com nível do nome diferente do nível do lote"):
        st.caption(
            "O nome do controle indica um nível e o último dígito do lote indica outro. O app usa o "
            "nível do nome; se o certo for o do lote, cadastre em `reference_data/niveis_controle.xlsx`."
        )
        st.dataframe(conflitos_nivel, hide_index=True, use_container_width=True)

nao_mapeados = sorted(set(r["Equipamento"] for r in recs if not r["Equip Mapeado"]))
if nao_mapeados:
    with st.expander(
        f"⚠ {len(nao_mapeados)} nome(s) de equipamento não cadastrado(s) em equipamentos.xlsx "
        "(aparecem com o nome bruto feio em vez do nome amigável, e podem 'sumir' de tabelas que "
        "esperam o nome já unificado)"
    ):
        st.caption(
            "Isso costuma acontecer quando o Infinity troca a identificação de um instrumento "
            "(lote, número de série) e o `equipamentos.xlsx` ainda não foi atualizado com o nome "
            "novo. Adicione uma linha nesse arquivo pra cada um dos nomes abaixo, com o Tipo/Módulo/"
            "Célula-Rotor corretos, e suba o app de novo."
        )
        for nome in nao_mapeados:
            st.write(f"- `{nome}`")

df = pd.DataFrame(recs)
df["Mês/Ano"] = df.apply(lambda r: f'{MESES_NOME.get(r["Mês"], "?")}/{str(r["Ano"])[2:]}', axis=1)
df["_ordem_tempo"] = df["Ano"] * 12 + df["Mês"]
# CIQ completo, antes dos filtros: o EP busca o CV/Média do mês de cada rodada
df_ciq_total = df.copy()

with st.sidebar:
    st.header("Filtros globais (valem para todos os módulos)")
    meses_ordenados_global = (df[["Ano", "Mês", "Mês/Ano"]].drop_duplicates()
                               .sort_values(["Ano", "Mês"])["Mês/Ano"].tolist())
    f_periodo_global = st.select_slider(
        "Período", options=meses_ordenados_global,
        value=(meses_ordenados_global[0], meses_ordenados_global[-1]), key="periodo_global",
    )
    tipos_global = sorted(df["Tipo"].dropna().unique())
    f_tipo_global = st.multiselect(
        "Módulo/Plataforma", tipos_global, default=[], key="tipo_global",
        help="Deixe vazio para incluir todos.",
    )

ini_ord_g = meses_ordenados_global.index(f_periodo_global[0])
fim_ord_g = meses_ordenados_global.index(f_periodo_global[1])
periodo_valido_g = set(meses_ordenados_global[ini_ord_g:fim_ord_g + 1])
# eixo de meses sempre em ordem cronológica (sem isso, um mês que só aparece na 2ª curva vai pro fim do eixo)
EIXO_MESES = {"xaxis_categoryorder": "array", "xaxis_categoryarray": meses_ordenados_global}
df = df[df["Mês/Ano"].isin(periodo_valido_g)]
if f_tipo_global:
    df = df[df["Tipo"].isin(f_tipo_global)]

with st.sidebar:
    st.caption(f"{len(df)} registros após os filtros globais (de {len(recs)} no total).")

PALETA = ["#4B2E5A", "#C2185B", "#00838F", "#F57F17", "#2E7D32", "#5D4037",
          "#1565C0", "#AD1457", "#00695C", "#EF6C00", "#6A1B9A", "#33691E"]
testes_disponiveis = sorted(df["Teste"].dropna().unique())


def cor_sigma(val):
    if pd.isna(val):
        return ""
    if val < 3:
        return "background-color: #F4CCCC"
    elif val < 6:
        return "background-color: #FFF2CC"
    return "background-color: #D9EAD3"


def metricas_status(contagem, com_avaliados=False):
    """Cartões Verde/Amarelo/Vermelho. Sem margem, o cartão amarelo não aparece."""
    itens = [("Avaliados", sum(contagem.values()))] if com_avaliados else []
    itens.append(("🟩 Dentro do limite", contagem.get("VERDE", 0)))
    if margem is not None:
        itens.append((f"🟨 Próximo ({margem_pct}%)", contagem.get("AMARELO", 0)))
    itens.append(("🟥 Acima do máximo", contagem.get("VERMELHO", 0)))
    for col, (rotulo, valor) in zip(st.columns(len(itens)), itens):
        col.metric(rotulo, valor)


def cor_faixa(val, maximo):
    """Cor da célula nas tabelas mensais: faixas de 5% e 10% só com a margem ligada."""
    if val > maximo:
        return "background-color: #F4CCCC"
    if margem is not None:
        if val >= maximo * 0.95:
            return "background-color: #FFCC80"
        if val >= maximo * 0.90:
            return "background-color: #FFF9C4"
    return "background-color: #D9EAD3"


LEGENDA_FAIXAS = ("🟩 dentro do limite · 🟧 dentro de 5% · 🟨 dentro de 10% · 🟥 acima do limite."
                  if margem is not None else "🟩 dentro do limite · 🟥 acima do limite.")


def card_pior_cenario(df_teste, coluna_valor, nome_metrica, usa_abs=False, maior_eh_pior=True):
    """Mostra um card único com o pior valor da métrica (e o mês), considerando todos
    os equipamentos e níveis do teste já filtrado em df_teste."""
    base = df_teste.dropna(subset=[coluna_valor]).copy()
    if base.empty:
        st.info(f"Sem dados de {nome_metrica} pra esse teste no filtro atual.")
        return
    base["N"] = base["N"].fillna(0)
    idx_maior_n = base.groupby(["Equipamento (nome)", "NívelNum", "_ordem_tempo"])["N"].idxmax()
    base = base.loc[idx_maior_n]
    valores = base[coluna_valor].abs() if usa_abs else base[coluna_valor]
    idx_pior = valores.idxmax() if maior_eh_pior else valores.idxmin()
    pior = base.loc[idx_pior]
    valor_pior = abs(pior[coluna_valor]) if usa_abs else pior[coluna_valor]
    rotulo = "Pior" if maior_eh_pior else "Pior (mais baixo)"
    with st.container(border=True):
        c1, c2 = st.columns([3, 1])
        with c1:
            st.caption(f"{rotulo} {nome_metrica} do período")
            st.markdown(
                f"<span style='font-size:30px; font-weight:700;'>{valor_pior:.2f}</span>",
                unsafe_allow_html=True,
            )
            st.caption(f"{pior['Equipamento (nome)']} · Nível {int(pior['NívelNum'])}")
        with c2:
            st.caption("Mês/Ano")
            st.markdown(f"**{pior['Mês/Ano']}**")


with st.sidebar:
    st.header("Teste em foco")
    teste_global = st.selectbox(
        "Usado nas análises por teste (CV, Bias, Erro Total, Sigma e EP)",
        testes_disponiveis, key="teste_global",
    )

# Três módulos, cada um com as suas análises em abas. Os blocos abaixo escrevem direto na aba de
# destino ("with tab_x:"), então a ordem do código não precisa seguir a ordem das abas.
mod_ciq, mod_ep, mod_comp = st.tabs(["🧫 Módulo CIQ", "🧪 Módulo EP", "⚖️ Comparabilidade"])
nomes_abas = ["📊 Dashboard", "📉 CV", "🎯 Bias", "⚠ Erro Total", "📋 Resultados Mensais"]
if completo:
    nomes_abas.append("🗓 Sigma por Período")
with mod_ciq:
    st.caption("Controle Interno da Qualidade — exports mensais do Infinity.")
    abas = st.tabs(nomes_abas)
tab_dash, tab_grafico, tab_bias, tab_et, tab_tabela = abas[:5]
tab_periodo = abas[5] if completo else None
tab_ep = mod_ep

with mod_comp:
    st.subheader("Comparabilidade — harmonização entre equipamentos")
    st.info(
        "Em construção. Este módulo vai comparar o mesmo teste entre os equipamentos que o realizam "
        "(harmonização). O export do Infinity já traz os controles de comparabilidade (\"COMP...\"), "
        "que hoje ficam fora do CIQ — eles serão a base deste módulo."
    )

# ---------------- DASHBOARD ----------------
with tab_dash:
    st.subheader("Visão geral — piores cenários (CV, Bias e Erro Total)")
    st.caption(
        "Considera os filtros globais da barra lateral (Período e Módulo/Plataforma). "
        "Pra investigar um teste específico em detalhe, use as abas CV, "
        "Bias ou Erro Total."
    )

    status_counts = Counter(df["Status CV"].dropna())
    status_b_counts = Counter(df["Status Bias"].dropna())
    status_e_counts = Counter(df["Status Erro Total"].dropna())

    st.markdown("**CV**")
    metricas_status(status_counts, com_avaliados=True)

    st.markdown("**Bias**")
    metricas_status(status_b_counts, com_avaliados=True)

    st.markdown("**Erro Total**")
    metricas_status(status_e_counts, com_avaliados=True)

    st.divider()
    st.subheader("Top 15 testes — CV fora da meta")
    top_verm = (df[df["Status CV"] == "VERMELHO"]["Teste"]
                .value_counts().head(15).reset_index())
    top_verm.columns = ["Teste", "Ocorrências"]
    st.dataframe(top_verm, hide_index=True, use_container_width=True)

    def bloco_ofensores(titulo, coluna_valor, coluna_maximo, nome_maximo, usa_abs=False):
        st.divider()
        st.subheader(titulo)
        if margem is not None:
            st.caption(
                f"Olha o mês mais recente de cada Teste + Equipamento + Nível, juntando as duas "
                f"faixas de proximidade numa lista só — 🟧 dentro de 5% do {nome_maximo} · "
                f"🟨 dentro de 10% · 🟥 já passou do limite."
            )
        else:
            st.caption(
                f"Olha o mês mais recente de cada Teste + Equipamento + Nível e lista só o que "
                f"passou do {nome_maximo} (margem de proximidade desativada na barra lateral)."
            )
        base = df.dropna(subset=[coluna_maximo, coluna_valor]).copy()
        if base.empty:
            st.info(f"Nenhum registro com {nome_maximo} definido no filtro atual.")
            return
        base["N"] = base["N"].fillna(0)
        idx_maior_n = base.groupby(["Teste", "Equipamento (nome)", "NívelNum", "_ordem_tempo"])["N"].idxmax()
        agg = base.loc[idx_maior_n, ["Teste", "Equipamento (nome)", "NívelNum", "_ordem_tempo",
                                      "Mês/Ano", coluna_valor, coluna_maximo, "N"]]
        idx_recentes = agg.sort_values("_ordem_tempo").groupby(
            ["Teste", "Equipamento (nome)", "NívelNum"]).tail(1).index
        recentes = agg.loc[idx_recentes].copy()
        valor_ref = recentes[coluna_valor].abs() if usa_abs else recentes[coluna_valor]
        recentes["_razao"] = valor_ref / recentes[coluna_maximo]

        def classifica(razao):
            if razao > 1.0:
                return "Vermelho (fora)"
            if margem is None:
                return None
            if razao >= 0.95:
                return "Dentro de 5%"
            if razao >= 0.90:
                return "Dentro de 10%"
            return None

        recentes["Faixa"] = recentes["_razao"].apply(classifica)
        alerta = recentes[recentes["Faixa"].notna()].sort_values("_razao", ascending=False)
        if alerta.empty:
            if margem is None:
                st.success("Nenhum teste acima do limite no momento mais recente.")
            else:
                st.success("Nenhum teste dentro de 10% do limite (ou acima) no momento mais recente.")
            return

        linhas = []
        for _, row in alerta.head(20).iterrows():
            serie = df[(df["Teste"] == row["Teste"]) &
                       (df["Equipamento (nome)"] == row["Equipamento (nome)"]) &
                       (df["NívelNum"] == row["NívelNum"])].copy()
            serie["N"] = serie["N"].fillna(0)
            idx_mes = serie.groupby("_ordem_tempo")["N"].idxmax()
            ultimos = serie.loc[idx_mes, ["_ordem_tempo", "Mês/Ano", coluna_valor]].sort_values("_ordem_tempo")
            linha = {
                "Teste": row["Teste"], "Equipamento": row["Equipamento (nome)"],
                "Nível": int(row["NívelNum"]), nome_maximo: row[coluna_maximo], "Faixa": row["Faixa"],
            }
            for _, mrow in ultimos.iterrows():
                linha[mrow["Mês/Ano"]] = round(mrow[coluna_valor], 2)
            linhas.append(linha)

        traj = pd.DataFrame(linhas)
        col_fixas = {"Teste", "Equipamento", "Nível", nome_maximo, "Faixa"}
        meses_presentes = [c for c in traj.columns if c not in col_fixas]
        ordem_mes = (df[["Mês/Ano", "_ordem_tempo"]].drop_duplicates()
                     .set_index("Mês/Ano")["_ordem_tempo"].to_dict())
        meses_cols = sorted(meses_presentes, key=lambda m: ordem_mes.get(m, 0))
        traj = traj[["Teste", "Equipamento", "Nível", nome_maximo, "Faixa"] + meses_cols]

        def cor_celula(row):
            maximo = row[nome_maximo]
            estilos = []
            for col in row.index:
                if col in meses_cols and pd.notna(row[col]) and pd.notna(maximo):
                    val = abs(row[col]) if usa_abs else row[col]
                    estilos.append(cor_faixa(val, maximo))
                else:
                    estilos.append("")
            return estilos

        styler = (traj.style.apply(cor_celula, axis=1)
                  .format("{:.2f}", subset=[nome_maximo] + meses_cols, na_rep="—"))
        st.dataframe(styler, hide_index=True, use_container_width=True)
        st.caption(
            f"{LEGENDA_FAIXAS} "
            f"{len(alerta)} combinação(ões) no total (mostrando até 20), todos os meses do "
            f"período selecionado na barra lateral."
        )

    bloco_ofensores("Piores cenários — CV", "CV (%)", "CV Máximo", "CV Máximo")
    bloco_ofensores("Piores cenários — Bias", "Bias Observado", "Bias Máximo", "Bias Máximo")
    bloco_ofensores("Piores cenários — Erro Total", "Erro Total Observado", "ETM (para comparação)", "ETM")

    # Sigma e a sugestão técnica ("Tente melhorar CV/Bias") só no modo completo;
    # no simples ficam só os semáforos de CV, Bias e Erro Total acima.
    if completo:
        st.divider()
        st.subheader("Piores cenários — Sigma")
        st.caption(
            "Olha o mês mais recente de cada Teste + Equipamento + Nível e mostra os piores "
            "Sigmas (quanto menor, pior) — junto com o CV e o Bias daquele mês, pra já dar uma "
            "pista se o Sigma baixo veio do CV, do Bias, ou dos dois. "
            "🟥 Sigma < 3 · 🟨 Sigma entre 3 e 6 · 🟩 Sigma ≥ 6."
        )

        df_watch_s = df.dropna(subset=["Sigma Mensal"]).copy()
        if df_watch_s.empty:
            st.info("Nenhum registro com Sigma calculado no filtro atual.")
        else:
            df_watch_s["N"] = df_watch_s["N"].fillna(0)
            idx_maior_n_s = (df_watch_s.groupby(["Teste", "Equipamento (nome)", "NívelNum", "_ordem_tempo"])
                              ["N"].idxmax())
            agg_s = df_watch_s.loc[idx_maior_n_s, ["Teste", "Equipamento (nome)", "NívelNum", "_ordem_tempo",
                                                    "Mês/Ano", "Sigma Mensal", "CV (%)", "Bias (%)", "Bias Observado",
                                                    "CV Máximo", "Bias Máximo"]]
            idx_recentes_s = agg_s.sort_values("_ordem_tempo").groupby(
                ["Teste", "Equipamento (nome)", "NívelNum"]).tail(1).index
            df_recentes_s = agg_s.loc[idx_recentes_s].copy()
            df_alerta_s = df_recentes_s[df_recentes_s["Sigma Mensal"] < 6].sort_values("Sigma Mensal")

            if df_alerta_s.empty:
                st.success("Nenhuma combinação com Sigma abaixo de 6 no mês mais recente do período filtrado.")
            else:
                def causa_provavel(row):
                    cv_frac = (row["CV (%)"] / row["CV Máximo"]) if pd.notna(row["CV Máximo"]) and row["CV Máximo"] else None
                    bias_frac = (row["Bias Observado"] / row["Bias Máximo"]) if pd.notna(row["Bias Máximo"]) and row["Bias Máximo"] else None
                    if cv_frac is None and bias_frac is None:
                        return "—"
                    if bias_frac is None or (cv_frac is not None and cv_frac >= bias_frac * 1.2):
                        return "Tente melhorar o CV"
                    if cv_frac is None or (bias_frac > cv_frac * 1.2):
                        return "Tente melhorar o Viés (Bias)"
                    return "Tente melhorar CV e Viés"

                df_alerta_s["Sugestão"] = df_alerta_s.apply(causa_provavel, axis=1)
                tabela_s = df_alerta_s.head(20)[[
                    "Teste", "Equipamento (nome)", "NívelNum", "Mês/Ano", "Sigma Mensal",
                    "CV (%)", "CV Máximo", "Bias (%)", "Bias Máximo", "Sugestão",
                ]].rename(columns={"Equipamento (nome)": "Equipamento", "NívelNum": "Nível"})
                tabela_s["Nível"] = tabela_s["Nível"].astype(int)

                styler_s = (tabela_s.style.map(cor_sigma, subset=["Sigma Mensal"])
                            .format("{:.2f}", subset=["Sigma Mensal", "CV (%)", "CV Máximo", "Bias (%)", "Bias Máximo"],
                                    na_rep="—"))
                st.dataframe(styler_s, hide_index=True, use_container_width=True)
                st.caption(
                    f"{len(df_alerta_s)} combinação(ões) com Sigma < 6 no total (mostrando até 20). "
                    "'Sugestão' compara o quanto do limite o CV consumiu vs. o quanto o Bias "
                    "consumiu — quem consumiu proporcionalmente mais é apontado como principal razão."
                )

# ---------------- CV ----------------
with tab_grafico:
    card_pior_cenario(df[df["Teste"] == teste_global], "CV (%)", "CV")

    modo = st.radio(
        "Comparar por:", ["Equipamento (mesmo teste/nível, entre equipamentos)",
                           "Nível (mesmo teste/equipamento, entre níveis)"],
        horizontal=True, key="modo_comparacao",
    )

    if modo.startswith("Equipamento"):
        st.subheader("Tendência de CV mensal por teste — todos os equipamentos")

        teste_sel = teste_global
        st.caption(f"Teste: **{teste_sel}** (mude na barra lateral, em 'Teste em foco')")

        df_teste = df[df["Teste"] == teste_sel]
        niveis_disponiveis = sorted(df_teste["NívelNum"].dropna().unique())
        if not niveis_disponiveis:
            st.info("Nenhum nível identificado para esse teste.")
            st.stop()
        nivel_sel = st.selectbox("Nível", [int(n) for n in niveis_disponiveis], key="tend_nivel")

        df_nivel = df_teste[df_teste["NívelNum"] == nivel_sel]

        equip_disponiveis = sorted(df_nivel["Equipamento (nome)"].dropna().unique())
        equip_sel = st.multiselect(
            "Equipamentos (filtro)", equip_disponiveis,
            default=equip_disponiveis, key="grafico_equip",
        )

        if not equip_sel:
            st.info("Selecione ao menos um equipamento.")
        else:
            df_plot = df_nivel[df_nivel["Equipamento (nome)"].isin(equip_sel)]

            fig = go.Figure()
            for i, equip in enumerate(equip_sel):
                serie_raw = df_plot[df_plot["Equipamento (nome)"] == equip].copy()
                serie_raw["N"] = serie_raw["N"].fillna(0)
                # se houver mais de um lote no mesmo mês, usa o de maior N
                idx_mes = serie_raw.groupby("_ordem_tempo")["N"].idxmax()
                serie = (serie_raw.loc[idx_mes, ["_ordem_tempo", "Mês/Ano", "CV (%)", "CV Máximo"]]
                         .sort_values("_ordem_tempo"))
                if serie.empty:
                    continue
                cor = PALETA[i % len(PALETA)]
                fig.add_trace(go.Scatter(
                    x=serie["Mês/Ano"], y=serie["CV (%)"],
                    mode="lines+markers", name=f"{equip} - CV mensal",
                    line=dict(color=cor, width=2), marker=dict(size=6),
                ))
                cvmax = serie["CV Máximo"].dropna()
                if not cvmax.empty:
                    fig.add_trace(go.Scatter(
                        x=serie["Mês/Ano"], y=[cvmax.iloc[-1]] * len(serie),
                        mode="lines", name=f"{equip} - CV Máximo",
                        line=dict(color=cor, width=1.2, dash="dash"), opacity=0.6,
                    ))

            fig.update_layout(
                xaxis_title="Mês/Ano", **EIXO_MESES, yaxis_title="CV (%)",
                height=560, hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                margin=dict(t=60),
            )
            st.plotly_chart(fig, use_container_width=True)
            st.caption(
                "Linha cheia = CV mensal do equipamento. Linha tracejada da mesma cor = "
                "CV Máximo daquele equipamento (pode variar entre plataformas para o mesmo teste)."
            )

    else:
        st.subheader("Comparação entre níveis — mesmo teste, mesmo equipamento")
        st.caption(
            "Se todos os níveis se movem juntos (mesma direção, mesma época), costuma apontar "
            "para calibração. Se só um nível específico foge do padrão enquanto os outros ficam "
            "estáveis, é mais sinal de problema pontual (material daquele nível, interferência, "
            "ou o próprio equipamento numa faixa específica de concentração)."
        )

        teste_sel_n = teste_global
        st.caption(f"Teste: **{teste_sel_n}** (mude na barra lateral, em 'Teste em foco')")
        df_teste_n = df[df["Teste"] == teste_sel_n]
        equip_n_opts = sorted(df_teste_n["Equipamento (nome)"].dropna().unique())
        if not equip_n_opts:
            st.info("Nenhum equipamento disponível para esse teste.")
        else:
            equip_sel_n = st.selectbox("Equipamento", equip_n_opts, key="niveis_equip")
            df_equip_n = df_teste_n[df_teste_n["Equipamento (nome)"] == equip_sel_n]
            niveis_n = sorted(df_equip_n["NívelNum"].dropna().unique())

            if len(niveis_n) < 2:
                st.warning(
                    f"Esse teste só tem {len(niveis_n)} nível registrado nesse equipamento — "
                    "não dá para comparar entre níveis."
                )
            else:
                def serie_por_nivel(df_base, nivel, coluna):
                    serie_raw = df_base[df_base["NívelNum"] == nivel].copy()
                    serie_raw["N"] = serie_raw["N"].fillna(0)
                    idx_mes = serie_raw.groupby("_ordem_tempo")["N"].idxmax()
                    return (serie_raw.loc[idx_mes, ["_ordem_tempo", "Mês/Ano", coluna]]
                            .sort_values("_ordem_tempo"))

                colg1, colg2 = st.columns(2)
                with colg1:
                    st.markdown("**Bias (%) por nível**")
                    fig_n_bias = go.Figure()
                    for i, nivel in enumerate(niveis_n):
                        serie = serie_por_nivel(df_equip_n, nivel, "Bias (%)")
                        if serie.empty:
                            continue
                        fig_n_bias.add_trace(go.Scatter(
                            x=serie["Mês/Ano"], y=serie["Bias (%)"],
                            mode="lines+markers", name=f"Nível {int(nivel)}",
                            line=dict(color=PALETA[i % len(PALETA)], width=2), marker=dict(size=6),
                        ))
                    fig_n_bias.add_hline(y=0, line=dict(color="#999999", width=1))
                    fig_n_bias.update_layout(
                        xaxis_title="Mês/Ano", **EIXO_MESES, yaxis_title="Bias (%)", height=450,
                        hovermode="x unified",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                        margin=dict(t=50),
                    )
                    st.plotly_chart(fig_n_bias, use_container_width=True)

                with colg2:
                    st.markdown("**CV (%) por nível**")
                    fig_n_cv = go.Figure()
                    for i, nivel in enumerate(niveis_n):
                        serie = serie_por_nivel(df_equip_n, nivel, "CV (%)")
                        if serie.empty:
                            continue
                        fig_n_cv.add_trace(go.Scatter(
                            x=serie["Mês/Ano"], y=serie["CV (%)"],
                            mode="lines+markers", name=f"Nível {int(nivel)}",
                            line=dict(color=PALETA[i % len(PALETA)], width=2), marker=dict(size=6),
                        ))
                    fig_n_cv.update_layout(
                        xaxis_title="Mês/Ano", **EIXO_MESES, yaxis_title="CV (%)", height=450,
                        hovermode="x unified",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                        margin=dict(t=50),
                    )
                    st.plotly_chart(fig_n_cv, use_container_width=True)

                st.info(
                    "💡 Leitura rápida: Bias subindo/descendo junto em todos os níveis no mesmo mês "
                    "→ suspeita de calibração. CV subindo só num nível → suspeita de material de "
                    "controle daquele nível ou problema pontual do equipamento naquela faixa."
                )

    st.divider()
    st.subheader("CV mensal, Média mensal e CV Pooled")
    st.caption(
        "Igual à planilha de desempenho: CV de cada mês por nível, a Média de cada mês por "
        "nível, e o CV Pooled (RMS = raiz da média dos quadrados dos CVs mensais) resumindo "
        "o período filtrado na barra lateral."
    )

    teste_cv_tabela = teste_global
    st.caption(f"Teste: **{teste_cv_tabela}** (mude na barra lateral, em 'Teste em foco')")
    df_teste_cvtab = df[df["Teste"] == teste_cv_tabela].dropna(subset=["NívelNum"]).copy()

    if df_teste_cvtab.empty:
        st.info("Nenhum registro para esse teste no filtro atual.")
    else:
        df_teste_cvtab["N"] = df_teste_cvtab["N"].fillna(0)
        idx_maior_n_ct = (df_teste_cvtab.groupby(["Equipamento (nome)", "NívelNum", "_ordem_tempo"])
                           ["N"].idxmax())
        base_ct = df_teste_cvtab.loc[idx_maior_n_ct]

        ordem_mes_ct = (df[["Mês/Ano", "_ordem_tempo"]].drop_duplicates()
                        .set_index("Mês/Ano")["_ordem_tempo"].to_dict())

        # --- Tabela de CV mensal + Pooled ---
        linhas_cv = []
        for (equip, nivel), grupo in base_ct.groupby(["Equipamento (nome)", "NívelNum"]):
            grupo = grupo.sort_values("_ordem_tempo")
            cvmax = grupo["CV Máximo"].dropna()
            cvmax_val = cvmax.iloc[0] if not cvmax.empty else None
            linha = {"Equipamento": equip, "Nível": int(nivel), "CV Máximo": cvmax_val}
            for _, r in grupo.iterrows():
                if pd.notna(r["CV (%)"]):
                    linha[r["Mês/Ano"]] = round(r["CV (%)"], 2)
            grupo_valido = grupo.dropna(subset=["CV (%)"])
            if len(grupo_valido) >= 2:
                cvs_validos = grupo_valido["CV (%)"]
                linha["CV Pooled (%)"] = round((sum(v ** 2 for v in cvs_validos) / len(cvs_validos)) ** 0.5, 2)
            else:
                linha["CV Pooled (%)"] = None
            if not grupo_valido.empty:
                idx_pior = grupo_valido["CV (%)"].idxmax()
                linha["Pior CV (%)"] = round(grupo_valido.loc[idx_pior, "CV (%)"], 2)
                linha["Mês do pior"] = grupo_valido.loc[idx_pior, "Mês/Ano"]
            else:
                linha["Pior CV (%)"] = None
                linha["Mês do pior"] = None
            linhas_cv.append(linha)

        df_cv_tab = pd.DataFrame(linhas_cv)
        col_fixas_ct = {"Equipamento", "Nível", "CV Máximo", "CV Pooled (%)", "Pior CV (%)", "Mês do pior"}
        meses_ct = sorted([c for c in df_cv_tab.columns if c not in col_fixas_ct],
                           key=lambda m: ordem_mes_ct.get(m, 0))
        df_cv_tab = df_cv_tab[["Equipamento", "Nível", "CV Máximo", "CV Pooled (%)", "Pior CV (%)",
                                "Mês do pior"] + meses_ct]
        df_cv_tab = df_cv_tab.sort_values(["Equipamento", "Nível"])

        def cor_cv_tab(row):
            cvmax = row["CV Máximo"]
            estilos = []
            for col in row.index:
                if col in (meses_ct + ["CV Pooled (%)", "Pior CV (%)"]) and pd.notna(row[col]) and pd.notna(cvmax):
                    val = row[col]
                    estilos.append(cor_faixa(val, cvmax))
                else:
                    estilos.append("")
            return estilos

        st.markdown("**CV mensal (%) por nível** — CV Pooled e pior mês em destaque, logo após o CV Máximo")
        styler_cv_tab = (df_cv_tab.style.apply(cor_cv_tab, axis=1)
                          .format("{:.2f}", subset=["CV Máximo"] + meses_ct + ["CV Pooled (%)", "Pior CV (%)"],
                                  na_rep="—"))
        st.dataframe(styler_cv_tab, hide_index=True, use_container_width=True)
        st.caption(LEGENDA_FAIXAS)

        # --- Tabela de Média mensal ---
        linhas_media_ct = []
        for (equip, nivel), grupo in base_ct.groupby(["Equipamento (nome)", "NívelNum"]):
            grupo = grupo.sort_values("_ordem_tempo")
            linha = {"Equipamento": equip, "Nível": int(nivel)}
            for _, r in grupo.iterrows():
                if pd.notna(r["Média"]):
                    linha[r["Mês/Ano"]] = round(r["Média"], 2)
            linhas_media_ct.append(linha)

        df_media_tab = pd.DataFrame(linhas_media_ct)
        col_fixas_mt = {"Equipamento", "Nível"}
        meses_mt = sorted([c for c in df_media_tab.columns if c not in col_fixas_mt],
                           key=lambda m: ordem_mes_ct.get(m, 0))
        df_media_tab = df_media_tab[["Equipamento", "Nível"] + meses_mt]
        df_media_tab = df_media_tab.sort_values(["Equipamento", "Nível"])

        st.markdown("**Média mensal do CIQ por nível**")
        styler_media_tab = df_media_tab.style.format("{:.2f}", subset=meses_mt, na_rep="—")
        st.dataframe(styler_media_tab, hide_index=True, use_container_width=True)

# ---------------- BIAS ----------------
with tab_bias:
    st.subheader("Bias mensal x Valor Alvo x Bias Máximo (ESM)")
    st.caption(
        "Bias recalculado como (Média − Valor Alvo) / Valor Alvo × 100, comparado com o "
        "Bias Máximo (ESM) da especificação. Quando o teste usa critério absoluto (concentração "
        "abaixo do cutoff), a comparação passa a ser em valor absoluto. Pra ver os piores "
        "cenários de todos os testes juntos, use a aba Dashboard."
    )

    teste_sel_b = teste_global
    st.caption(f"Teste: **{teste_sel_b}** (mude na barra lateral, em 'Teste em foco')")
    df_teste_b = df[df["Teste"] == teste_sel_b]

    card_pior_cenario(df_teste_b, "Bias Observado", "Bias")

    modo_b = st.radio(
        "Comparar por:", ["Equipamento (mesmo nível, entre equipamentos)",
                           "Nível (mesmo equipamento, entre níveis)"],
        horizontal=True, key="modo_bias",
    )

    def tabela_media_alvo(df_plot, equip_ou_nivel_sel, rotulo_linha_fn):
        st.markdown("**Média mensal do CIQ x Valor Alvo (Config. Infinity)**")
        linhas_media = []
        for chave in equip_ou_nivel_sel:
            serie_raw = rotulo_linha_fn(df_plot, chave).copy()
            serie_raw["N"] = serie_raw["N"].fillna(0)
            idx_mes = serie_raw.groupby("_ordem_tempo")["N"].idxmax()
            serie = (serie_raw.loc[idx_mes, ["_ordem_tempo", "Mês/Ano", "Média", "Config. valor alvo"]]
                     .sort_values("_ordem_tempo"))
            if serie.empty:
                continue
            linha_media = {"": chave, "Métrica": "Média CIQ"}
            linha_alvo = {"": chave, "Métrica": "Valor Alvo"}
            for _, mrow in serie.iterrows():
                linha_media[mrow["Mês/Ano"]] = mrow["Média"]
                linha_alvo[mrow["Mês/Ano"]] = mrow["Config. valor alvo"]
            linhas_media.append(linha_media)
            linhas_media.append(linha_alvo)

        if linhas_media:
            df_media = pd.DataFrame(linhas_media)
            col_fixas_m = {"", "Métrica"}
            meses_presentes_m = [c for c in df_media.columns if c not in col_fixas_m]
            ordem_mes_m = (df[["Mês/Ano", "_ordem_tempo"]].drop_duplicates()
                           .set_index("Mês/Ano")["_ordem_tempo"].to_dict())
            meses_cols_m = sorted(meses_presentes_m, key=lambda m: ordem_mes_m.get(m, 0))
            df_media = df_media[["", "Métrica"] + meses_cols_m]
            styler_media = df_media.style.format("{:.2f}", subset=meses_cols_m, na_rep="—")
            st.dataframe(styler_media, hide_index=True, use_container_width=True)
            st.caption(
                "Valores usados no cálculo do Bias mensal (Bias % = (Média − Valor Alvo) "
                "/ Valor Alvo × 100). O Valor Alvo pode variar de mês a mês quando o lote "
                "do controle muda."
            )

    if modo_b.startswith("Equipamento"):
        niveis_b = sorted(df_teste_b["NívelNum"].dropna().unique())
        if not niveis_b:
            st.info("Nenhum nível identificado para esse teste.")
        else:
            nivel_sel_b = st.selectbox("Nível", [int(n) for n in niveis_b], key="bias_nivel")
            df_nivel_b = df_teste_b[df_teste_b["NívelNum"] == nivel_sel_b]

            equip_b = sorted(df_nivel_b["Equipamento (nome)"].dropna().unique())
            equip_sel_b = st.multiselect("Equipamentos (filtro)", equip_b, default=equip_b, key="bias_equip")

            if not equip_sel_b:
                st.info("Selecione ao menos um equipamento.")
            else:
                df_plot_b = df_nivel_b[df_nivel_b["Equipamento (nome)"].isin(equip_sel_b)]
                status_b = Counter(df_plot_b["Status Bias"].dropna())
                metricas_status(status_b)

                fig_b = go.Figure()
                for i, equip in enumerate(equip_sel_b):
                    serie_raw = df_plot_b[df_plot_b["Equipamento (nome)"] == equip].copy()
                    serie_raw["N"] = serie_raw["N"].fillna(0)
                    idx_mes = serie_raw.groupby("_ordem_tempo")["N"].idxmax()
                    serie = (serie_raw.loc[idx_mes, ["_ordem_tempo", "Mês/Ano", "Bias Observado (sinal)",
                                                      "Bias Máximo"]]
                             .sort_values("_ordem_tempo"))
                    if serie.empty:
                        continue
                    cor = PALETA[i % len(PALETA)]
                    fig_b.add_trace(go.Scatter(
                        x=serie["Mês/Ano"], y=serie["Bias Observado (sinal)"],
                        mode="lines+markers", name=f"{equip} - Bias",
                        line=dict(color=cor, width=2), marker=dict(size=6),
                    ))
                    bmax = serie["Bias Máximo"].dropna()
                    if not bmax.empty:
                        fig_b.add_trace(go.Scatter(
                            x=serie["Mês/Ano"], y=[bmax.iloc[-1]] * len(serie),
                            mode="lines", name=f"{equip} - Bias Máximo (+)",
                            line=dict(color=cor, width=1.2, dash="dash"), opacity=0.6,
                        ))
                        fig_b.add_trace(go.Scatter(
                            x=serie["Mês/Ano"], y=[-bmax.iloc[-1]] * len(serie),
                            mode="lines", name=f"{equip} - Bias Máximo (−)",
                            line=dict(color=cor, width=1.2, dash="dash"), opacity=0.6,
                            showlegend=False,
                        ))
                fig_b.add_hline(y=0, line=dict(color="#999999", width=1))
                fig_b.update_layout(
                    xaxis_title="Mês/Ano", **EIXO_MESES, yaxis_title="Bias (%) ou valor absoluto",
                    height=560, hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                    margin=dict(t=60),
                )
                st.plotly_chart(fig_b, use_container_width=True)
                st.caption(
                    "Linha cheia = Bias mensal do equipamento. Linhas tracejadas = faixa de "
                    "Bias Máximo (+/-) daquele equipamento."
                )

                tabela_media_alvo(
                    df_plot_b, equip_sel_b,
                    lambda dfp, equip: dfp[dfp["Equipamento (nome)"] == equip],
                )

    else:
        equip_n_opts_b = sorted(df_teste_b["Equipamento (nome)"].dropna().unique())
        if not equip_n_opts_b:
            st.info("Nenhum equipamento disponível para esse teste.")
        else:
            equip_sel_nb = st.selectbox("Equipamento", equip_n_opts_b, key="bias_equip_niveis")
            df_equip_nb = df_teste_b[df_teste_b["Equipamento (nome)"] == equip_sel_nb]
            niveis_nb = sorted(df_equip_nb["NívelNum"].dropna().unique())

            if len(niveis_nb) < 2:
                st.warning(
                    f"Esse teste só tem {len(niveis_nb)} nível registrado nesse equipamento — "
                    "não dá para comparar entre níveis."
                )
            else:
                status_b = Counter(df_equip_nb["Status Bias"].dropna())
                metricas_status(status_b)

                fig_b = go.Figure()
                for i, nivel in enumerate(niveis_nb):
                    serie_raw = df_equip_nb[df_equip_nb["NívelNum"] == nivel].copy()
                    serie_raw["N"] = serie_raw["N"].fillna(0)
                    idx_mes = serie_raw.groupby("_ordem_tempo")["N"].idxmax()
                    serie = (serie_raw.loc[idx_mes, ["_ordem_tempo", "Mês/Ano", "Bias Observado (sinal)",
                                                      "Bias Máximo"]]
                             .sort_values("_ordem_tempo"))
                    if serie.empty:
                        continue
                    cor = PALETA[i % len(PALETA)]
                    fig_b.add_trace(go.Scatter(
                        x=serie["Mês/Ano"], y=serie["Bias Observado (sinal)"],
                        mode="lines+markers", name=f"Nível {int(nivel)} - Bias",
                        line=dict(color=cor, width=2), marker=dict(size=6),
                    ))
                    bmax = serie["Bias Máximo"].dropna()
                    if not bmax.empty:
                        fig_b.add_trace(go.Scatter(
                            x=serie["Mês/Ano"], y=[bmax.iloc[-1]] * len(serie),
                            mode="lines", name=f"Nível {int(nivel)} - Bias Máximo (+)",
                            line=dict(color=cor, width=1.2, dash="dash"), opacity=0.6,
                        ))
                        fig_b.add_trace(go.Scatter(
                            x=serie["Mês/Ano"], y=[-bmax.iloc[-1]] * len(serie),
                            mode="lines", name=f"Nível {int(nivel)} - Bias Máximo (−)",
                            line=dict(color=cor, width=1.2, dash="dash"), opacity=0.6,
                            showlegend=False,
                        ))
                fig_b.add_hline(y=0, line=dict(color="#999999", width=1))
                fig_b.update_layout(
                    xaxis_title="Mês/Ano", **EIXO_MESES, yaxis_title="Bias (%) ou valor absoluto",
                    height=560, hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                    margin=dict(t=60),
                )
                st.plotly_chart(fig_b, use_container_width=True)
                st.caption(
                    "Linha cheia = Bias mensal daquele nível. Linhas tracejadas = faixa de "
                    "Bias Máximo (+/-). Níveis se movendo juntos → suspeita de calibração; "
                    "só um nível fora do padrão → problema pontual."
                )

                niveis_label = [f"Nível {int(n)}" for n in niveis_nb]
                tabela_media_alvo(
                    df_equip_nb, niveis_label,
                    lambda dfp, lbl: dfp[dfp["NívelNum"] == int(lbl.replace("Nível ", ""))],
                )

    st.divider()
    st.subheader("Bias mensal — todos os meses")
    st.caption(
        "Igual à planilha de desempenho: mostra o Bias de cada mês por nível, coloridos, "
        "não só os meses que estão ruins."
    )

    teste_bias_tabela = teste_global
    st.caption(f"Teste: **{teste_bias_tabela}** (mude na barra lateral, em 'Teste em foco')")
    df_teste_btab = df[df["Teste"] == teste_bias_tabela].dropna(subset=["NívelNum"]).copy()

    if df_teste_btab.empty:
        st.info("Nenhum registro para esse teste no filtro atual.")
    else:
        df_teste_btab["N"] = df_teste_btab["N"].fillna(0)
        idx_maior_n_bt = (df_teste_btab.groupby(["Equipamento (nome)", "NívelNum", "_ordem_tempo"])
                           ["N"].idxmax())
        base_bt = df_teste_btab.loc[idx_maior_n_bt]

        ordem_mes_bt = (df[["Mês/Ano", "_ordem_tempo"]].drop_duplicates()
                        .set_index("Mês/Ano")["_ordem_tempo"].to_dict())

        linhas_bt = []
        for (equip, nivel), grupo in base_bt.groupby(["Equipamento (nome)", "NívelNum"]):
            grupo = grupo.sort_values("_ordem_tempo")
            bmax = grupo["Bias Máximo"].dropna()
            bmax_val = bmax.iloc[0] if not bmax.empty else None
            linha = {"Equipamento": equip, "Nível": int(nivel), "Bias Máximo": bmax_val}
            for _, r in grupo.iterrows():
                if pd.notna(r["Bias Observado (sinal)"]):
                    linha[r["Mês/Ano"]] = round(r["Bias Observado (sinal)"], 2)
            grupo_valido_b = grupo.dropna(subset=["Bias Observado (sinal)"])
            if not grupo_valido_b.empty:
                linha["Bias Médio"] = round(grupo_valido_b["Bias Observado (sinal)"].mean(), 2)
                idx_pior_b = grupo_valido_b["Bias Observado"].idxmax()
                linha["Pior Bias"] = round(grupo_valido_b.loc[idx_pior_b, "Bias Observado (sinal)"], 2)
                linha["Mês do pior"] = grupo_valido_b.loc[idx_pior_b, "Mês/Ano"]
            else:
                linha["Bias Médio"] = None
                linha["Pior Bias"] = None
                linha["Mês do pior"] = None
            linhas_bt.append(linha)

        df_bias_tab = pd.DataFrame(linhas_bt)
        col_fixas_bt = {"Equipamento", "Nível", "Bias Máximo", "Bias Médio", "Pior Bias", "Mês do pior"}
        meses_bt = sorted([c for c in df_bias_tab.columns if c not in col_fixas_bt],
                           key=lambda m: ordem_mes_bt.get(m, 0))
        df_bias_tab = df_bias_tab[["Equipamento", "Nível", "Bias Máximo", "Bias Médio", "Pior Bias",
                                    "Mês do pior"] + meses_bt]
        df_bias_tab = df_bias_tab.sort_values(["Equipamento", "Nível"])

        def cor_bias_tab(row):
            bmax = row["Bias Máximo"]
            estilos = []
            for col in row.index:
                if col in (meses_bt + ["Bias Médio", "Pior Bias"]) and pd.notna(row[col]) and pd.notna(bmax):
                    val = abs(row[col])
                    estilos.append(cor_faixa(val, bmax))
                else:
                    estilos.append("")
            return estilos

        styler_bias_tab = (df_bias_tab.style.apply(cor_bias_tab, axis=1)
                            .format("{:.2f}", subset=["Bias Máximo", "Bias Médio", "Pior Bias"] + meses_bt,
                                    na_rep="—"))
        st.dataframe(styler_bias_tab, hide_index=True, use_container_width=True)
        st.caption(LEGENDA_FAIXAS)

# ---------------- ERRO TOTAL ----------------
with tab_et:
    st.subheader("Erro Total Observado x ETM")
    st.caption(
        "Erro Total Observado = |Bias%| + 1,65 × CV% (padrão CLIA/RCPA), comparado com o "
        "ETM da especificação. No critério absoluto, usa |Média − Alvo| + 1,65 × DP. Pra ver "
        "os piores cenários de todos os testes juntos, use a aba Dashboard."
    )

    teste_sel_e = teste_global
    st.caption(f"Teste: **{teste_sel_e}** (mude na barra lateral, em 'Teste em foco')")
    df_teste_e = df[df["Teste"] == teste_sel_e]

    card_pior_cenario(df_teste_e, "Erro Total Observado", "Erro Total")

    modo_e = st.radio(
        "Comparar por:", ["Equipamento (mesmo nível, entre equipamentos)",
                           "Nível (mesmo equipamento, entre níveis)"],
        horizontal=True, key="modo_et",
    )

    if modo_e.startswith("Equipamento"):
        niveis_e = sorted(df_teste_e["NívelNum"].dropna().unique())
        if not niveis_e:
            st.info("Nenhum nível identificado para esse teste.")
        else:
            nivel_sel_e = st.selectbox("Nível", [int(n) for n in niveis_e], key="et_nivel")
            df_nivel_e = df_teste_e[df_teste_e["NívelNum"] == nivel_sel_e]

            equip_e = sorted(df_nivel_e["Equipamento (nome)"].dropna().unique())
            equip_sel_e = st.multiselect("Equipamentos (filtro)", equip_e, default=equip_e, key="et_equip")

            if not equip_sel_e:
                st.info("Selecione ao menos um equipamento.")
            else:
                df_plot_e = df_nivel_e[df_nivel_e["Equipamento (nome)"].isin(equip_sel_e)]
                status_e = Counter(df_plot_e["Status Erro Total"].dropna())
                metricas_status(status_e)

                fig_e = go.Figure()
                for i, equip in enumerate(equip_sel_e):
                    serie_raw = df_plot_e[df_plot_e["Equipamento (nome)"] == equip].copy()
                    serie_raw["N"] = serie_raw["N"].fillna(0)
                    idx_mes = serie_raw.groupby("_ordem_tempo")["N"].idxmax()
                    serie = (serie_raw.loc[idx_mes, ["_ordem_tempo", "Mês/Ano", "Erro Total Observado",
                                                      "ETM (para comparação)"]]
                             .sort_values("_ordem_tempo"))
                    if serie.empty:
                        continue
                    cor = PALETA[i % len(PALETA)]
                    fig_e.add_trace(go.Scatter(
                        x=serie["Mês/Ano"], y=serie["Erro Total Observado"],
                        mode="lines+markers", name=f"{equip} - Erro Total Observado",
                        line=dict(color=cor, width=2), marker=dict(size=6),
                    ))
                    etm = serie["ETM (para comparação)"].dropna()
                    if not etm.empty:
                        fig_e.add_trace(go.Scatter(
                            x=serie["Mês/Ano"], y=[etm.iloc[-1]] * len(serie),
                            mode="lines", name=f"{equip} - ETM",
                            line=dict(color=cor, width=1.2, dash="dash"), opacity=0.6,
                        ))
                fig_e.update_layout(
                    xaxis_title="Mês/Ano", **EIXO_MESES, yaxis_title="Erro Total (%) ou valor absoluto",
                    height=560, hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                    margin=dict(t=60),
                )
                st.plotly_chart(fig_e, use_container_width=True)
                st.caption(
                    "Linha cheia = Erro Total Observado mensal. Linha tracejada = ETM daquele equipamento."
                )

    else:
        equip_n_opts_e = sorted(df_teste_e["Equipamento (nome)"].dropna().unique())
        if not equip_n_opts_e:
            st.info("Nenhum equipamento disponível para esse teste.")
        else:
            equip_sel_ne = st.selectbox("Equipamento", equip_n_opts_e, key="et_equip_niveis")
            df_equip_ne = df_teste_e[df_teste_e["Equipamento (nome)"] == equip_sel_ne]
            niveis_ne = sorted(df_equip_ne["NívelNum"].dropna().unique())

            if len(niveis_ne) < 2:
                st.warning(
                    f"Esse teste só tem {len(niveis_ne)} nível registrado nesse equipamento — "
                    "não dá para comparar entre níveis."
                )
            else:
                status_e = Counter(df_equip_ne["Status Erro Total"].dropna())
                metricas_status(status_e)

                fig_e = go.Figure()
                for i, nivel in enumerate(niveis_ne):
                    serie_raw = df_equip_ne[df_equip_ne["NívelNum"] == nivel].copy()
                    serie_raw["N"] = serie_raw["N"].fillna(0)
                    idx_mes = serie_raw.groupby("_ordem_tempo")["N"].idxmax()
                    serie = (serie_raw.loc[idx_mes, ["_ordem_tempo", "Mês/Ano", "Erro Total Observado",
                                                      "ETM (para comparação)"]]
                             .sort_values("_ordem_tempo"))
                    if serie.empty:
                        continue
                    cor = PALETA[i % len(PALETA)]
                    fig_e.add_trace(go.Scatter(
                        x=serie["Mês/Ano"], y=serie["Erro Total Observado"],
                        mode="lines+markers", name=f"Nível {int(nivel)} - Erro Total Observado",
                        line=dict(color=cor, width=2), marker=dict(size=6),
                    ))
                    etm = serie["ETM (para comparação)"].dropna()
                    if not etm.empty:
                        fig_e.add_trace(go.Scatter(
                            x=serie["Mês/Ano"], y=[etm.iloc[-1]] * len(serie),
                            mode="lines", name=f"Nível {int(nivel)} - ETM",
                            line=dict(color=cor, width=1.2, dash="dash"), opacity=0.6,
                        ))
                fig_e.update_layout(
                    xaxis_title="Mês/Ano", **EIXO_MESES, yaxis_title="Erro Total (%) ou valor absoluto",
                    height=560, hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                    margin=dict(t=60),
                )
                st.plotly_chart(fig_e, use_container_width=True)
                st.caption(
                    "Linha cheia = Erro Total Observado mensal daquele nível. Linha tracejada = "
                    "ETM. Níveis se movendo juntos → suspeita de calibração; só um nível fora "
                    "do padrão → problema pontual."
                )

# ---------------- TABELA COMPLETA ----------------
with tab_tabela:
    st.subheader("Resultados mensais completos")

    colf1, colf2, colf3, colf4 = st.columns(4)
    f_status = colf1.multiselect("Status CV", ["VERDE", "AMARELO", "VERMELHO"] if margem is not None
                                 else ["VERDE", "VERMELHO"])
    f_teste = colf2.multiselect("Teste", sorted(df["Teste"].dropna().unique()))
    f_equip_tab = colf3.multiselect("Equipamento", sorted(df["Equipamento (nome)"].dropna().unique()))
    f_tend = colf4.checkbox("Só com tendência de alta")

    df_filtrado = df.copy()
    if f_status:
        df_filtrado = df_filtrado[df_filtrado["Status CV"].isin(f_status)]
    if f_teste:
        df_filtrado = df_filtrado[df_filtrado["Teste"].isin(f_teste)]
    if f_equip_tab:
        df_filtrado = df_filtrado[df_filtrado["Equipamento (nome)"].isin(f_equip_tab)]
    if f_tend:
        df_filtrado = df_filtrado[df_filtrado["Tendência CV"] != "—"]

    cols_show = ["Mês/Ano", "Tipo", "Equipamento (nome)", "Teste", "Spec - Analito", "Nível", "NívelNum",
                 "NívelOrigem", "Número de lote", "N", "Config. valor alvo", "Média", "CV (%)", "CV Máximo",
                 "Status CV", "Tendência CV", "Bias (%)", "Sigma Mensal", "Sigma Mínimo", "Status Sigma",
                 "Classificação Sigma (Westgard)", "Critério Sigma"]
    if not completo:
        cols_show = [c for c in cols_show if "Sigma" not in c]
    exibe = df_filtrado[cols_show].rename(columns={"Equipamento (nome)": "Equipamento"})
    exibe["Mês/Ano"] = pd.to_datetime(dict(year=df_filtrado["Ano"], month=df_filtrado["Mês"], day=1))
    st.dataframe(exibe.sort_values("Mês/Ano", ascending=False, kind="stable"), hide_index=True,
                 use_container_width=True, height=500,
                 column_config={"Mês/Ano": st.column_config.DateColumn("Mês/Ano", format="MM/YYYY")})
    st.caption(f"{len(df_filtrado)} de {len(df)} registros exibidos")

    csv = df_filtrado[cols_show].to_csv(index=False).encode("utf-8-sig")
    st.download_button("Baixar CSV filtrado", csv, "ciq_resultados.csv", "text/csv")

# ---------------- EP (ENSAIO DE PROFICIÊNCIA) ----------------
ARQ_DEPARA_EP = REF_DIR / "ep_depara_testes.xlsx"
# equipamento informado pelo ControlLab → "Tipo" do equipamento no CIQ (equipamentos.xlsx)
PLATAFORMA_EP = [("c702", "BIOQ"), ("c503", "BIOQ"), ("cobas c", "BIOQ"), ("e801", "IMUNO"), ("cobas e", "IMUNO"),
                 ("atellica", "ATELLICA"), ("liaison", "LIAISON"), ("bn ii", "BN"), ("immulite", "IMMULITE"),
                 ("maglumi", "MAGLUMI"), ("diestro", "DIESTRO"), ("c513", "C513")]
CHAVE_RODADA = ["Provedor", "Programa", "Rodada", "Mneumonico", "Sistema"]
MOSTRAR_DEPARA_EP = False   # situação do de-para escondida enquanto a ferramenta é validada
# opções de cálculo do Sigma do EP (rótulo na tela → valor do ep_calculo)
OPCOES_VIES_EP = {"Automático": "Automático", "Médio (%)": "Médio", "Regressão": "Regressão"}
OPCOES_FORMULA_EP = {"Automático": "Automático", "σ %": "σ %", "σ absoluto": "σ abs"}
res_ep_global = pd.DataFrame()   # resultados do EP por amostra, usados também na aba de Sigma


def _mtime(caminho):
    return caminho.stat().st_mtime_ns if caminho.exists() else 0


# Obs.: o st.cache_data ignora parâmetros que começam com "_"; a assinatura (data do arquivo) precisa
# de nome sem "_" pra que o cache seja refeito quando o arquivo muda.
@st.cache_data(show_spinner=False)
def carrega_depara_ep(assinatura, conteudo=None):
    """De-para de testes EP → Infinity: o que veio no pacote de EP (conteudo) ou o de reference_data."""
    colunas = ["Provedor", "Módulo/Programa", "Teste Provedor", "Unidade Provedor", "Mneumonico Infinity",
               "Fator", "Confiança", "Observação"]
    if conteudo is None and not ARQ_DEPARA_EP.exists():
        return pd.DataFrame(columns=colunas)
    d = pd.read_excel(io.BytesIO(conteudo) if conteudo is not None else ARQ_DEPARA_EP, dtype=str, **SO_VAZIO_E_NA)
    for c in colunas:
        if c not in d.columns:
            d[c] = ""
    for c in ["Provedor", "Módulo/Programa", "Teste Provedor", "Unidade Provedor", "Confiança", "Observação"]:
        d[c] = d[c].fillna("").astype(str).str.strip()
    d["Mneumonico Infinity"] = d["Mneumonico Infinity"].fillna("").astype(str).str.strip().str.upper()
    d["Fator"] = pd.to_numeric(d["Fator"], errors="coerce").fillna(1.0)
    return d[colunas]


@st.cache_data(show_spinner=False)
def carrega_base_ep(assinatura):
    return ep_base.carrega_base()


def mes_ano_curto(d):
    return f"{MESES_NOME[d.month]}/{str(d.year)[2:]}" if pd.notna(d) and d else ""


def mes_ano_ep(d):
    return f"{MESES_NOME[d.month]}/{d.year}" if pd.notna(d) and d else ""


def cor_limite_ep(v, limite):
    """Só verde (dentro do limite) ou vermelho (fora)."""
    if v is None or pd.isna(v):
        return ""
    return "background-color: #F4CCCC" if abs(v) > limite else "background-color: #D9EAD3"


CORES_AMOSTRA_EP = ["#1565C0", "#2E7D32", "#EC407A", "#F57F17", "#6A1B9A", "#00838F", "#5D4037"]


def grafico_historico_ep(dados, coluna, titulo, faixa=None, limite_vermelho=None, eixo_minimo=4.0,
                         rotulo_faixa=None):
    """Um ponto por amostra (cor = posição da amostra na rodada, como no gráfico do ControlLab), faixa
    aceitável em verde e a média de cada rodada ligada por uma linha cinza."""
    validos = dados.dropna(subset=[coluna])
    # posição no eixo = mês da rodada (CAP e ControlLab do mesmo mês ficam na mesma coluna)
    validos = validos.assign(_mes=validos["_data"].map(lambda d: pd.Timestamp(d.year, d.month, 1)))
    fig = go.Figure()
    if faixa is not None and pd.notna(faixa):
        fig.add_hrect(y0=-faixa, y1=faixa, fillcolor="#2E7D32", opacity=0.13, line_width=0)
        if rotulo_faixa:   # rótulo só na linha de cima (sem texto, o plotly escreve "new text")
            fig.add_hline(y=faixa, line=dict(color="#C62828", width=1.2, dash="dash"),
                          annotation_text=rotulo_faixa, annotation_position="top left")
            fig.add_hline(y=-faixa, line=dict(color="#C62828", width=1.2, dash="dash"))
    if limite_vermelho is not None:
        for s in (1, -1):
            fig.add_hline(y=s * limite_vermelho, line=dict(color="#C62828", width=1, dash="dash"))
    fig.add_hline(y=0, line=dict(color="#999999", width=1))
    for prov, gp in validos.groupby("Provedor"):
        medias = gp.groupby(["_mes", "Programa", "Rodada"])[coluna].mean().reset_index().sort_values("_mes")
        fig.add_trace(go.Scatter(
            x=medias["_mes"], y=medias[coluna], mode="lines", hoverinfo="skip", name=f"Média da rodada ({prov})",
            line=dict(color="#757575", width=1.5, dash="solid" if prov == "ControlLab" else "dot")))
    for pos, gp in validos.groupby("Posição"):
        fig.add_trace(go.Scatter(
            x=gp["_mes"], y=gp[coluna], mode="markers", name=f"Amostra {int(pos)}",
            marker=dict(size=10, color=CORES_AMOSTRA_EP[(int(pos) - 1) % len(CORES_AMOSTRA_EP)],
                        symbol=["diamond" if p == "ControlLab" else "circle" for p in gp["Provedor"]],
                        line=dict(width=0.5, color="white")),
            text=[f"{p} · {e} · {mes_ano_ep(d)}" for p, e, d in zip(gp["Provedor"], gp["Especime"], gp["_data"])],
            hovertemplate="%{text}<br>%{y:.2f}<extra></extra>"))
    datas = sorted(set(validos["_mes"]))
    fig.update_xaxes(tickvals=datas, ticktext=[mes_ano_ep(d) for d in datas],
                     tickangle=-45 if len(datas) > 6 else 0)
    if eixo_minimo is not None:
        maior = max([eixo_minimo] + [abs(v) * 1.1 for v in validos[coluna]])
        fig.update_yaxes(range=[-maior, maior])
    if validos.empty:
        fig.add_annotation(text="sem dados", showarrow=False, x=0.5, y=0.5, xref="paper", yref="paper")
    fig.update_layout(title=titulo, height=380, margin=dict(t=40, b=10), legend=dict(orientation="h", y=-0.3))
    return fig


with tab_ep:
    st.subheader("Ensaio de Proficiência (EP) — CAP e ControlLab")
    st.caption(
        "Cada amostra do EP é comparada com o CIQ do equipamento escolhido, no mês da rodada. "
        "O período da barra lateral também filtra as rodadas daqui."
        + ("" if completo else " No modo simples aparece só o Bias do EP, sem Sigma.")
    )
    avisos_ep = st.container()   # avisos gerais do EP, acima das abas internas
    ep_tab_res, ep_tab_hist, ep_tab_base = st.tabs(
        ["📋 Rodadas e resultados", "📈 Histórico do teste", "⚙️ Atualizar base"])

    # ---- 1. Atualizar a base ----
    pdftotext_exe = ep_base.localiza_pdftotext()

    # ---- Importação automática das pastas configuradas (uma vez por sessão do navegador) ----
    # No app online (Streamlit Cloud, servidor Linux fora da rede do Sabin) não há acesso às pastas
    # da rede: a atualização por pasta só roda com o app aberto num computador Windows do Sabin.
    app_online = os.name != "nt"
    if app_online and pacote_ep is None:
        with avisos_ep:
            st.info("📦 Para ver a base de EP, suba o **pacote_ep.zip** no campo **\"Base de EP (.zip)\"** da "
                    "barra lateral (à esquerda, logo abaixo do zip do CIQ).")
    config_ep = ep_base.carrega_config()
    if (not app_online and config_ep["automatico"] and config_ep["pastas"]
            and not st.session_state.get("ep_sync_feito")):
        with st.spinner("Procurando arquivos novos de EP nas pastas configuradas..."):
            st.session_state["ep_sync_resumo"] = ep_base.sincroniza_pastas(
                config_ep["pastas"], pdftotext=pdftotext_exe or "pdftotext")
        st.session_state["ep_sync_feito"] = True
    with ep_tab_base:
        resumo_sync = st.session_state.get("ep_sync_resumo")
        if resumo_sync:
            if resumo_sync["novos"]:
                st.success(f"Atualização automática: {resumo_sync['novos']} arquivo(s) novo(s), "
                           f"{resumo_sync['linhas']} resultado(s) importado(s) das pastas configuradas.")
            if resumo_sync["inexistentes"] and len(resumo_sync["inexistentes"]) == len(config_ep["pastas"]):
                # computador sem acesso às pastas: usa a base já atualizada (sincronizada pelo OneDrive)
                ultima = ""
                if ep_base.ARQ_LOG.exists():
                    linhas_log = [l for l in ep_base.ARQ_LOG.read_text(encoding="utf-8").splitlines() if l[:1].isdigit()]
                    ultima = f" Última atualização da base: {linhas_log[-1][:16].replace('T', ' às ')}." if linhas_log else ""
                st.info("As pastas de EP configuradas não estão acessíveis deste computador — os dados mostrados "
                        f"são os da última atualização feita num computador com acesso.{ultima}")
            else:
                for p in resumo_sync["inexistentes"]:
                    st.warning(f"Pasta de EP não encontrada: {p}")
            if resumo_sync["erros"]:
                with st.expander(f"⚠ {len(resumo_sync['erros'])} arquivo(s) de EP não puderam ser lidos"):
                    for e in resumo_sync["erros"]:
                        st.write(f"- {e}")

        with st.expander("📁 Atualização automática por pasta"):
            if app_online:
                st.info(
                    "Esta é a versão online do app: ela roda num servidor fora do Sabin e não enxerga as pastas da "
                    "rede (CAP e ControlLab). A base de EP mostrada aqui é a que foi publicada junto com o app. "
                    "A atualização automática por pasta funciona com o app aberto num computador do Sabin que "
                    "tenha acesso às pastas; aqui dá para subir CSVs do ControlLab manualmente (abaixo), mas o "
                    "que for importado online não fica guardado depois que o app reinicia."
                )
            else:
                st.caption(
                    "Informe as pastas onde ficam os arquivos do CAP (PDFs \"Original Evaluation\") e do ControlLab "
                    "(CSVs ou .zip) — uma por linha; subpastas entram também. Com a importação automática ligada, "
                    "toda vez que o app é aberto ele procura arquivos novos ou alterados nessas pastas e grava na "
                    "base sozinho (o que já foi lido não é lido de novo). Para atualizar sem ninguém abrir o app, "
                    "o mesmo pode ser agendado no Windows com o comando `python ep_base.py`."
                )
                pastas_txt = st.text_area("Pastas de entrada", value="\n".join(config_ep["pastas"]),
                                          key="ep_pastas_cfg",
                                          placeholder=r"D:\OneDrive - ...\EP\CAP" + "\n" + r"D:\OneDrive - ...\EP\ControlLab")
                auto_ep = st.checkbox("Importar automaticamente ao abrir o app", value=config_ep["automatico"],
                                      key="ep_auto_cfg")
                ca, cb = st.columns(2)
                pastas_novas = [p.strip().strip('"') for p in pastas_txt.splitlines() if p.strip()]
                if ca.button("Salvar configuração", key="ep_salvar_cfg"):
                    ep_base.grava_config(pastas_novas, auto_ep)
                    faltando = [p for p in pastas_novas if not Path(p).exists()]
                    if faltando:
                        st.warning("Configuração salva, mas estas pastas não foram encontradas: " + "; ".join(faltando))
                    else:
                        st.success("Configuração salva.")
                if cb.button("Verificar as pastas agora", key="ep_sync_agora", disabled=not pastas_novas):
                    with st.spinner("Procurando arquivos novos..."):
                        st.session_state["ep_sync_resumo"] = ep_base.sincroniza_pastas(
                            pastas_novas, pdftotext=pdftotext_exe or "pdftotext")
                    r = st.session_state["ep_sync_resumo"]
                    st.info(f"{r['verificados']} arquivo(s) verificado(s), {r['novos']} novo(s), "
                            f"{r['ignorados']} repetido(s), {r['linhas']} resultado(s) importado(s).")
                if ep_base.ARQ_LOG.exists():
                    st.caption("Últimas atualizações (dados_ep/ep_log.txt)")
                    st.code("".join(ep_base.ARQ_LOG.read_text(encoding="utf-8").splitlines(True)[-15:]))

        with st.expander("⬆ Atualizar base de EP manualmente (subir arquivos)"):
            st.caption(
                "PDFs \"Original Evaluation\" do CAP, CSVs de avaliação do ControlLab ou um .zip com eles. "
                "Arquivos já processados são reconhecidos e ignorados (não duplicam). Para uma carga grande, "
                "informe a pasta em vez de subir os arquivos. No ControlLab entram só os resultados do(s) "
                f"laboratório(s) {', '.join(sorted(ep_base.PARTICIPANTES_CONTROLLAB))}."
            )
            if not pdftotext_exe:
                st.warning("Leitor de PDF do CAP indisponível aqui (precisa do pdftotext do Git for Windows, que não "
                           "existe no app online) — os PDFs do CAP não serão lidos; os CSVs do ControlLab funcionam.")
            arquivos_ep = st.file_uploader("Arquivos", type=["pdf", "csv", "zip"], accept_multiple_files=True,
                                           key="ep_upload")
            pasta_ep = st.text_input("…ou uma pasta (ou .zip) do computador", key="ep_pasta",
                                     placeholder=r"D:\Ligia\EP_entrada")
            if st.button("Ler arquivos", key="ep_ler"):
                entradas = []
                for f in (arquivos_ep or []):
                    conteudo = f.getvalue()
                    if f.name.lower().endswith(".zip"):
                        try:
                            eh_pacote = any(Path(n).name == "ep_resultados.csv"
                                            for n in zipfile.ZipFile(io.BytesIO(conteudo)).namelist())
                        except zipfile.BadZipFile:
                            eh_pacote = False
                        if eh_pacote:
                            st.warning(f"\"{f.name}\" é um pacote de EP — suba-o no campo \"Base de EP (.zip)\" da "
                                       "barra lateral (à esquerda), não aqui.")
                            continue
                    entradas.append((f.name, conteudo))
                if pasta_ep.strip():
                    if Path(pasta_ep.strip()).exists():
                        entradas = [*entradas, *ep_base.entradas_de_caminhos([pasta_ep.strip()])]
                    else:
                        st.error(f"Pasta não encontrada: {pasta_ep}")
                hashes = set(ep_base.carrega_registro()["Hash"])
                barra = st.progress(0.0, text="Lendo arquivos...")
                linhas_ep, registro_ep, ignorados_ep = ep_base.processa_entradas(
                    entradas, hashes, pdftotext=pdftotext_exe or "pdftotext",
                    progresso=lambda n: barra.progress(min(1.0, n / max(len(entradas), 1)),
                                                       text=f"{n} arquivo(s) lido(s)..."))
                barra.empty()
                st.session_state["ep_pendente"] = (linhas_ep, registro_ep, ignorados_ep)

            pendente = st.session_state.get("ep_pendente")
            if pendente:
                linhas_ep, registro_ep, ignorados_ep = pendente
                st.info(f"{len(registro_ep)} arquivo(s) novo(s), {ignorados_ep} já processado(s) antes · "
                        f"{len(linhas_ep)} resultado(s) quantitativo(s) lido(s).")
                if linhas_ep:
                    prev = pd.DataFrame(linhas_ep)
                    st.dataframe(prev.groupby(["Provedor", "Programa"]).agg(
                        Rodadas=("Rodada", "nunique"), Resultados=("RL", "size")).reset_index(),
                        hide_index=True, use_container_width=True)
                    if st.checkbox("Mostrar prévia das linhas lidas", key="ep_ver_previa"):
                        st.dataframe(prev.head(500), hide_index=True, use_container_width=True)
                c_g, c_d = st.columns(2)
                if c_g.button("Gravar na base de EP", type="primary", key="ep_gravar"):
                    ep_base.grava(linhas_ep, registro_ep)
                    del st.session_state["ep_pendente"]
                    st.success("Base de EP atualizada.")
                    st.rerun()
                if c_d.button("Descartar", key="ep_descartar"):
                    del st.session_state["ep_pendente"]
                    st.rerun()

    # base de EP: pacote enviado na barra lateral + o que houver na base local (sem duplicar)
    rodadas_pacote = depara_pacote = None
    base_ep = carrega_base_ep(_mtime(ep_base.ARQ_RESULTADOS))
    if pacote_ep is not None:
        try:
            base_pacote, rodadas_pacote, depara_pacote = ep_base.le_pacote(pacote_ep.getvalue())
            base_ep = (ep_base.consolida(base_pacote, base_ep.to_dict("records")) if len(base_ep)
                       else base_pacote)
            with ep_tab_base:
                st.caption(f"📦 Base de EP do pacote enviado ({pacote_ep.name}): {len(base_pacote)} resultado(s)"
                           + (f"; {len(base_ep)} ao juntar com a base local." if len(base_ep) != len(base_pacote) else "."))
        except (ValueError, zipfile.BadZipFile) as e:
            with avisos_ep:
                st.error(f"Não foi possível usar o pacote de EP: {e}")
    # de-para: o que veio no pacote tem prioridade (o app online não depende do arquivo no GitHub)
    depara_ep = carrega_depara_ep(_mtime(ARQ_DEPARA_EP), depara_pacote)
    if depara_ep.empty and not base_ep.empty:
        with avisos_ep:
            st.warning("De-para de testes não encontrado (nem no pacote de EP, nem em "
                       "reference_data/ep_depara_testes.xlsx) — sem ele nenhum teste do EP é correlacionado com o "
                       "Infinity. Gere o pacote de novo num computador que tenha o de-para (ele vai junto).")
    if base_ep.empty:
        with avisos_ep:
            st.info("A base de EP ainda está vazia — suba o pacote de EP na barra lateral ou use "
                    "a aba ⚙️ Atualizar base.")
    else:
        b = base_ep.copy()
        for c in ("Lim Inf", "Lim Sup"):   # pacotes gerados antes de existirem essas colunas
            if c not in b.columns:
                b[c] = None
        b["Módulo/Programa"] = b["Programa"].where(b["Provedor"] != "CAP",
                                                   b["Programa"].str.replace(r"-[A-Z]$", "", regex=True))
        b = b.merge(depara_ep, how="left",
                    left_on=["Provedor", "Módulo/Programa", "Teste Provedor", "Unidade"],
                    right_on=["Provedor", "Módulo/Programa", "Teste Provedor", "Unidade Provedor"])
        # o de-para pode trazer mais de um mnemônico ("ACTH;ACT", quando o teste mudou de equipamento e de
        # código): o 1º dá nome ao teste; o CIQ usado é o do primeiro que existir no mês da rodada
        b["Mneumonicos"] = b["Mneumonico Infinity"].fillna("").astype(str).map(
            lambda s: tuple(m.strip() for m in s.split(";") if m.strip()))
        b["Mneumonico"] = b["Mneumonicos"].map(lambda t: t[0] if t else "")
        b["Fator"] = b["Fator"].fillna(1.0)
        b["Mês/Ano"] = b["Data Envio"].map(mes_ano_curto)
        b = b[b["Mês/Ano"].isin(periodo_valido_g)]

        if MOSTRAR_DEPARA_EP:
            with ep_tab_base:
                # ---- Situação do de-para ----
                sem_depara = (b[b["Confiança"].isna()]
                              .groupby(["Provedor", "Módulo/Programa", "Teste Provedor", "Unidade"])
                              .size().reset_index(name="Resultados"))
                duvidas = depara_ep[depara_ep["Confiança"] == "dúvida"]
                m1, m2, m3 = st.columns(3)
                m1.metric("Rodadas com teste correlacionado", b[b["Mneumonico"] != ""].groupby(
                    ["Provedor", "Programa", "Rodada", "Mneumonico", "Sistema"]).ngroups)
                m2.metric("Testes sem de-para no período", len(sem_depara))
                m3.metric("Dúvidas no de-para", len(duvidas))
                if len(sem_depara) or len(duvidas):
                    with st.expander("Ver testes sem de-para e dúvidas"):
                        st.caption(
                            "O de-para fica em `reference_data/ep_depara_testes.xlsx` (colunas Mneumonico Infinity e "
                            "Fator). Edite, salve e recarregue o app. Teste sem mnemônico não entra no cálculo."
                        )
                        if len(sem_depara):
                            st.markdown("**Testes do EP que ainda não estão no de-para**")
                            st.dataframe(sem_depara, hide_index=True, use_container_width=True)
                        if len(duvidas):
                            st.markdown("**Correlações marcadas como dúvida**")
                            st.dataframe(duvidas[["Provedor", "Módulo/Programa", "Teste Provedor",
                                                  "Mneumonico Infinity", "Fator", "Observação"]],
                                         hide_index=True, use_container_width=True)

        escolhas_atuais = None
        resumo_total = pd.DataFrame()
        with ep_tab_res:
            mapeado = b[b["Mneumonico"] != ""].copy()
            if mapeado.empty:
                st.info("Nenhuma rodada com teste correlacionado no período selecionado.")
            else:
                ciq = df_ciq_total.copy()
                ciq["TesteU"] = ciq["Teste"].str.upper()
                ciq["N"] = ciq["N"].fillna(0)
                ciq_por_teste = {t: g for t, g in ciq.groupby("TesteU")}
                ciq_valido = ciq.dropna(subset=["CV (%)", "Média", "NívelNum"])
                ciq_nivel = ciq_valido.loc[ciq_valido.groupby(
                    ["TesteU", "Equipamento (nome)", "Ano", "Mês", "NívelNum"])["N"].idxmax()]
                ciq_nivel = {k: g for k, g in ciq_nivel.groupby(["TesteU", "Equipamento (nome)", "Ano", "Mês"])}
                # especificação (ETM/ESM) por teste, pra análise de tendência mesmo sem CIQ no mês da rodada
                spec_por_teste = {t: next((s for s in g["Spec"] if isinstance(s, dict)), None)
                                  for t, g in ciq.groupby("TesteU")}

                def plataforma_do_provedor(nome):
                    n = str(nome or "").lower()
                    return next((tipo for chave, tipo in PLATAFORMA_EP if chave in n), None)

                def sugere_equipamento(mneus, d_envio, equip_prov):
                    partes = [ciq_por_teste[m] for m in mneus if m in ciq_por_teste]
                    if not partes:
                        return ""
                    cand = pd.concat(partes)
                    tipo = plataforma_do_provedor(equip_prov)
                    if tipo:
                        mesma_plat = cand[cand["Tipo"].astype(str).str.upper() == tipo]
                        cand = mesma_plat if not mesma_plat.empty else cand
                    no_mes = cand[(cand["Ano"] == d_envio.year) & (cand["Mês"] == d_envio.month)]
                    cand = no_mes if not no_mes.empty else cand
                    return cand.groupby("Equipamento (nome)")["N"].sum().idxmax()

                def indice_desvio(a):
                    """ID: o ControlLab informa; no CAP sai dos limites de aceitação (|ID| > 1 = fora da faixa)."""
                    if a["Provedor"] == "ControlLab":
                        return a["SDI"] if pd.notna(a["SDI"]) else None
                    rl, vd, li, ls = a["RL"], a["VD"], a.get("Lim Inf"), a.get("Lim Sup")
                    if any(v is None or pd.isna(v) for v in (rl, vd, li, ls)):
                        return None
                    meia_faixa = (ls - vd) if rl >= vd else (vd - li)
                    return (rl - vd) / meia_faixa if meia_faixa > 0 else None

                rodadas = (mapeado.groupby(CHAVE_RODADA).agg(
                    Data=("Data Envio", "min"),
                    Mneus=("Mneumonicos", "first"),
                    Teste_prov=("Teste Provedor", lambda s: "; ".join(sorted(set(s)))),
                    Equip_prov=("Equipamento Provedor", lambda s: "; ".join(sorted({x for x in s if x}))),
                    Amostras=("Especime", "nunique")).reset_index())
                salvas = ep_base.carrega_rodadas()
                if rodadas_pacote is not None:
                    # escolhas que vieram no pacote; as salvas nesta sessão (arquivo local) valem por cima
                    salvas = (pd.concat([rodadas_pacote, salvas], ignore_index=True)
                              .drop_duplicates(CHAVE_RODADA, keep="last"))
                rodadas = rodadas.merge(salvas[CHAVE_RODADA + ["Equipamento"]], how="left", on=CHAVE_RODADA)
                rodadas["Equipamento"] = rodadas["Equipamento"].fillna("")
                rodadas["Origem"] = rodadas["Equipamento"].map(lambda e: "salvo" if e else "")
                faltam = rodadas["Equipamento"] == ""
                rodadas.loc[faltam, "Equipamento"] = [
                    sugere_equipamento(r.Mneus, r.Data, r.Equip_prov) for r in rodadas[faltam].itertuples()]
                rodadas.loc[faltam & (rodadas["Equipamento"] != ""), "Origem"] = "sugerido"
                rodadas["Mês"] = pd.to_datetime(rodadas["Data"])
                rodadas = rodadas.sort_values(["Data", "Provedor", "Programa", "Mneumonico"],
                                              ascending=[False, True, True, True])

                # ---- Filtros (só mudam o que é exibido; o cálculo é feito para todas as rodadas) ----
                fe1, fe2, fe3 = st.columns(3)
                f_prov_ep = fe1.multiselect("Provedor", sorted(rodadas["Provedor"].unique()), key="ep_f_prov")
                f_teste_ep = fe2.multiselect("Teste (Infinity)", sorted(rodadas["Mneumonico"].unique()),
                                             key="ep_f_teste")
                f_equip_ep = fe3.multiselect("Equipamento", sorted(e for e in rodadas["Equipamento"].unique() if e),
                                             key="ep_f_equip")

                # ---- Como o Sigma do EP é calculado (vale para todas as rodadas) ----
                if completo:
                    oc1, oc2, oc3 = st.columns(3)
                    vies_ep = oc1.selectbox(
                        "Viés usado no Sigma", list(OPCOES_VIES_EP), key="ep_vies",
                        help="Médio (%): um viés só para a rodada. Regressão: o viés muda com a concentração. "
                             "Automático: o app escolhe por rodada (veja \"O que escolher\").")
                    formula_ep = oc2.selectbox(
                        "Fórmula do Sigma", list(OPCOES_FORMULA_EP), key="ep_formula",
                        help="σ %: viés e CV em %. σ absoluto: viés e DP na unidade do exame. Automático: σ absoluto "
                             "quando há amostra abaixo do cutoff do ETM absoluto.")
                    pareamento_ep = oc3.selectbox(
                        "Amostra × nível do CIQ", ep_calculo.PAREAMENTOS, key="ep_pareamento",
                        help="Por concentração: nível do CIQ com média mais próxima do valor designado. "
                             "Por posição: amostra 1 → nível 1, amostra 2 → nível 2 (como a planilha).")
                    with st.expander("ℹ O que escolher"):
                        st.markdown(
                            "**Viés usado no Sigma**\n"
                            "- **Médio (%)**: (média dos resultados do laboratório − média dos valores designados) ÷ "
                            "média dos valores designados × 100. Um viés só, igual para todos os níveis do CIQ. Use "
                            "quando as amostras da rodada têm concentrações parecidas.\n"
                            "- **Regressão**: reta resultado × valor designado com as amostras da rodada; o viés é "
                            "lido na concentração de cada nível do CIQ. Use quando as concentrações são bem "
                            "diferentes, porque o viés pode mudar com a concentração. Precisa de 3 amostras ou mais "
                            f"e boa correlação (r ≥ {ep_calculo.R_MINIMO_REGRESSAO:g}).\n"
                            f"- **Automático** (recomendado): regressão quando a maior concentração da rodada é "
                            f"≥ {ep_calculo.RAZAO_CONCENTRACAO_REGRESSAO:g}× a menor, há 3 amostras ou mais e "
                            f"r ≥ {ep_calculo.R_MINIMO_REGRESSAO:g}; senão, médio.\n\n"
                            "**Fórmula do Sigma**\n"
                            "- **σ %** = (ETM % − |viés %|) ÷ CV % do CIQ. Serve para a maioria dos testes.\n"
                            "- **σ absoluto** = (ETM % × média do CIQ − |viés na unidade|) ÷ DP do CIQ. Use quando o "
                            "teste tem ETM absoluto (concentrações baixas, abaixo do cutoff).\n"
                            "- **Automático** (recomendado): σ absoluto quando alguma amostra da rodada está no cutoff "
                            "do ETM absoluto ou abaixo; senão σ %.\n\n"
                            "Com viés por regressão as duas fórmulas dão o mesmo resultado. Quando a média do CIQ do "
                            "nível está no cutoff ou abaixo, vale sempre o ETM absoluto.\n\n"
                            "**Amostra × nível do CIQ** (qual CV entra no Sigma)\n"
                            "- **Por concentração** (recomendado): cada amostra usa o nível do CIQ com média mais "
                            "próxima do seu valor designado — o CV é o da mesma faixa em que o viés foi medido.\n"
                            "- **Por posição** (como a planilha): amostra 1 → nível 1, amostra 2 → nível 2..."
                        )
                else:
                    vies_ep, formula_ep, pareamento_ep = "Automático", "Automático", ep_calculo.PAREAMENTOS[0]

                vis = rodadas
                if f_prov_ep:
                    vis = vis[vis["Provedor"].isin(f_prov_ep)]
                if f_teste_ep:
                    vis = vis[vis["Mneumonico"].isin(f_teste_ep)]
                if f_equip_ep:
                    vis = vis[vis["Equipamento"].isin(f_equip_ep)]

                st.markdown("**Rodadas e equipamento**")
                st.caption(
                    "Os relatórios não dizem em qual equipamento a amostra foi analisada: confira o equipamento de "
                    "cada rodada (\"sugerido\" = escolhido pelo app, pelo equipamento informado pelo provedor ou pelo "
                    "de maior volume de CIQ no mês) e clique em Salvar."
                )
                opcoes_equip = [""] + sorted(df_ciq_total["Equipamento (nome)"].dropna().astype(str).unique())
                colunas_editor = ["Mês", "Provedor", "Programa", "Mneumonico", "Teste_prov", "Sistema", "Equip_prov",
                                  "Amostras", "Equipamento", "Origem"]
                vis = vis.reset_index(drop=True)
                # a chave muda com os filtros: a edição não "escorrega" para outra linha quando a lista muda
                chave_editor = "ep_editor|" + "|".join(["/".join(f_prov_ep), "/".join(f_teste_ep), "/".join(f_equip_ep)])
                editado = st.data_editor(
                    vis[colunas_editor], key=chave_editor, hide_index=True, use_container_width=True, height=300,
                    disabled=[c for c in colunas_editor if c != "Equipamento"],
                    column_config={
                        "Mês": st.column_config.DateColumn("Mês", format="MM/YYYY"),
                        "Mneumonico": st.column_config.TextColumn("Teste (Infinity)"),
                        "Teste_prov": st.column_config.TextColumn("Teste no provedor"),
                        "Equip_prov": st.column_config.TextColumn("Equip. informado"),
                        "Equipamento": st.column_config.SelectboxColumn("Equipamento (S.A.)", options=opcoes_equip),
                    },
                )
                vis["Equipamento"] = editado["Equipamento"].to_numpy()
                escolhas_atuais = (pd.concat([salvas, vis[vis["Equipamento"] != ""][CHAVE_RODADA + ["Equipamento"]]],
                                             ignore_index=True).drop_duplicates(CHAVE_RODADA, keep="last")
                                   .reindex(columns=ep_base.COLUNAS_RODADAS).fillna(""))
                if st.button("Salvar equipamentos", key="ep_salvar_rodadas"):
                    ep_base.grava_rodadas(escolhas_atuais)
                    st.success(f"{len(escolhas_atuais)} escolha(s) gravada(s)."
                               + (" No app online elas valem só nesta sessão: para guardá-las, baixe o pacote de EP "
                                  "na aba ⚙️ Atualizar base." if app_online else ""))

                # ---- Resultados ----
                grupos_ep = {k: g for k, g in mapeado.groupby(CHAVE_RODADA)}
                linhas_res, resumo_res = [], []
                # calcula todas as rodadas (com as edições da tabela) — os filtros só mudam o que é exibido
                todas_rodadas = rodadas.set_index(CHAVE_RODADA)
                todas_rodadas.update(vis.set_index(CHAVE_RODADA)[["Equipamento"]])
                for rd in todas_rodadas.reset_index().to_dict("records"):
                    g = grupos_ep.get(tuple(rd[c] for c in CHAVE_RODADA))
                    if g is None:
                        continue
                    # CAP pode trazer o mesmo espécime em dois grupos de comparação: fica o que foi avaliado
                    g = (g.assign(_aval=g["Nota"].isin(["Acceptable", "Unacceptable"]).astype(int))
                         .sort_values("_aval", ascending=False).drop_duplicates("Especime"))
                    amostras = [{"Especime": a["Especime"], "Num": a["Num"], "RL": a["RL"] * a["Fator"],
                                 "VD": a["VD"] * a["Fator"], "Qualificador": a["Qualificador"],
                                 "DP Grupo": a["DP Grupo"] * a["Fator"] if pd.notna(a["DP Grupo"]) else None,
                                 "Índice Provedor": indice_desvio(a)}
                                for a in g.to_dict("records")]
                    d_envio = g["Data Envio"].iloc[0]
                    niveis = next((ciq_nivel[k] for k in ((m, rd["Equipamento"], d_envio.year, d_envio.month)
                                                          for m in rd["Mneus"]) if k in ciq_nivel), None)
                    ciq_rodada, spec = {}, None
                    if niveis is not None:
                        for _, n in niveis.iterrows():
                            ciq_rodada[int(n["NívelNum"])] = {"Média": n["Média"], "CV (%)": n["CV (%)"]}
                        spec = next((s for s in niveis["Spec"] if isinstance(s, dict)), None)
                    linhas, info = ep_calculo.calcula_rodada(amostras, ciq_rodada, spec, OPCOES_VIES_EP[vies_ep],
                                                             OPCOES_FORMULA_EP[formula_ep], pareamento_ep)
                    spec_teste = spec or next((spec_por_teste[m] for m in rd["Mneus"] if spec_por_teste.get(m)), None)
                    izs = [abs(l["IZ"]) for l in linhas if l["IZ"] is not None]
                    base_linha = {"Mês": pd.Timestamp(d_envio), "_data": d_envio, "Provedor": rd["Provedor"],
                                  "Programa": rd["Programa"], "Rodada": rd["Rodada"], "Teste": rd["Mneumonico"],
                                  "Sistema": rd["Sistema"], "Equipamento": rd["Equipamento"] or "—"}
                    for l in linhas:
                        linhas_res.append({**base_linha, **l})
                    sigmas = [l["Sigma EP"] for l in linhas if l["Sigma EP"] is not None]
                    resumo_res.append({**base_linha, "Viés médio (%)": info["Viés médio %"],
                                       "Pior |IZ|": max(izs) if izs else None,
                                       "Pior Sigma EP": min(sigmas) if sigmas else None,
                                       "ESM (%)": (spec_teste or {}).get("ESM (%)")})

                res_ep_global = pd.DataFrame(linhas_res)
                resumo_total = pd.DataFrame(resumo_res)
                res = res_ep_global
                if f_prov_ep and not res.empty:
                    res = res[res["Provedor"].isin(f_prov_ep)]
                if f_teste_ep and not res.empty:
                    res = res[res["Teste"].isin(f_teste_ep)]
                if f_equip_ep and not res.empty:
                    res = res[res["Equipamento"].isin(f_equip_ep)]

                st.divider()
                if completo and not res.empty:
                    ep_t = res[(res["Teste"] == str(teste_global).upper()) & res["Sigma EP"].notna()]
                    if not ep_t.empty:
                        pior = ep_t.loc[ep_t["Sigma EP"].idxmin()]
                        with st.container(border=True):
                            st.caption(f"Pior Sigma do EP no período — {teste_global} (separado do pior cenário do CIQ)")
                            st.markdown(f"<span style='font-size:30px; font-weight:700;'>{pior['Sigma EP']:.2f}</span>",
                                        unsafe_allow_html=True)
                            st.caption(f"{pior['Provedor']} · {pior['Programa']} · {mes_ano_ep(pior['_data'])} · "
                                       f"{pior['Equipamento']} · amostra {pior['Especime']}")

                st.markdown("**Resultados por amostra**")
                if res.empty:
                    st.info("Nenhum resultado para os filtros escolhidos.")
                else:
                    cols_res = ["Mês", "Provedor", "Teste", "Equipamento", "Especime", "RL", "VD", "Bias % amostra",
                                "Índice Provedor", "IZ"]
                    if completo:
                        cols_res += ["Sigma EP", "CV CIQ (%)", "Viés %"]
                    cols_res += ["Programa", "Rodada"]
                    fmt = {c: "{:.2f}" for c in ["Bias % amostra", "Índice Provedor", "IZ", "CV CIQ (%)", "Viés %",
                                                 "Sigma EP"] if c in cols_res}
                    estilo = (res[cols_res].style.format(fmt, na_rep="—")
                              .map(lambda v: cor_limite_ep(v, ep_calculo.LIMITE_ID), subset=["Índice Provedor"])
                              .map(lambda v: cor_limite_ep(v, ep_calculo.LIMITE_IZ_ALERTA), subset=["IZ"]))
                    st.dataframe(estilo, hide_index=True, use_container_width=True, height=420, column_config={
                        "Mês": st.column_config.DateColumn("Mês", format="MM/YYYY"),
                        "Especime": "Amostra", "Bias % amostra": "Bias %", "Índice Provedor": "ID",
                        "IZ": "Z grupo", "Viés %": "Viés % (Sigma)"})
                    st.caption(
                        "**Bias %** = (resultado − valor designado) ÷ valor designado, da própria amostra. "
                        "**ID** = índice de desvio (ControlLab: o que eles informam; CAP: calculado pelos limites de "
                        "aceitação) — 🟩 |ID| ≤ 1 dentro da faixa aceita · 🟥 |ID| > 1 fora. "
                        "**Z grupo** = (resultado − média do grupo) ÷ DP do grupo — 🟩 |Z| ≤ 2 · 🟥 |Z| > 2."
                        + (" **Viés % (Sigma)** = viés da rodada usado no Sigma (médio ou da regressão, conforme a "
                           "escolha acima)." if completo else ""))
                    tabela_csv = (res[cols_res].assign(Mês=res["_data"].map(mes_ano_ep))
                                  .rename(columns={"Especime": "Amostra", "Bias % amostra": "Bias %",
                                                   "Índice Provedor": "ID", "IZ": "Z grupo",
                                                   "Viés %": "Viés % (Sigma)"}))
                    st.download_button("⬇ Baixar esta tabela (CSV, abre no Excel)",
                                       tabela_csv.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                                       "ep_resultados.csv", "text/csv")

        # ---- Histórico do teste: tendência entre rodadas (sem os filtros da aba de resultados) ----
        with ep_tab_hist:
            if resumo_total.empty:
                st.info("Nenhuma rodada com teste correlacionado no período selecionado.")
            else:
                h1, h2 = st.columns([2, 1])
                testes_hist = sorted(resumo_total["Teste"].unique())
                padrao_hist = str(teste_global).upper()
                teste_hist = h1.selectbox("Teste", testes_hist, key="ep_hist_teste",
                                          index=testes_hist.index(padrao_hist) if padrao_hist in testes_hist else 0)
                prov_hist = h2.radio("Provedor", ["Todos", "ControlLab", "CAP"], horizontal=True, key="ep_hist_prov")
                hist = resumo_total[resumo_total["Teste"] == teste_hist]
                am_hist = res_ep_global[res_ep_global["Teste"] == teste_hist]
                if prov_hist != "Todos":
                    hist, am_hist = hist[hist["Provedor"] == prov_hist], am_hist[am_hist["Provedor"] == prov_hist]
                hist = hist.sort_values("_data")
                if hist.empty:
                    st.info(f"Sem rodadas do {prov_hist} para {teste_hist} no período.")
                else:
                    alertas = ep_calculo.analisa_historico(
                        [{"Rótulo": f"{r['Provedor']} {mes_ano_ep(r['_data'])}", "Viés %": r["Viés médio (%)"],
                          "Pior |IZ|": r["Pior |IZ|"]} for r in hist.to_dict("records")])
                    for alerta in alertas:
                        grave = any(p in alerta for p in ("persistente", "piora", "investigar"))
                        (st.warning if grave else st.info)(alerta)
                    if not alertas:
                        st.caption("Poucas rodadas para avaliar tendência (são necessárias pelo menos 2).")

                    esm_hist = next((e for e in hist["ESM (%)"] if pd.notna(e)), None)
                    g1, g2, g3 = st.columns(3)
                    g1.plotly_chart(grafico_historico_ep(am_hist, "Índice Provedor", "Índice de Desvio (ID)",
                                                         faixa=ep_calculo.LIMITE_ID), use_container_width=True)
                    g2.plotly_chart(grafico_historico_ep(am_hist, "IZ", "Índice Z (grupo de comparação)",
                                                         faixa=ep_calculo.LIMITE_IZ_ALERTA,
                                                         limite_vermelho=ep_calculo.LIMITE_IZ_ACAO),
                                    use_container_width=True)
                    g3.plotly_chart(grafico_historico_ep(am_hist, "Bias % amostra", "Viés (%) por amostra",
                                                         faixa=esm_hist, eixo_minimo=None, rotulo_faixa="ESM"),
                                    use_container_width=True)
                    st.caption("Cada cor é uma amostra da rodada (1ª, 2ª, 3ª...); ◆ ControlLab · ● CAP. A linha cinza "
                               "liga a média das amostras de cada rodada. Faixa verde: ID ±1, Z ±2 e viés dentro do ESM.")

        # pacote de EP (base + escolhas de equipamento): para o app online, que não guarda nada
        with ep_tab_base:
            st.divider()
            st.download_button(
                "📦 Baixar pacote de EP (.zip)", ep_base.monta_pacote(base_ep, escolhas_atuais, depara_pacote),
                f"pacote_ep_{pd.Timestamp.today():%Y-%m-%d}.zip", "application/zip", key="ep_baixar_pacote",
                help="Base de EP + os equipamentos escolhidos na aba Rodadas e resultados. É o arquivo que se sobe "
                     "na barra lateral do app online (ou para guardar uma cópia de segurança).")

# ---------------- SIGMA POR PERÍODO ----------------
if completo:
    with tab_periodo:
        st.subheader("Sigma — mensal e por período")
        st.caption(
            "Recalculado a partir dos dados já filtrados pela barra lateral (Período e "
            "Módulo/Plataforma) — mudar o período lá em cima muda o que aparece aqui."
        )

        # Critério para definição do pior cenário do Sigma (mesma célula da planilha), por analito:
        # None = pior cenário entre os níveis; número = só aquele nível. A escolha fica guardada
        # por teste durante a sessão; o padrão pode vir da coluna "Nível Sigma" da tabela_mestre.
        def nivel_sigma_padrao(teste):
            for linha in mestre_by_mneu.get(str(teste).strip().upper(), []):
                try:
                    return int(float(linha.get("Nível Sigma")))
                except (TypeError, ValueError):
                    continue
            return None

        criterios_sigma = st.session_state.setdefault("criterio_sigma", {})
        niveis_teste_sigma = sorted(int(n) for n in df.loc[df["Teste"] == teste_global, "NívelNum"].dropna().unique())
        crit_atual = criterios_sigma.get(teste_global, nivel_sigma_padrao(teste_global))

        cs1, cs2 = st.columns([2, 1])
        modo_sigma = cs1.radio(
            f"Critério para definição do pior cenário do Sigma — {teste_global}",
            ["Pior cenário", "Nível específico"], index=0 if crit_atual is None else 1,
            horizontal=True, key=f"crit_sigma_modo_{teste_global}",
            help="Pior cenário = menor Sigma entre todos os níveis. Nível específico = o Sigma passa a "
                 "refletir só o nível escolhido. A escolha vale para este teste (analito).",
        )
        nivel_sigma = None
        if modo_sigma == "Nível específico":
            if niveis_teste_sigma:
                nivel_sigma = cs2.selectbox(
                    "Nível", niveis_teste_sigma,
                    index=niveis_teste_sigma.index(crit_atual) if crit_atual in niveis_teste_sigma else 0,
                    key=f"crit_sigma_nivel_{teste_global}",
                )
            else:
                cs2.info("Nenhum nível identificado para esse teste no filtro atual.")
        criterios_sigma[teste_global] = nivel_sigma

        def aplica_criterio_sigma(df_in):
            """Mantém só o nível escolhido nos testes com critério "Nível específico"."""
            alvo = {t: nivel_sigma_padrao(t) for t in df_in["Teste"].dropna().unique()}
            alvo.update(criterios_sigma)
            alvo = {t: n for t, n in alvo.items() if n is not None}
            if not alvo:
                return df_in
            nivel_alvo = df_in["Teste"].map(alvo)
            return df_in[nivel_alvo.isna() | (df_in["NívelNum"] == nivel_alvo)]

        df_sigma = aplica_criterio_sigma(df)
        rotulo_crit = "pior cenário entre os níveis" if nivel_sigma is None else f"Nível {nivel_sigma}"

        # CIQ + EP: os pontos do EP entram no gráfico com outro marcador e têm card próprio —
        # o pior cenário do CIQ e o do EP não se misturam (decisão com as regionais)
        mostrar_ep = st.toggle(
            "Mostrar também o EP (CIQ + EP)", value=False, key="sigma_com_ep",
            help="Os pontos do EP aparecem como losangos no gráfico e ganham um card de pior cenário "
                 "próprio. O pior cenário do CIQ continua calculado só com o CIQ.")
        ep_teste = pd.DataFrame()
        if mostrar_ep and not res_ep_global.empty:
            ep_teste = res_ep_global[(res_ep_global["Teste"] == str(teste_global).upper())
                                     & res_ep_global["Sigma EP"].notna()].copy()
            if nivel_sigma is not None:
                ep_teste = ep_teste[ep_teste["Nível CIQ"] == nivel_sigma]
            ep_teste["Mês/Ano"] = ep_teste["_data"].map(mes_ano_curto)
            ep_teste["_ordem_tempo"] = ep_teste["_data"].map(lambda d: d.year * 12 + d.month)

        if mostrar_ep:
            cc1, cc2 = st.columns(2)
            with cc1:
                card_pior_cenario(df_sigma[df_sigma["Teste"] == teste_global], "Sigma Mensal", "Sigma",
                                  maior_eh_pior=False)
            with cc2:
                if ep_teste.empty:
                    st.info("Sem Sigma de EP para esse teste no período (confira a aba EP).")
                else:
                    pior_ep = ep_teste.loc[ep_teste["Sigma EP"].idxmin()]
                    with st.container(border=True):
                        c1, c2 = st.columns([3, 1])
                        with c1:
                            st.caption("Pior (mais baixo) Sigma do EP do período")
                            st.markdown(f"<span style='font-size:30px; font-weight:700;'>"
                                        f"{pior_ep['Sigma EP']:.2f}</span>", unsafe_allow_html=True)
                            st.caption(f"{pior_ep['Provedor']} {pior_ep['Especime']} · {pior_ep['Equipamento']} · "
                                       f"Nível {int(pior_ep['Nível CIQ'])}")
                        with c2:
                            st.caption("Mês/Ano")
                            st.markdown(f"**{pior_ep['Mês/Ano']}**")
        else:
            card_pior_cenario(df_sigma[df_sigma["Teste"] == teste_global], "Sigma Mensal", "Sigma",
                              maior_eh_pior=False)

        st.markdown(f"**Pior Sigma por mês — {teste_global} ({rotulo_crit})**")
        base_pior_mes = df_sigma[(df_sigma["Teste"] == teste_global)].dropna(subset=["Sigma Mensal"]).copy()
        if base_pior_mes.empty:
            st.info("Nenhum registro de Sigma para esse teste no filtro atual.")
        else:
            base_pior_mes["N"] = base_pior_mes["N"].fillna(0)
            # se houver mais de um lote no mesmo mês/equipamento/nível, usa o de maior N
            idx_pm = base_pior_mes.groupby(["Equipamento (nome)", "NívelNum", "_ordem_tempo"])["N"].idxmax()
            base_pior_mes = base_pior_mes.loc[idx_pm]
            # pior Sigma entre TODOS os equipamentos/níveis daquele teste, mês a mês
            idx_pior_mes = base_pior_mes.groupby("_ordem_tempo")["Sigma Mensal"].idxmin()
            serie_pior = base_pior_mes.loc[idx_pior_mes].sort_values("_ordem_tempo")

            fig_sigma_mes = go.Figure()
            fig_sigma_mes.add_trace(go.Scatter(
                x=serie_pior["Mês/Ano"], y=serie_pior["Sigma Mensal"],
                mode="lines+markers", name="Pior Sigma do mês",
                line=dict(color=PALETA[0], width=2), marker=dict(size=6),
                text=[f"{eq} · Nível {int(n)}" for eq, n in
                      zip(serie_pior["Equipamento (nome)"], serie_pior["NívelNum"])],
                hovertemplate="%{x}<br>Sigma: %{y:.2f}<br>%{text}<extra></extra>",
            ))
            sigma_min_serie = (serie_pior["Sigma Mínimo"].dropna()
                                if "Sigma Mínimo" in serie_pior.columns else pd.Series(dtype=float))
            if not sigma_min_serie.empty:
                meta = sigma_min_serie.iloc[-1]
                fig_sigma_mes.add_trace(go.Scatter(
                    x=serie_pior["Mês/Ano"], y=[meta] * len(serie_pior),
                    mode="lines", name="Sigma Mínimo (meta)",
                    line=dict(color="red", width=1.2, dash="dash"), opacity=0.6,
                ))
            if not ep_teste.empty:
                fig_sigma_mes.add_trace(go.Scatter(
                    x=ep_teste["Mês/Ano"], y=ep_teste["Sigma EP"], mode="markers", name="Sigma do EP (amostra)",
                    marker=dict(symbol="diamond", size=10, color=PALETA[1], line=dict(width=1, color="white")),
                    text=[f"{p} {e} · {eq} · Nível {int(n)}" for p, e, eq, n in
                          zip(ep_teste["Provedor"], ep_teste["Especime"], ep_teste["Equipamento"],
                              ep_teste["Nível CIQ"])],
                    hovertemplate="%{x}<br>Sigma EP: %{y:.2f}<br>%{text}<extra></extra>",
                ))
            fig_sigma_mes.update_layout(
                xaxis_title="Mês/Ano", **EIXO_MESES, yaxis_title="Sigma",
                height=420, hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                margin=dict(t=60),
            )
            st.plotly_chart(fig_sigma_mes, use_container_width=True)
            st.caption(
                "Pior Sigma entre todos os equipamentos/níveis daquele teste, em cada mês do período "
                "selecionado na barra lateral — mesma lógica do card \"Pior Sigma (Valor) por Nome_Mês\" "
                "montado no Power BI, pra comparar lado a lado."
            )

        st.divider()
        st.markdown(f"**Resumo do período — {teste_global}**")
        st.caption(
            "Sigma Médio (CIQ) = média simples dos Sigmas mensais já limitados entre 0 e 10 — "
            "mesmo princípio da 'Média Robusta' da planilha de desempenho, mas calculado só com "
            "dados de CIQ (a planilha original mistura CIQ com EP; quando o módulo de EP existir "
            "aqui, dá pra combinar os dois do jeito certo)."
        )
        base_resumo_sigma = df_sigma[(df_sigma["Teste"] == teste_global)].dropna(subset=["Sigma Mensal"]).copy()
        if base_resumo_sigma.empty:
            st.info("Nenhum registro de Sigma para esse teste no filtro atual.")
        else:
            base_resumo_sigma["N"] = base_resumo_sigma["N"].fillna(0)
            idx_maior_n_rs = (base_resumo_sigma.groupby(["Equipamento (nome)", "NívelNum", "_ordem_tempo"])
                               ["N"].idxmax())
            base_resumo_sigma = base_resumo_sigma.loc[idx_maior_n_rs]

            linhas_rs = []
            for (equip, nivel), grupo in base_resumo_sigma.groupby(["Equipamento (nome)", "NívelNum"]):
                grupo = grupo.sort_values("_ordem_tempo")
                sigma_medio = round(grupo["Sigma Mensal"].mean(), 2)
                idx_pior_s = grupo["Sigma Mensal"].idxmin()
                pior_sigma = round(grupo.loc[idx_pior_s, "Sigma Mensal"], 2)
                mes_pior_s = grupo.loc[idx_pior_s, "Mês/Ano"]
                linhas_rs.append({
                    "Equipamento": equip, "Nível": int(nivel),
                    "Sigma Médio (CIQ)": sigma_medio, "Pior Sigma": pior_sigma,
                    "Mês do pior": mes_pior_s, "Nº meses": len(grupo),
                })

            df_resumo_sigma = pd.DataFrame(linhas_rs).sort_values(["Equipamento", "Nível"])
            styler_resumo_sigma = (df_resumo_sigma.style
                                    .map(cor_sigma, subset=["Sigma Médio (CIQ)", "Pior Sigma"])
                                    .format("{:.2f}", subset=["Sigma Médio (CIQ)", "Pior Sigma"], na_rep="—"))
            st.dataframe(styler_resumo_sigma, hide_index=True, use_container_width=True)

        st.divider()
        df_per_filtrado = calcula_sigma_periodos_df(df_sigma)

        pf1, pf2 = st.columns(2)
        periodo_sel = pf1.radio("Período", ["Mensal", "Trimestral", "Semestral", "Anual"], horizontal=True)

        if periodo_sel == "Mensal":
            base_mensal = df_sigma.dropna(subset=["Sigma Mensal"]).copy()
            base_mensal["N"] = base_mensal["N"].fillna(0)
            idx_maior_n = (base_mensal.groupby(["Teste", "Equipamento (nome)", "NívelNum", "_ordem_tempo"])
                            ["N"].idxmax())
            base_mensal = base_mensal.loc[idx_maior_n]
            f_teste_p = pf2.multiselect("Teste", sorted(base_mensal["Teste"].dropna().unique()), key="periodo_teste")
            if f_teste_p:
                base_mensal = base_mensal[base_mensal["Teste"].isin(f_teste_p)]
            base_mensal["Mês/Ano"] = pd.to_datetime(dict(year=base_mensal["Ano"], month=base_mensal["Mês"], day=1))
            df_p = (base_mensal[["Teste", "Equipamento (nome)", "NívelNum", "Ano", "Mês/Ano",
                                  "Sigma Mensal", "CV (%)", "Bias (%)"]]
                    .rename(columns={"Equipamento (nome)": "Equipamento", "NívelNum": "Nível",
                                      "Sigma Mensal": "Sigma (pior cenário)", "Mês/Ano": "Sub-período"}))
            df_p = df_p.sort_values(["Teste", "Equipamento", "Nível", "Sub-período"])
            df_p["Nível"] = df_p["Nível"].astype(int)
        else:
            f_teste_p = pf2.multiselect("Teste", sorted(df_per_filtrado["Teste"].dropna().unique()), key="periodo_teste")
            df_p = df_per_filtrado[df_per_filtrado["Período"] == periodo_sel]
            if f_teste_p:
                df_p = df_p[df_p["Teste"].isin(f_teste_p)]
            df_p = df_p.sort_values(["Teste", "Equipamento", "Nível", "Ano", "Sub-período"]).drop(columns=["Período"])

        cols_numericas = [c for c in ["Sigma (pior cenário)", "CV (%)", "Bias (%)"] if c in df_p.columns]
        styler_p = (df_p.style.map(cor_sigma, subset=["Sigma (pior cenário)"])
                    .format("{:.2f}", subset=cols_numericas, na_rep="—"))
        st.dataframe(styler_p, hide_index=True, use_container_width=True, height=500,
                     column_config={"Sub-período": st.column_config.DateColumn("Mês", format="MM/YYYY")}
                     if periodo_sel == "Mensal" else None)
        caption_extra = " (aqui é o Sigma do mês mesmo, não um pior cenário agregado)" if periodo_sel == "Mensal" else ""
        st.caption(
            f"🟥 Sigma < 3 (inaceitável) · 🟨 Sigma entre 3 e 6 (aceitável/precisa melhorar) · "
            f"🟩 Sigma ≥ 6 (padrão Six Sigma clássico){caption_extra}. As colunas CV (%) e Bias (%) "
            "mostram os valores do mês que gerou esse Sigma, pra ajudar a identificar a causa."
        )
        testes_nivel_fixo = sorted(t for t, n in criterios_sigma.items() if n is not None)
        if testes_nivel_fixo:
            st.caption("Testes com critério \"Nível específico\" (só aquele nível entra na tabela): " +
                       ", ".join(f"{t} → Nível {criterios_sigma[t]}" for t in testes_nivel_fixo))
