"""
Módulo de CIQ - Grupo Sabin
Dashboard automático de Controle Interno da Qualidade (Infinity)

Como rodar:
    pip install streamlit pandas plotly openpyxl
    streamlit run app.py
"""
import io
import json
import re
import zipfile
from collections import defaultdict, Counter
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(page_title="Módulo CIQ - Grupo Sabin", layout="wide")

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
@st.cache_data
def load_reference():
    df_mestre = pd.read_excel(REF_DIR / "tabela_mestre.xlsx").astype(object)
    mestre = df_mestre.where(pd.notna(df_mestre), None).to_dict("records")

    df_bd = pd.read_excel(REF_DIR / "bd_fallback.xlsx").astype(object)
    df_bd = df_bd.where(pd.notna(df_bd), None)
    bd_fallback = {
        str(row["Mneumonico Infinity"]).strip().upper(): {
            "Analito": row["Analito"], "ETM (%)": row["ETM (%)"],
            "ESM (%)": row["ESM (%)"], "CV Máx (%)": row["CV Máx (%)"],
        }
        for _, row in df_bd.iterrows() if row["Mneumonico Infinity"]
    }

    df_equip = pd.read_excel(REF_DIR / "equipamentos.xlsx").astype(object)
    df_equip = df_equip.where(pd.notna(df_equip), None)
    equip_depara = {
        str(row["Equipamento"]): {
            "Tipo": row["Tipo"], "Módulo": row["Módulo"], "Célula/Rotor": row["Célula/Rotor"],
        }
        for _, row in df_equip.iterrows() if row["Equipamento"]
    }

    testes_excluidos_path = REF_DIR / "testes_excluidos.xlsx"
    if testes_excluidos_path.exists():
        df_excl = pd.read_excel(testes_excluidos_path)
        testes_excluidos = {str(t).strip().upper() for t in df_excl["Teste"].dropna()}
    else:
        testes_excluidos = set()

    mestre_by_mneu = defaultdict(list)
    for r in mestre:
        if r["Mneumonico Infinity"]:
            mestre_by_mneu[str(r["Mneumonico Infinity"]).strip().upper()].append(r)
    return mestre_by_mneu, bd_fallback, equip_depara, testes_excluidos


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
        v = v.replace(",", ".")
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
            reader = pd.read_csv(io.StringIO(raw), sep=";", dtype=str)
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
                    "Número de lote": str(row.get("Número de lote") or "").strip(),
                    "N": to_float(row.get("N")),
                    "Config. valor alvo": to_float(row.get("Config. valor alvo")),
                    "Média": to_float(row.get("Média")),
                    "CV (%)": to_float(row.get("CV (%)")),
                })
    return records, skipped, len(csv_names)


def atribui_niveis(recs):
    """Extrai o nível (1-4) do nome do controle; fallback: ranking por concentração dentro do mesmo mês."""
    for r in recs:
        r["NívelNum"] = extrai_nivel(r["Nível"])
        r["NívelOrigem"] = "nome do controle" if r["NívelNum"] is not None else None

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
        ordenados = sorted(niveis, key=lambda n: media_local.get((teste, equip, ano, mes, n), 0))
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


def processa(recs, mestre_by_mneu, bd_fallback, equip_depara, margem=0.05):
    for r in recs:
        r["Equip Mapeado"] = r["Equipamento"] in equip_depara
        r["Equipamento (nome)"] = equip_depara.get(r["Equipamento"], {}).get("Célula/Rotor", r["Equipamento"])
        r["Tipo"] = equip_depara.get(r["Equipamento"], {}).get("Tipo", "?")

        spec = get_spec(r["Teste"], r["Módulo/Arquivo"], mestre_by_mneu, bd_fallback)
        r["Spec"] = spec
        r["Bias (%)"] = calc_bias(r["Média"], r["Config. valor alvo"])

        if spec is None:
            r["CV Máximo"] = None
            r["Status CV"] = None
            r["Sigma Mensal"] = None
            r["Spec - Analito"] = None
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
        elif cv >= cvmax * (1 - margem):
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
        elif bias_obs >= bias_max * (1 - margem):
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
        elif et_obs >= etm_ref * (1 - margem):
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
st.title("Módulo de CIQ — Controle Interno da Qualidade")
st.caption("Grupo Sabin · Análise automática dos exports do Infinity")

mestre_by_mneu, bd_fallback, equip_depara, testes_excluidos = load_reference()


def nome_amigavel_equip(raw):
    info = equip_depara.get(raw)
    return info["Célula/Rotor"] if info else raw

with st.sidebar:
    st.header("Dados de entrada")
    uploaded_zip = st.file_uploader("ZIP com os CSVs mensais do Infinity", type="zip")
    st.caption(
        "A tabela de especificações (ETM, ESM, CV Máximo por nível) já vem "
        "embutida no app — só é preciso subir os dados novos do Infinity."
    )
    st.header("Margem de proximidade")
    margem_pct = st.radio(
        "Sinalizar em amarelo quando estiver a quantos % do limite?",
        [5, 10], index=1, horizontal=True,
    )
    margem = margem_pct / 100
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
    recs = atribui_niveis(recs)
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

with st.sidebar:
    st.header("Filtros globais (afetam todas as abas)")
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
        "Usado nas abas Tendência, Bias e Erro Total",
        testes_disponiveis, key="teste_global",
    )

tab_dash, tab_grafico, tab_bias, tab_et, tab_tabela, tab_periodo = st.tabs(
    ["📊 Dashboard", "📉 CV", "🎯 Bias", "⚠ Erro Total",
     "📋 Resultados Mensais", "🗓 Sigma por Período"]
)

# ---------------- DASHBOARD ----------------
with tab_dash:
    st.subheader("Visão geral — piores cenários (CV, Bias e Erro Total)")
    st.caption(
        "Considera os filtros globais da barra lateral (Período e Módulo/Plataforma). "
        "Pra investigar um teste específico em detalhe, use as abas Tendência/Comparação, "
        "Bias ou Erro Total."
    )

    status_counts = Counter(df["Status CV"].dropna())
    status_b_counts = Counter(df["Status Bias"].dropna())
    status_e_counts = Counter(df["Status Erro Total"].dropna())

    st.markdown("**CV**")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Avaliados", sum(status_counts.values()))
    c2.metric("🟩 Dentro do limite", status_counts.get("VERDE", 0))
    c3.metric(f"🟨 Próximo ({margem_pct}%)", status_counts.get("AMARELO", 0))
    c4.metric("🟥 Acima do máximo", status_counts.get("VERMELHO", 0))

    st.markdown("**Bias**")
    b1, b2, b3, b4 = st.columns(4)
    b1.metric("Avaliados", sum(status_b_counts.values()))
    b2.metric("🟩 Dentro do limite", status_b_counts.get("VERDE", 0))
    b3.metric(f"🟨 Próximo ({margem_pct}%)", status_b_counts.get("AMARELO", 0))
    b4.metric("🟥 Acima do máximo", status_b_counts.get("VERMELHO", 0))

    st.markdown("**Erro Total**")
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Avaliados", sum(status_e_counts.values()))
    e2.metric("🟩 Dentro do limite", status_e_counts.get("VERDE", 0))
    e3.metric(f"🟨 Próximo ({margem_pct}%)", status_e_counts.get("AMARELO", 0))
    e4.metric("🟥 Acima do máximo", status_e_counts.get("VERMELHO", 0))

    st.divider()
    st.subheader("Top 15 testes — CV fora da meta")
    top_verm = (df[df["Status CV"] == "VERMELHO"]["Teste"]
                .value_counts().head(15).reset_index())
    top_verm.columns = ["Teste", "Ocorrências"]
    st.dataframe(top_verm, hide_index=True, use_container_width=True)

    def bloco_ofensores(titulo, coluna_valor, coluna_maximo, nome_maximo, usa_abs=False):
        st.divider()
        st.subheader(titulo)
        st.caption(
            f"Olha o mês mais recente de cada Teste + Equipamento + Nível, juntando as duas "
            f"faixas de proximidade numa lista só — 🟧 dentro de 5% do {nome_maximo} · "
            f"🟨 dentro de 10% · 🟥 já passou do limite."
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
            elif razao >= 0.95:
                return "Dentro de 5%"
            elif razao >= 0.90:
                return "Dentro de 10%"
            return None

        recentes["Faixa"] = recentes["_razao"].apply(classifica)
        alerta = recentes[recentes["Faixa"].notna()].sort_values("_razao", ascending=False)
        if alerta.empty:
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
                    if val > maximo:
                        estilos.append("background-color: #F4CCCC")
                    elif val >= maximo * 0.95:
                        estilos.append("background-color: #FFCC80")
                    elif val >= maximo * 0.90:
                        estilos.append("background-color: #FFF9C4")
                    else:
                        estilos.append("background-color: #D9EAD3")
                else:
                    estilos.append("")
            return estilos

        styler = (traj.style.apply(cor_celula, axis=1)
                  .format("{:.2f}", subset=[nome_maximo] + meses_cols, na_rep="—"))
        st.dataframe(styler, hide_index=True, use_container_width=True)
        st.caption(
            f"🟩 dentro do limite · 🟧 dentro de 5% · 🟨 dentro de 10% · 🟥 acima do limite. "
            f"{len(alerta)} combinação(ões) no total (mostrando até 20), todos os meses do "
            f"período selecionado na barra lateral."
        )

    bloco_ofensores("Piores cenários — CV", "CV (%)", "CV Máximo", "CV Máximo")
    bloco_ofensores("Piores cenários — Bias", "Bias Observado", "Bias Máximo", "Bias Máximo")
    bloco_ofensores("Piores cenários — Erro Total", "Erro Total Observado", "ETM (para comparação)", "ETM")

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
                xaxis_title="Mês/Ano", yaxis_title="CV (%)",
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
                        xaxis_title="Mês/Ano", yaxis_title="Bias (%)", height=450,
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
                        xaxis_title="Mês/Ano", yaxis_title="CV (%)", height=450,
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
                    if val > cvmax:
                        estilos.append("background-color: #F4CCCC")
                    elif val >= cvmax * 0.95:
                        estilos.append("background-color: #FFCC80")
                    elif val >= cvmax * 0.90:
                        estilos.append("background-color: #FFF9C4")
                    else:
                        estilos.append("background-color: #D9EAD3")
                else:
                    estilos.append("")
            return estilos

        st.markdown("**CV mensal (%) por nível** — CV Pooled e pior mês em destaque, logo após o CV Máximo")
        styler_cv_tab = (df_cv_tab.style.apply(cor_cv_tab, axis=1)
                          .format("{:.2f}", subset=["CV Máximo"] + meses_ct + ["CV Pooled (%)", "Pior CV (%)"],
                                  na_rep="—"))
        st.dataframe(styler_cv_tab, hide_index=True, use_container_width=True)
        st.caption("🟩 dentro do limite · 🟧 dentro de 5% do CV Máximo · 🟨 dentro de 10% · 🟥 acima do CV Máximo.")

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
                c1, c2, c3 = st.columns(3)
                c1.metric("🟩 Dentro do limite", status_b.get("VERDE", 0))
                c2.metric("🟨 Próximo (5%)", status_b.get("AMARELO", 0))
                c3.metric("🟥 Acima do máximo", status_b.get("VERMELHO", 0))

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
                    xaxis_title="Mês/Ano", yaxis_title="Bias (%) ou valor absoluto",
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
                c1, c2, c3 = st.columns(3)
                c1.metric("🟩 Dentro do limite", status_b.get("VERDE", 0))
                c2.metric("🟨 Próximo (5%)", status_b.get("AMARELO", 0))
                c3.metric("🟥 Acima do máximo", status_b.get("VERMELHO", 0))

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
                    xaxis_title="Mês/Ano", yaxis_title="Bias (%) ou valor absoluto",
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
            else:
                linha["Bias Médio"] = None
            linhas_bt.append(linha)

        df_bias_tab = pd.DataFrame(linhas_bt)
        col_fixas_bt = {"Equipamento", "Nível", "Bias Máximo", "Bias Médio"}
        meses_bt = sorted([c for c in df_bias_tab.columns if c not in col_fixas_bt],
                           key=lambda m: ordem_mes_bt.get(m, 0))
        df_bias_tab = df_bias_tab[["Equipamento", "Nível", "Bias Máximo", "Bias Médio"] + meses_bt]
        df_bias_tab = df_bias_tab.sort_values(["Equipamento", "Nível"])

        def cor_bias_tab(row):
            bmax = row["Bias Máximo"]
            estilos = []
            for col in row.index:
                if col in (meses_bt + ["Bias Médio"]) and pd.notna(row[col]) and pd.notna(bmax):
                    val = abs(row[col])
                    if val > bmax:
                        estilos.append("background-color: #F4CCCC")
                    elif val >= bmax * 0.95:
                        estilos.append("background-color: #FFCC80")
                    elif val >= bmax * 0.90:
                        estilos.append("background-color: #FFF9C4")
                    else:
                        estilos.append("background-color: #D9EAD3")
                else:
                    estilos.append("")
            return estilos

        styler_bias_tab = (df_bias_tab.style.apply(cor_bias_tab, axis=1)
                            .format("{:.2f}", subset=["Bias Máximo", "Bias Médio"] + meses_bt,
                                    na_rep="—"))
        st.dataframe(styler_bias_tab, hide_index=True, use_container_width=True)
        st.caption(
            "🟩 dentro do limite · 🟧 dentro de 5% do Bias Máximo · 🟨 dentro de 10% · "
            "🟥 acima do Bias Máximo."
        )

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
                c1, c2, c3 = st.columns(3)
                c1.metric("🟩 Dentro do limite", status_e.get("VERDE", 0))
                c2.metric("🟨 Próximo (5%)", status_e.get("AMARELO", 0))
                c3.metric("🟥 Acima do máximo", status_e.get("VERMELHO", 0))

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
                    xaxis_title="Mês/Ano", yaxis_title="Erro Total (%) ou valor absoluto",
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
                c1, c2, c3 = st.columns(3)
                c1.metric("🟩 Dentro do limite", status_e.get("VERDE", 0))
                c2.metric("🟨 Próximo (5%)", status_e.get("AMARELO", 0))
                c3.metric("🟥 Acima do máximo", status_e.get("VERMELHO", 0))

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
                    xaxis_title="Mês/Ano", yaxis_title="Erro Total (%) ou valor absoluto",
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
    f_status = colf1.multiselect("Status CV", ["VERDE", "AMARELO", "VERMELHO"])
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
                 "Número de lote", "N", "Config. valor alvo", "Média", "CV (%)", "CV Máximo",
                 "Status CV", "Tendência CV", "Bias (%)", "Sigma Mensal", "Critério Sigma"]
    st.dataframe(df_filtrado[cols_show].rename(columns={"Equipamento (nome)": "Equipamento"}),
                 hide_index=True, use_container_width=True, height=500)
    st.caption(f"{len(df_filtrado)} de {len(df)} registros exibidos")

    csv = df_filtrado[cols_show].to_csv(index=False).encode("utf-8-sig")
    st.download_button("Baixar CSV filtrado", csv, "ciq_resultados.csv", "text/csv")

# ---------------- SIGMA POR PERÍODO ----------------
with tab_periodo:
    st.subheader("Sigma — mensal e por período")
    st.caption(f"Teste: **{teste_global}** (mude na barra lateral, em 'Teste em foco')")
    st.caption(
        "Recalculado a partir dos dados já filtrados pela barra lateral (Período e "
        "Módulo/Plataforma) — mudar o período lá em cima muda o que aparece aqui."
    )

    card_pior_cenario(df[df["Teste"] == teste_global], "Sigma Mensal", "Sigma", maior_eh_pior=False)

    st.markdown(f"**Resumo do período — {teste_global}**")
    st.caption(
        "Sigma Médio (CIQ) = média simples dos Sigmas mensais já limitados entre 0 e 10 — "
        "mesmo princípio da 'Média Robusta' da planilha de desempenho, mas calculado só com "
        "dados de CIQ (a planilha original mistura CIQ com EP; quando o módulo de EP existir "
        "aqui, dá pra combinar os dois do jeito certo)."
    )
    base_resumo_sigma = df[(df["Teste"] == teste_global)].dropna(subset=["Sigma Mensal"]).copy()
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
            linhas_rs.append({
                "Equipamento": equip, "Nível": int(nivel),
                "Sigma Médio (CIQ)": sigma_medio, "Nº meses": len(grupo),
            })

        df_resumo_sigma = pd.DataFrame(linhas_rs).sort_values(["Equipamento", "Nível"])
        styler_resumo_sigma = (df_resumo_sigma.style
                                .map(cor_sigma, subset=["Sigma Médio (CIQ)"])
                                .format("{:.2f}", subset=["Sigma Médio (CIQ)"], na_rep="—"))
        st.dataframe(styler_resumo_sigma, hide_index=True, use_container_width=True)

    st.divider()
    df_per_filtrado = calcula_sigma_periodos_df(df[df["Teste"] == teste_global])

    periodo_sel = st.radio("Período", ["Mensal", "Trimestral", "Semestral", "Anual"], horizontal=True)

    if periodo_sel == "Mensal":
        base_mensal = df[df["Teste"] == teste_global].dropna(subset=["Sigma Mensal"]).copy()
        base_mensal["N"] = base_mensal["N"].fillna(0)
        idx_maior_n = (base_mensal.groupby(["Teste", "Equipamento (nome)", "NívelNum", "_ordem_tempo"])
                        ["N"].idxmax())
        base_mensal = base_mensal.loc[idx_maior_n]
        df_p = (base_mensal[["Teste", "Equipamento (nome)", "NívelNum", "Ano", "Mês/Ano",
                              "Sigma Mensal", "CV (%)", "Bias (%)"]]
                .rename(columns={"Equipamento (nome)": "Equipamento", "NívelNum": "Nível",
                                  "Sigma Mensal": "Sigma (pior cenário)", "Mês/Ano": "Sub-período"}))
        df_p = df_p.sort_values(["Teste", "Equipamento", "Nível", "Ano"])
        df_p["Nível"] = df_p["Nível"].astype(int)
    else:
        df_p = df_per_filtrado[df_per_filtrado["Período"] == periodo_sel]
        df_p = df_p.sort_values(["Teste", "Equipamento", "Nível", "Ano", "Sub-período"]).drop(columns=["Período"])

    cols_numericas = [c for c in ["Sigma (pior cenário)", "CV (%)", "Bias (%)"] if c in df_p.columns]
    styler_p = (df_p.style.map(cor_sigma, subset=["Sigma (pior cenário)"])
                .format("{:.2f}", subset=cols_numericas, na_rep="—"))
    st.dataframe(styler_p, hide_index=True, use_container_width=True, height=500)
    caption_extra = " (aqui é o Sigma do mês mesmo, não um pior cenário agregado)" if periodo_sel == "Mensal" else ""
    st.caption(
        f"🟥 Sigma < 3 (inaceitável) · 🟨 Sigma entre 3 e 6 (aceitável/precisa melhorar) · "
        f"🟩 Sigma ≥ 6 (padrão Six Sigma clássico){caption_extra}. As colunas CV (%) e Bias (%) "
        "mostram os valores do mês que gerou esse Sigma, pra ajudar a identificar a causa."
    )
