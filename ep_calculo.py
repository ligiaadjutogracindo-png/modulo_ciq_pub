"""
Cálculo do Sigma do EP (Ensaio de Proficiência) — mesma lógica da aba EP da planilha
(conferido contra "06 - Bili Direta.xlsx", rodadas ControlLab jan/abr/jul 2026).

Por rodada (provedor + kit + teste), com o equipamento informado pelo usuário:
1. Bias% de cada amostra: (RL − VD) / VD × 100.
2. Amostra na posição N usa o Nível N de CIQ do equipamento no mês da rodada
   (posição fixa, como a planilha — não é por concentração mais próxima).
3. DPCIQ = CV × MédiaCIQ / 100.
4. Viés da rodada (na planilha, lista "Critério Viés para estimar o Sigma"):
   - Médio: média(RL) − média(VD), um só para todas as amostras;
   - Regressão: RL = b·VD + a (b e a com 3 casas), avaliado na média do CIQ de cada nível;
   - Automático: regressão quando as amostras têm concentrações bem diferentes entre si
     (maior VD ≥ 2 × menor VD), há ≥ 3 amostras e r ≥ 0,95; senão, médio.
   Sem pares suficientes pra regressão, cai pro viés médio (com aviso).
5. Sigma (na planilha, lista "σ abs" / "σ %"):
   - σ abs: (MédiaCIQ × ETM%/100 − |Viés abs|) / DPCIQ
   - σ %:   (ETM% − |Viés %|) / CVCIQ
   - Automático: σ abs quando alguma amostra da rodada está no cutoff do ETM absoluto ou
     abaixo; senão σ %.
   Quando a MédiaCIQ do nível está no cutoff ou abaixo, vale o ETM absoluto nas duas:
   (ETM abs − |Viés abs|) / DPCIQ. (A planilha, no σ %, subtrai um viés em % do ETM em
   mg/dL nesse caso; aqui usa a forma absoluta, que é a coerente.)
   Com viés por regressão as duas fórmulas dão o mesmo resultado (o viés % é o do nível).
   Limitado a 0–10 (a planilha limita em 8).
"""
import math

MODOS_VIES = ["Médio", "Regressão", "Automático"]
FORMULAS_SIGMA = ["Automático", "σ abs", "σ %"]
RAZAO_CONCENTRACAO_REGRESSAO = 2.0   # maior VD / menor VD a partir do qual o Automático usa regressão
R_MINIMO_REGRESSAO = 0.95


def _valido(v):
    return v is not None and not (isinstance(v, float) and math.isnan(v))


def clip(v, lo=0.0, hi=10.0):
    return None if v is None else max(lo, min(hi, v))


def regressao(vd, rl):
    """RL = b·VD + a, com b e a arredondados a 3 casas (como a planilha), e o r de Pearson.
    None quando não dá pra calcular (menos de 2 pares ou VD sem variação)."""
    pares = [(x, y) for x, y in zip(vd, rl) if _valido(x) and _valido(y)]
    if len(pares) < 2:
        return None
    n = len(pares)
    mx = sum(x for x, _ in pares) / n
    my = sum(y for _, y in pares) / n
    sxx = sum((x - mx) ** 2 for x, _ in pares)
    syy = sum((y - my) ** 2 for _, y in pares)
    sxy = sum((x - mx) * (y - my) for x, y in pares)
    if sxx == 0:
        return None
    b = sxy / sxx
    a = my - b * mx
    r = sxy / math.sqrt(sxx * syy) if syy > 0 else None
    return round(b, 3), round(a, 3), r


def _cutoff(spec):
    if not spec:
        return None, None
    cutoff, etm_abs = spec.get("ETM Absoluto - Cutoff"), spec.get("ETM Absoluto - Valor")
    return (cutoff, etm_abs) if _valido(cutoff) and _valido(etm_abs) else (None, None)


def modo_automatico(validas, reg):
    """Regressão quando as concentrações da rodada são bem diferentes e a reta é confiável."""
    vds = [a["VD"] for a in validas if a["VD"] > 0]
    if len(validas) < 3 or not reg or reg[2] is None or len(vds) < 2:
        return "Médio", "menos de 3 amostras válidas ou regressão indisponível"
    razao = max(vds) / min(vds)
    if razao >= RAZAO_CONCENTRACAO_REGRESSAO and reg[2] >= R_MINIMO_REGRESSAO:
        return "Regressão", f"concentrações variam {razao:.1f}× e r = {reg[2]:.3f}"
    if razao < RAZAO_CONCENTRACAO_REGRESSAO:
        return "Médio", f"concentrações próximas (variam {razao:.1f}×)"
    return "Médio", f"r = {reg[2]:.3f} abaixo de {R_MINIMO_REGRESSAO}"


def formula_automatica(validas, spec):
    cutoff, _ = _cutoff(spec)
    if cutoff is not None and any(a["VD"] <= cutoff for a in validas):
        return "σ abs", f"há amostra ≤ cutoff ({cutoff:g})"
    return "σ %", "nenhuma amostra ≤ cutoff" if cutoff is not None else "teste sem ETM absoluto"


def sigma_ep(media_ciq, cv_ciq, vies_abs, vies_pct, spec, formula="σ abs"):
    """Sigma de uma amostra. Devolve (sigma, critério)."""
    if not all(_valido(v) for v in (media_ciq, cv_ciq, vies_abs)) or not spec or not cv_ciq:
        return None, None
    dp_ciq = cv_ciq * media_ciq / 100
    if dp_ciq == 0:
        return None, None
    cutoff, etm_abs = _cutoff(spec)
    if cutoff is not None and media_ciq <= cutoff:
        return clip((etm_abs - abs(vies_abs)) / dp_ciq), "ETM absoluto (≤ cutoff)"
    etm_pct = spec.get("ETM (%)")
    if not _valido(etm_pct):
        return None, None
    if formula == "σ %" and _valido(vies_pct):
        return clip((etm_pct - abs(vies_pct)) / cv_ciq), "σ %"
    # atenção à escala: ETM% / 100 × MédiaCIQ (não ETM% direto)
    return clip((media_ciq * etm_pct / 100 - abs(vies_abs)) / dp_ciq), "σ abs"


def calcula_rodada(amostras, ciq_por_nivel, spec, modo="Médio", formula="Automático"):
    """
    amostras: lista de dicts com "Especime", "Num", "RL", "VD" (já na unidade do laboratório)
              e "Qualificador" ("<", ">" ou ""). A posição (1, 2, 3...) vem da ordem do Num.
    ciq_por_nivel: {nível: {"Média": ..., "CV (%)": ...}} do equipamento no mês da rodada.
    spec: especificação da tabela_mestre (ETM %, ETM absoluto e cutoff).
    modo: "Médio", "Regressão" ou "Automático"; formula: "Automático", "σ abs" ou "σ %".
    Retorna (linhas, info): uma linha por amostra e um resumo da rodada.
    """
    ordenadas = sorted(amostras, key=lambda a: a["Num"])
    validas = [a for a in ordenadas
               if _valido(a.get("RL")) and _valido(a.get("VD")) and not a.get("Qualificador")]

    info = {"Modo pedido": modo, "Modo usado": None, "Fórmula pedida": formula, "Fórmula usada": None,
            "Motivo": "", "Aviso": None, "b": None, "a": None, "r": None,
            "Viés médio abs": None, "Viés médio %": None}
    vies_medio = None
    if len(validas) >= 2:
        media_rl = sum(a["RL"] for a in validas) / len(validas)
        media_vd = sum(a["VD"] for a in validas) / len(validas)
        vies_medio = media_rl - media_vd
        info["Viés médio abs"] = vies_medio
        info["Viés médio %"] = vies_medio / media_vd * 100 if media_vd else None

    reg = regressao([a["VD"] for a in validas], [a["RL"] for a in validas])
    if reg:
        info["b"], info["a"], info["r"] = reg
    motivos = []
    modo_efetivo = modo
    if modo == "Automático":
        modo_efetivo, motivo = modo_automatico(validas, reg)
        motivos.append(f"viés {modo_efetivo.lower()}: {motivo}")
    if modo_efetivo == "Regressão" and reg:
        info["Modo usado"] = "Regressão"
    elif vies_medio is not None:
        info["Modo usado"] = "Médio"
        if modo_efetivo == "Regressão":
            info["Aviso"] = "Não foi possível calcular a regressão — usado o viés médio para estimar o Sigma."
    else:
        info["Aviso"] = "Menos de 2 amostras válidas — sem viés para estimar o Sigma."

    formula_efetiva = formula
    if formula == "Automático":
        formula_efetiva, motivo = formula_automatica(validas, spec)
        motivos.append(f"{formula_efetiva}: {motivo}")
    info["Fórmula usada"] = formula_efetiva
    info["Motivo"] = "; ".join(motivos)

    linhas = []
    for pos, a in enumerate(ordenadas, start=1):
        rl, vd = a.get("RL"), a.get("VD")
        ok = a in validas
        ciq = ciq_por_nivel.get(pos) or {}
        media_ciq, cv_ciq = ciq.get("Média"), ciq.get("CV (%)")
        if info["Modo usado"] == "Regressão" and _valido(media_ciq):
            vies_abs = (media_ciq * info["b"] + info["a"]) - media_ciq
            vies_pct = vies_abs / media_ciq * 100 if media_ciq else None
        elif info["Modo usado"] == "Médio":
            vies_abs, vies_pct = vies_medio, info["Viés médio %"]
        else:
            vies_abs = vies_pct = None
        sigma, criterio = (sigma_ep(media_ciq, cv_ciq, vies_abs, vies_pct, spec, formula_efetiva)
                           if ciq else (None, None))
        linhas.append({
            "Especime": a["Especime"], "Posição": pos, "Nível CIQ": pos if ciq else None,
            "RL": rl, "VD": vd, "Qualificador": a.get("Qualificador") or "",
            "Bias % amostra": (rl - vd) / vd * 100 if ok and vd else None,
            "Média CIQ": media_ciq, "CV CIQ (%)": cv_ciq,
            "DP CIQ": cv_ciq * media_ciq / 100 if _valido(cv_ciq) and _valido(media_ciq) else None,
            "Viés abs": vies_abs, "Viés %": vies_pct,
            "Sigma EP": sigma, "Critério Sigma": criterio,
        })
    return linhas, info
