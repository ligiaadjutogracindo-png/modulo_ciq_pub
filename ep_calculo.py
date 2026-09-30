"""
Cálculo do Sigma do EP (Ensaio de Proficiência) — mesma lógica da aba EP da planilha
(conferido contra "06 - Bili Direta.xlsx", rodadas ControlLab jan/abr/jul 2026).

Por rodada (provedor + kit + teste), com o equipamento informado pelo usuário:
1. Bias% de cada amostra: (RL − VD) / VD × 100.
2. Pareamento amostra × nível de CIQ do equipamento no mês da rodada:
   - Por concentração (padrão): cada amostra usa o nível de CIQ com média mais próxima do VD
     (em escala relativa). O CV do Sigma passa a ser o da mesma faixa de concentração em que o
     viés foi medido. Quando o nível mais próximo ainda está longe (VD < ½ ou > 2× a média do
     CIQ), a amostra é marcada — o Sigma é só indicativo, não há controle naquela faixa.
   - Por posição (como a planilha): amostra 1 → Nível 1, amostra 2 → Nível 2...
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
PAREAMENTOS = ["Por concentração", "Por posição"]
RAZAO_CONCENTRACAO_REGRESSAO = 2.0   # maior VD / menor VD a partir do qual o Automático usa regressão
R_MINIMO_REGRESSAO = 0.95
FAIXA_NIVEL_PROXIMO = (0.5, 2.0)     # VD / média do CIQ fora disso: nível "longe" da amostra


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


def nivel_mais_proximo(vd, ciq_por_nivel):
    """Nível de CIQ com média mais próxima do VD, em escala relativa (|log(VD/média)|)."""
    if not _valido(vd) or vd <= 0:
        return None
    candidatos = [(abs(math.log(vd / c["Média"])), n) for n, c in ciq_por_nivel.items()
                  if _valido(c.get("Média")) and c["Média"] > 0]
    return min(candidatos)[1] if candidatos else None


def calcula_rodada(amostras, ciq_por_nivel, spec, modo="Médio", formula="Automático",
                   pareamento="Por concentração"):
    """
    amostras: lista de dicts com "Especime", "Num", "RL", "VD" (já na unidade do laboratório),
              "Qualificador" ("<", ">" ou ""), e opcionais "DP Grupo" e "Índice Provedor".
              A posição (1, 2, 3...) vem da ordem do Num.
    ciq_por_nivel: {nível: {"Média": ..., "CV (%)": ...}} do equipamento no mês da rodada.
    spec: especificação da tabela_mestre (ETM %, ETM absoluto e cutoff).
    modo: "Médio", "Regressão" ou "Automático"; formula: "Automático", "σ abs" ou "σ %";
    pareamento: "Por concentração" ou "Por posição".
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
        nivel = (nivel_mais_proximo(vd, ciq_por_nivel) if pareamento == "Por concentração"
                 else (pos if pos in ciq_por_nivel else None))
        ciq = ciq_por_nivel.get(nivel) or {}
        media_ciq, cv_ciq = ciq.get("Média"), ciq.get("CV (%)")
        razao = vd / media_ciq if _valido(vd) and _valido(media_ciq) and media_ciq else None
        longe = razao is not None and not (FAIXA_NIVEL_PROXIMO[0] <= razao <= FAIXA_NIVEL_PROXIMO[1])
        dp_grupo = a.get("DP Grupo")
        iz = (rl - vd) / dp_grupo if ok and _valido(dp_grupo) and dp_grupo else None
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
            "Especime": a["Especime"], "Posição": pos, "Nível CIQ": nivel if ciq else None,
            "RL": rl, "VD": vd, "Qualificador": a.get("Qualificador") or "",
            "Bias % amostra": (rl - vd) / vd * 100 if ok and vd else None,
            "Índice Provedor": a.get("Índice Provedor"), "IZ": iz,
            "Média CIQ": media_ciq, "VD/Média CIQ": razao,
            "Nível longe": "sim — Sigma indicativo" if longe else "",
            "CV CIQ (%)": cv_ciq,
            "DP CIQ": cv_ciq * media_ciq / 100 if _valido(cv_ciq) and _valido(media_ciq) else None,
            "Viés abs": vies_abs, "Viés %": vies_pct,
            "Sigma EP": sigma, "Critério Sigma": criterio,
        })
    info["Pareamento"] = pareamento
    return linhas, info


# ------------------------------------------------------------------
# Análises de tendência
# ------------------------------------------------------------------
LIMITE_IZ_ALERTA, LIMITE_IZ_ACAO = 2.0, 3.0      # |IZ| (escore z) — critérios usuais da ISO 13528


def analisa_rodada(linhas, info, esm_pct=None):
    """Padrões dentro da rodada: viés do mesmo lado, erro proporcional/constante, IZ alto."""
    msgs = []
    validas = [l for l in linhas if _valido(l.get("Bias % amostra"))]
    bias = [l["Bias % amostra"] for l in validas]
    if len(bias) >= 2 and (all(b < 0 for b in bias) or all(b > 0 for b in bias)):
        media = sum(bias) / len(bias)
        lado = "abaixo" if media < 0 else "acima"
        grave = _valido(esm_pct) and abs(media) > esm_pct / 2
        msgs.append(f"todas as amostras {lado} do valor designado (média {media:+.1f}%)"
                    + (" — erro sistemático relevante (> metade do ESM)" if grave else ""))
    vds = [l["VD"] for l in validas if l["VD"] and l["VD"] > 0]
    b, a, r = info.get("b"), info.get("a"), info.get("r")
    if len(validas) >= 3 and b is not None and r is not None and vds and max(vds) / min(vds) >= 1.5:
        if r < R_MINIMO_REGRESSAO:
            msgs.append(f"resultados pouco alinhados (r = {r:.2f}) — sugere variação aleatória")
        else:
            # viés previsto pela reta nas pontas da faixa da rodada; só é padrão relevante se mudar
            # mais que metade do ESM (com 3 amostras próximas, inclinação e intercepto se compensam)
            vies_min = ((b * min(vds) + a) - min(vds)) / min(vds) * 100
            vies_max = ((b * max(vds) + a) - max(vds)) / max(vds) * 100
            limite = esm_pct / 2 if _valido(esm_pct) else 5.0
            if abs(vies_max - vies_min) > limite:
                tipo = "proporcional" if abs(b - 1) >= abs(a) / (sum(vds) / len(vds)) else "constante"
                msgs.append(f"o viés muda com a concentração ({vies_min:+.1f}% em {min(vds):g} → {vies_max:+.1f}% "
                            f"em {max(vds):g}; erro {tipo}) — prefira o viés por regressão")
    izs = [abs(l["IZ"]) for l in linhas if _valido(l.get("IZ"))]
    if izs:
        pior = max(izs)
        if pior > LIMITE_IZ_ACAO:
            msgs.append(f"|IZ| {pior:.1f} > {LIMITE_IZ_ACAO:g} — resultado insatisfatório")
        elif pior > LIMITE_IZ_ALERTA:
            msgs.append(f"|IZ| {pior:.1f} > {LIMITE_IZ_ALERTA:g} — resultado questionável")
    return "; ".join(msgs) or "sem padrão de erro na rodada"


def analisa_historico(rodadas):
    """rodadas: lista (em ordem de data) de dicts com "Viés %", "Pior |IZ|" e "Rótulo".
    Devolve alertas de tendência entre rodadas."""
    alertas = []
    vieses = [(r["Rótulo"], r["Viés %"]) for r in rodadas if _valido(r.get("Viés %"))]
    if len(vieses) >= 3:
        ult = [v for _, v in vieses]
        n = 1
        while n < len(ult) and (ult[-1 - n] > 0) == (ult[-1] > 0) and ult[-1 - n] != 0:
            n += 1
        if n >= 3:
            lado = "acima" if ult[-1] > 0 else "abaixo"
            alertas.append(f"viés {lado} do valor designado nas últimas {n} rodadas seguidas — "
                           "possível erro sistemático persistente")
        a3 = [abs(v) for v in ult[-3:]]
        if a3[0] < a3[1] < a3[2]:
            alertas.append(f"viés aumentando nas últimas 3 rodadas ({ult[-3]:+.1f}% → {ult[-2]:+.1f}% → "
                           f"{ult[-1]:+.1f}%) — tendência de piora")
    if len(vieses) >= 2:
        (r_ant, v_ant), (r_ult, v_ult) = vieses[-2], vieses[-1]
        alertas.append(f"última rodada ({r_ult}): viés {v_ult:+.1f}%, {v_ult - v_ant:+.1f} pontos em relação à "
                       f"anterior ({r_ant})")
    izs = [r.get("Pior |IZ|") for r in rodadas[-3:] if _valido(r.get("Pior |IZ|"))]
    altos = sum(1 for z in izs if z > LIMITE_IZ_ALERTA)
    if altos >= 2:
        alertas.append(f"|IZ| > {LIMITE_IZ_ALERTA:g} em {altos} das últimas {len(izs)} rodadas — investigar")
    return alertas
