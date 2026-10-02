"""
Relatório de análise crítica (PDF) do app Desempenho Analítico.

Feito com o matplotlib (sem outra biblioteca de PDF): para o teste e o mês escolhidos, cada parte
marcada pelo usuário entra com a tabela ou os gráficos e, logo abaixo, a análise crítica escrita no
app — ou um quadro em branco, quando a análise vai ser feita em outra ferramenta.
"""
import io
import textwrap
from datetime import datetime

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Rectangle  # noqa: E402

# partes que o usuário pode levar para a análise crítica (chave → título no PDF / nome curto na tela)
SECOES = {
    "resumo_ciq": ("CIQ do mês — resumo por equipamento e nível", "CIQ do mês (tabela)"),
    "cv": ("CV ao longo do período", "CV"),
    "bias": ("Bias ao longo do período", "Bias"),
    "erro_total": ("Erro total ao longo do período", "Erro total"),
    "sigma": ("Sigma ao longo do período", "Sigma"),
    "ep_resultados": ("Ensaio de proficiência — resultados das rodadas", "EP — resultados"),
    "ep_historico": ("Ensaio de proficiência — histórico (ID, Z e viés)", "EP — histórico"),
}
MESES = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
         7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}

A4 = (11.69, 8.27)                 # paisagem, em polegadas
ESQ, DIR, TOPO, BASE = 0.045, 0.955, 0.94, 0.075
COR = "#4B2E5A"
VERDE, VERMELHO, CINZA = "#D9EAD3", "#F4CCCC", "#F3F3F3"
PALETA = ["#4B2E5A", "#C2185B", "#00838F", "#F57F17", "#2E7D32", "#5D4037",
          "#1565C0", "#AD1457", "#00695C", "#EF6C00", "#6A1B9A", "#33691E"]
CORES_AMOSTRA = ["#1565C0", "#2E7D32", "#EC407A", "#F57F17", "#6A1B9A", "#00838F", "#5D4037"]


def _txt(v):
    """Texto seguro para o matplotlib ("$" ativaria fórmulas)."""
    return str(v).replace("$", r"\$")


def _num(v, casas=2):
    return "—" if v is None or pd.isna(v) else f"{v:.{casas}f}".replace(".", ",")


def _g(v):
    """Número curto com vírgula decimal (13, 6,5, 4,33)."""
    return f"{v:g}".replace(".", ",")


class _Relatorio:
    """Páginas A4 com fluxo de cima para baixo: texto, tabela e gráficos, quebrando página quando falta espaço."""

    def __init__(self, rodape):
        self.buf = io.BytesIO()
        self.pdf = PdfPages(self.buf, metadata={"Title": rodape, "Creator": "Desempenho Analítico"})
        self.rodape = rodape
        self.fig = None
        self.y = TOPO
        self.paginas = 0

    def _fecha_pagina(self):
        if self.fig is None:
            return
        self.fig.text(ESQ, 0.035, _txt(self.rodape), fontsize=7, color="#777777")
        self.fig.text(DIR, 0.035, f"Página {self.paginas}", fontsize=7, color="#777777", ha="right")
        self.pdf.savefig(self.fig)
        plt.close(self.fig)
        self.fig = None

    def nova_pagina(self):
        self._fecha_pagina()
        self.fig = plt.figure(figsize=A4)
        self.paginas += 1
        self.y = TOPO

    def cabe(self, altura):
        if self.fig is None or self.y - altura < BASE:
            self.nova_pagina()

    def espaco(self, altura=0.012):
        self.y -= altura

    def texto(self, texto, tamanho=9.0, negrito=False, cor="black", recuo=0.0):
        altura = tamanho * 1.5 / 72 / A4[1]
        chars = int((DIR - ESQ - recuo) * A4[0] * 72 / (tamanho * 0.56))
        linhas = []
        for paragrafo in str(texto).split("\n"):
            linhas += textwrap.wrap(paragrafo, chars) or [""]
        for linha in linhas:
            self.cabe(altura)
            self.fig.text(ESQ + recuo, self.y, _txt(linha), fontsize=tamanho, color=cor, va="top",
                          weight="bold" if negrito else "normal")
            self.y -= altura

    def titulo_secao(self, texto, reserva=0.12):
        self.cabe(reserva)    # o título não fica sozinho no pé da página: já reserva o espaço do conteúdo
        self.espaco(0.012)
        self.fig.add_artist(Line2D([ESQ, DIR], [self.y, self.y], color=COR, lw=0.8))
        self.espaco(0.01)
        self.texto(texto, tamanho=11.5, negrito=True, cor=COR)
        self.espaco(0.004)

    def quadro_analise(self, texto, rotulo="Análise crítica"):
        """Texto da análise crítica, ou um quadro em branco para preencher em outra ferramenta."""
        self.espaco(0.006)
        self.texto(f"{rotulo}:", tamanho=9, negrito=True, cor=COR)
        if str(texto or "").strip():
            self.texto(texto, tamanho=9, recuo=0.01)
        else:
            self.cabe(0.09)
            self.fig.add_artist(Rectangle((ESQ, self.y - 0.08), DIR - ESQ, 0.08, facecolor="none",
                                          edgecolor="#BBBBBB", lw=0.6))
            self.y -= 0.085

    def tabela(self, dados, cores=None, tamanho=7.5):
        """dados: DataFrame já em texto; cores: DataFrame igual com a cor de fundo de cada célula ("" = sem cor)."""
        if dados.empty:
            self.texto("(sem dados no período)", tamanho=8.5, cor="#777777")
            return
        alt = tamanho * 2.0 / 72 / A4[1]
        colunas = list(dados.columns)
        pesos = [min(max(len(str(c)), *(len(str(v)) for v in dados[c]), 4), 34) for c in colunas]
        larg = DIR - ESQ
        xs = [ESQ]
        for p in pesos:
            xs.append(xs[-1] + larg * p / sum(pesos))

        def cabecalho():
            self.cabe(alt * 2)
            self.fig.add_artist(Rectangle((ESQ, self.y - alt), larg, alt, facecolor=COR, edgecolor="none"))
            for i, c in enumerate(colunas):
                self.fig.text(xs[i] + 0.003, self.y - alt / 2, _txt(c), fontsize=tamanho, color="white",
                              weight="bold", va="center")
            self.y -= alt

        cabecalho()
        for r in range(len(dados)):
            if self.y - alt < BASE:
                self.nova_pagina()
                cabecalho()
            for i, c in enumerate(colunas):
                fundo = cores.iat[r, i] if cores is not None else ""
                fundo = fundo or (CINZA if r % 2 else "")
                if fundo:
                    self.fig.add_artist(Rectangle((xs[i], self.y - alt), xs[i + 1] - xs[i], alt,
                                                  facecolor=fundo, edgecolor="none"))
                valor = str(dados.iat[r, i])
                limite = pesos[i] + 1
                valor = valor if len(valor) <= limite else valor[:limite - 1] + "…"
                self.fig.text(xs[i] + 0.003, self.y - alt / 2, _txt(valor), fontsize=tamanho, va="center")
            self.fig.add_artist(Line2D([ESQ, DIR], [self.y - alt, self.y - alt], color="#DDDDDD", lw=0.4))
            self.y -= alt
        self.espaco(0.008)

    def graficos(self, desenhos, altura=0.34, margem_inferior=0.075):
        """desenhos: funções f(ax), lado a lado. margem_inferior: espaço para os meses (e a legenda) abaixo."""
        self.cabe(altura)
        n = len(desenhos)
        intervalo = 0.045
        larg = (DIR - ESQ - intervalo * (n - 1)) / n
        for i, desenha in enumerate(desenhos):
            ax = self.fig.add_axes([ESQ + 0.02 + i * (larg + intervalo), self.y - altura + margem_inferior,
                                    larg - 0.02, altura - margem_inferior - 0.035])
            desenha(ax)
        self.y -= altura

    def fecha(self):
        self._fecha_pagina()
        self.pdf.close()
        return self.buf.getvalue()


# ------------------------------------------------------------------
# Partes do relatório
# ------------------------------------------------------------------
def _um_por_mes(ciq):
    """Mesmo critério das abas: com mais de um lote no mês/equipamento/nível, fica o de maior N."""
    if ciq.empty:
        return ciq
    ciq = ciq.assign(N=ciq["N"].fillna(0))
    idx = ciq.groupby(["Equipamento (nome)", "NívelNum", "_ordem_tempo"])["N"].idxmax()
    return ciq.loc[idx]


def _eixo_meses(ax, rotulos):
    passo = max(1, -(-len(rotulos) // 12))    # até ~12 rótulos no eixo (com muitos meses, mostra um sim, outro não)
    ax.set_xticks(range(0, len(rotulos), passo))
    ax.set_xticklabels(rotulos[::passo], rotation=45 if len(rotulos) > 6 else 0,
                       ha="right" if len(rotulos) > 6 else "center", fontsize=6.5)
    ax.set_xlim(-0.5, len(rotulos) - 0.5)
    ax.tick_params(axis="y", labelsize=7)
    ax.grid(alpha=0.25)


def _grafico_ciq(dados, meses, coluna, limite, titulo, simetrico=False, maior_melhor=False):
    pos = {o: i for i, o in enumerate(meses["_ordem_tempo"])}

    def desenha(ax):
        for k, (equip, g) in enumerate(dados.groupby("Equipamento (nome)")):
            g = g.dropna(subset=[coluna]).sort_values("_ordem_tempo")
            if g.empty:
                continue
            ax.plot([pos[o] for o in g["_ordem_tempo"]], g[coluna], marker="o", ms=3, lw=1.2,
                    color=PALETA[k % len(PALETA)], label=_txt(equip))
        valores_lim = dados[limite].dropna() if limite in dados else pd.Series(dtype=float)
        if len(valores_lim):
            v = float(valores_lim.iloc[-1])
            ax.axhline(v, color="#C62828", ls="--", lw=1)
            if simetrico:
                ax.axhline(-v, color="#C62828", ls="--", lw=1)
        if simetrico:
            ax.axhline(0, color="#999999", lw=0.6)
        ax.set_title(_txt(titulo), fontsize=9)
        _eixo_meses(ax, list(meses["Mês/Ano"]))
        if ax.get_legend_handles_labels()[0]:   # legenda embaixo do gráfico: não cobre as linhas
            ax.legend(fontsize=5.5, loc="upper center", bbox_to_anchor=(0.5, -0.27), ncol=3, frameon=False)
        if maior_melhor:
            ax.set_ylim(bottom=0)

    return desenha


# gráficos do CIQ por nível: coluna do valor, coluna do limite, limite dos dois lados, maior é melhor, rótulo
GRAFICOS_CIQ = {
    "cv": ("CV (%)", "CV Máximo", False, False, "CV (%)"),
    "bias": ("Bias Observado (sinal)", "Bias Máximo", True, False, "Bias"),
    "erro_total": ("Erro Total Observado", "ETM (para comparação)", False, False, "Erro total"),
    "sigma": ("Sigma Mensal", "Sigma Mínimo", False, True, "Sigma"),
}
LEGENDA_CIQ = "Linha vermelha tracejada: limite da especificação. Uma cor por equipamento."
LEGENDA_EP = ("Cada cor é uma amostra da rodada (1ª, 2ª, 3ª...); losango = ControlLab, círculo = CAP. "
              "Linha cinza: média da rodada. Faixa verde: ID ±1, Z ±2 e viés dentro do ESM.")


def _resumo_ciq(rel, ciq_mes, completo):
    if ciq_mes.empty:
        rel.texto("(sem CIQ do teste no mês)", tamanho=8.5, cor="#777777")
        return
    d = ciq_mes.sort_values(["Equipamento (nome)", "NívelNum"])
    tab = pd.DataFrame({
        "Equipamento": d["Equipamento (nome)"], "Nível": d["NívelNum"].map(lambda n: f"{int(n)}"),
        "Controle": d["Nível"], "Lote": d["Número de lote"], "N": d["N"].map(lambda v: _num(v, 0)),
        "Média": d["Média"].map(_num), "CV %": d["CV (%)"].map(_num), "CV máx": d["CV Máximo"].map(_num),
        "Bias": d["Bias Observado (sinal)"].map(_num), "Bias máx": d["Bias Máximo"].map(_num),
        "Erro total": d["Erro Total Observado"].map(_num), "ETM": d["ETM (para comparação)"].map(_num),
    })
    cores = pd.DataFrame("", index=tab.index, columns=tab.columns)

    def pinta(coluna, fora, valido):
        cores.loc[valido, coluna] = [VERMELHO if f else VERDE for f in fora[valido]]

    pinta("CV %", d["CV (%)"] > d["CV Máximo"], d["CV (%)"].notna() & d["CV Máximo"].notna())
    pinta("Bias", d["Bias Observado (sinal)"].abs() > d["Bias Máximo"],
          d["Bias Observado (sinal)"].notna() & d["Bias Máximo"].notna())
    pinta("Erro total", d["Erro Total Observado"] > d["ETM (para comparação)"],
          d["Erro Total Observado"].notna() & d["ETM (para comparação)"].notna())
    if completo:
        tab["Sigma"] = d["Sigma Mensal"].map(_num)
        tab["Sigma mín"] = d["Sigma Mínimo"].map(lambda v: _num(v, 1))
        cores["Sigma"] = ""
        cores["Sigma mín"] = ""
        pinta("Sigma", d["Sigma Mensal"] < d["Sigma Mínimo"], d["Sigma Mensal"].notna() & d["Sigma Mínimo"].notna())
    rel.tabela(tab.reset_index(drop=True), cores.reset_index(drop=True))
    rel.texto("Verde/vermelho = dentro/fora do limite. Bias e erro total em % (ou na unidade do exame, quando a "
              "especificação é absoluta).", tamanho=7, cor="#777777")


def _ep_resultados(rel, ep, completo, limite_id, limite_z):
    if ep.empty:
        rel.texto("(sem rodadas de EP do teste no período)", tamanho=8.5, cor="#777777")
        return
    d = ep.sort_values(["_data", "Provedor", "Programa", "Posição"])
    tab = pd.DataFrame({
        "Mês": d["_data"].map(lambda x: f"{MESES[x.month]}/{x.year}"), "Provedor": d["Provedor"],
        "Programa": d["Programa"], "Equipamento": d["Equipamento"], "Amostra": d["Especime"],
        "Resultado": d["RL"].map(_num), "Valor designado": d["VD"].map(_num),
        "Bias %": d["Bias % amostra"].map(_num), "ID": d["Índice Provedor"].map(_num), "Z grupo": d["IZ"].map(_num),
    })
    cores = pd.DataFrame("", index=tab.index, columns=tab.columns)
    for coluna, origem, lim in (("ID", "Índice Provedor", limite_id), ("Z grupo", "IZ", limite_z)):
        ok = d[origem].notna()
        cores.loc[ok, coluna] = [VERMELHO if abs(v) > lim else VERDE for v in d.loc[ok, origem]]
    if completo:
        tab["Sigma"] = d["Sigma EP"].map(_num)
        cores["Sigma"] = ""
    rel.tabela(tab.reset_index(drop=True), cores.reset_index(drop=True))
    rel.texto(f"ID: verde |ID| ≤ {limite_id:g}, vermelho fora da faixa aceita pelo provedor. Z grupo = (resultado − média "
              f"do grupo) ÷ DP do grupo: verde |Z| ≤ {limite_z:g}.", tamanho=7, cor="#777777")


def _grafico_ep(ep, meses, coluna, titulo, faixa=None, limite_vermelho=None, eixo_minimo=4.0):
    pos = {m: i for i, m in enumerate(meses)}

    def desenha(ax):
        d = ep.dropna(subset=[coluna])
        if faixa is not None and pd.notna(faixa):
            ax.axhspan(-faixa, faixa, color="#2E7D32", alpha=0.13, lw=0)
        if limite_vermelho is not None:
            for s in (1, -1):
                ax.axhline(s * limite_vermelho, color="#C62828", ls="--", lw=0.9)
        ax.axhline(0, color="#999999", lw=0.6)
        for prov, g in d.groupby("Provedor"):
            medias = g.groupby("_mes")[coluna].mean().reset_index().sort_values("_mes")
            ax.plot([pos[m] for m in medias["_mes"]], medias[coluna], color="#757575",
                    ls="-" if prov == "ControlLab" else ":", lw=1.2)
        for (prov, posicao), g in d.groupby(["Provedor", "Posição"]):
            ax.scatter([pos[m] for m in g["_mes"]], g[coluna], s=26, zorder=3,
                       color=CORES_AMOSTRA[(int(posicao) - 1) % len(CORES_AMOSTRA)],
                       marker="D" if prov == "ControlLab" else "o", edgecolors="white", linewidths=0.4)
        if eixo_minimo is not None:
            maior = max([eixo_minimo] + [abs(v) * 1.1 for v in d[coluna]])
            ax.set_ylim(-maior, maior)
        ax.set_title(_txt(titulo), fontsize=9)
        _eixo_meses(ax, [f"{MESES[m.month]}/{m.year}" for m in meses])

    return desenha


def _prepara(ciq, ep):
    """CIQ com um registro por mês/equipamento/nível e EP com a coluna do mês (posição no eixo)."""
    ciq = _um_por_mes(ciq)
    meses = ciq[["_ordem_tempo", "Mês/Ano"]].drop_duplicates().sort_values("_ordem_tempo")
    if not ep.empty:
        ep = ep.assign(_mes=ep["_data"].map(lambda x: pd.Timestamp(x.year, x.month, 1)))
    meses_ep = sorted(set(ep["_mes"])) if not ep.empty else []
    return ciq, meses, ep, meses_ep


def _desenhos(secao, ciq, meses, ep, meses_ep, spec, limite_id, limite_z):
    """Gráficos de uma parte (funções f(ax), lado a lado). Lista vazia = sem gráfico ou sem dados."""
    if secao in GRAFICOS_CIQ:
        coluna, limite, simetrico, maior_melhor, rotulo = GRAFICOS_CIQ[secao]
        if ciq.empty:
            return []
        # só os níveis com valor no período (a especificação pode prever um nível que não tem controle)
        niveis = sorted(int(n) for n in ciq.loc[ciq[coluna].notna(), "NívelNum"].dropna().unique())
        return [_grafico_ciq(ciq[ciq["NívelNum"] == n], meses, coluna, limite, f"{rotulo} — Nível {n}",
                             simetrico, maior_melhor) for n in niveis[:4]]
    if secao == "ep_historico" and not ep.empty:
        return [_grafico_ep(ep, meses_ep, "Índice Provedor", "Índice de Desvio (ID)", faixa=limite_id),
                _grafico_ep(ep, meses_ep, "IZ", "Índice Z (grupo de comparação)", faixa=limite_z,
                            limite_vermelho=3.0),
                _grafico_ep(ep, meses_ep, "Bias % amostra", "Viés (%) por amostra",
                            faixa=(spec or {}).get("ESM (%)"), eixo_minimo=None)]
    return []


def figura_png(secao, ciq, ep, spec, limite_id=1.0, limite_z=2.0, dpi=110):
    """PNG com os gráficos da parte — o mesmo desenho que vai para o PDF — ou None quando não há gráfico."""
    ciq, meses, ep, meses_ep = _prepara(ciq, ep)
    desenhos = _desenhos(secao, ciq, meses, ep, meses_ep, spec, limite_id, limite_z)
    if not desenhos:
        return None
    fig = plt.figure(figsize=(11.0, 3.3))
    for i, desenha in enumerate(desenhos):
        desenha(fig.add_subplot(1, len(desenhos), i + 1))
    fig.subplots_adjust(wspace=0.25)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")   # "tight" inclui a legenda abaixo do gráfico
    plt.close(fig)
    return buf.getvalue()


def gera_pdf(teste, analito, mes, periodo, spec, ciq, ordem_mes, ep, secoes, textos, conclusao,
             responsavel, data_analise, completo=True, limite_id=1.0, limite_z=2.0):
    """
    teste/analito/mes/periodo: textos do cabeçalho; spec: especificação da tabela_mestre (dict);
    ciq: linhas do CIQ do teste no período (até o mês de referência); ordem_mes: ano × 12 + mês de referência;
    ep: resultados do EP do teste (colunas da aba EP); secoes: chaves de SECOES, na ordem;
    textos: {seção: análise crítica}. Devolve os bytes do PDF.
    """
    rel = _Relatorio(f"Desempenho Analítico — análise crítica — {teste} — {mes}")
    rel.nova_pagina()
    rel.texto("Análise crítica do desempenho analítico", tamanho=17, negrito=True, cor=COR)
    rel.espaco(0.006)
    rel.texto(f"Teste: {teste}" + (f" — {analito}" if analito else "") + f"     Mês de referência: {mes}     "
              f"Gráficos: {periodo}", tamanho=10.5, negrito=True)
    spec = spec or {}
    partes = []
    if pd.notna(spec.get("ETM (%)")):
        partes.append(f"ETM {_g(spec['ETM (%)'])}%")
    if pd.notna(spec.get("ETM Absoluto - Valor")) and pd.notna(spec.get("ETM Absoluto - Cutoff")):
        partes.append(f"ETM absoluto {_g(spec['ETM Absoluto - Valor'])} até {_g(spec['ETM Absoluto - Cutoff'])}")
    if pd.notna(spec.get("ESM (%)")):
        partes.append(f"ESM {_g(spec['ESM (%)'])}%")
    cvs = [f"N{i} {_g(spec[f'CV{i}'])}%" for i in range(1, 5) if pd.notna(spec.get(f"CV{i}"))]
    if cvs:
        partes.append("CV máximo " + ", ".join(cvs))
    if completo and pd.notna(spec.get("Sigma Mínimo")):
        partes.append(f"Sigma mínimo {_g(spec['Sigma Mínimo'])}")
    if partes:
        rel.texto("Especificações: " + " · ".join(partes), tamanho=9)
    rel.texto(f"Gerado em {datetime.now():%d/%m/%Y %H:%M} pelo app Desempenho Analítico.", tamanho=8, cor="#777777")

    ciq, meses, ep, meses_ep = _prepara(ciq, ep)
    for secao in secoes:
        tem_grafico = secao in GRAFICOS_CIQ or secao == "ep_historico"
        rel.titulo_secao(SECOES[secao][0], reserva=0.52 if tem_grafico else 0.16)
        if secao == "resumo_ciq":
            _resumo_ciq(rel, ciq[ciq["_ordem_tempo"] == ordem_mes], completo)
        elif secao == "ep_resultados":
            _ep_resultados(rel, ep, completo, limite_id, limite_z)
        else:
            desenhos = _desenhos(secao, ciq, meses, ep, meses_ep, spec, limite_id, limite_z)
            if not desenhos:
                rel.texto("(sem rodadas de EP do teste no período)" if secao == "ep_historico"
                          else "(sem dados no período)", tamanho=8.5, cor="#777777")
            elif secao == "ep_historico":
                rel.graficos(desenhos)
                rel.texto(LEGENDA_EP, tamanho=7, cor="#777777")
            else:
                rel.graficos(desenhos, altura=0.44, margem_inferior=0.16)
                rel.texto(LEGENDA_CIQ, tamanho=7, cor="#777777")
        rel.quadro_analise(textos.get(secao, ""))

    rel.titulo_secao("Conclusão e ações", reserva=0.26)   # conclusão e assinatura na mesma página
    rel.quadro_analise(conclusao, rotulo="Conclusão")
    rel.espaco(0.025)
    rel.texto(f"Responsável: {responsavel or '_' * 40}          Data: "
              f"{data_analise:%d/%m/%Y}" + "          Assinatura: " + "_" * 40, tamanho=9.5)
    return rel.fecha()
