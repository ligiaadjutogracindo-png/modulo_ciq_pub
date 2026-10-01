"""
Base consolidada do EP (Ensaio de Proficiência): leitura do ControlLab, gravação e
controle dos arquivos já processados.

É a mesma rotina usada pelo app (botão "Atualizar base de EP" e importação automática
das pastas configuradas) e por um robô agendado — para rodar sem o app:

    python ep_base.py                      varre as pastas configuradas (dados_ep/config_ep.json)
    python ep_base.py "pasta ou .zip" ...  importa esses caminhos

Arquivos gravados em dados_ep/ (ao lado deste script):
    ep_resultados.csv   uma linha por amostra (CAP e ControlLab), já sem duplicatas
    ep_arquivos.csv     hash de cada arquivo processado (reprocessar não duplica)
    ep_origens.csv      arquivos das pastas já vistos (caminho + tamanho + data), pra varredura rápida
    ep_rodadas.csv      equipamento, viés e fórmula escolhidos por rodada (vem do app)
    config_ep.json      pastas de entrada e se a importação automática está ligada
    ep_log.txt          histórico das atualizações
"""
import hashlib
import io
import json
import re
import shutil
import subprocess
import sys
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path

import pandas as pd

import ep_cap

PASTA_DADOS = Path(__file__).parent / "dados_ep"
ARQ_RESULTADOS = PASTA_DADOS / "ep_resultados.csv"
ARQ_ARQUIVOS = PASTA_DADOS / "ep_arquivos.csv"
ARQ_RODADAS = PASTA_DADOS / "ep_rodadas.csv"
ARQ_ORIGENS = PASTA_DADOS / "ep_origens.csv"
ARQ_CONFIG = PASTA_DADOS / "config_ep.json"
ARQ_LOG = PASTA_DADOS / "ep_log.txt"
ARQ_PACOTE = PASTA_DADOS / "pacote_ep.zip"   # refeito a cada atualização — é o que vai para o app online
PASTA_ORIGINAIS = PASTA_DADOS / "originais"

# Laboratório (participante) do ControlLab considerado — cada regional tem um número.
PARTICIPANTES_CONTROLLAB = {"779"}   # 779 = Brasília (mesmo laboratório dos relatórios do CAP)

COLUNAS = ["Provedor", "Programa", "Rodada", "Data Envio", "Teste Provedor", "Unidade", "Grupo", "Sistema",
           "Equipamento Provedor", "ID Equipamento", "Especime", "Num", "Qualificador", "RL", "VD",
           "DP Grupo", "N Labs", "SDI", "Nota", "Data Avaliação", "Arquivo", "Lim Inf", "Lim Sup"]
CHAVE = ["Provedor", "Programa", "Rodada", "Teste Provedor", "Unidade", "Grupo", "Sistema", "Especime"]
TEXTO = ["Provedor", "Programa", "Rodada", "Teste Provedor", "Unidade", "Grupo", "Sistema",
         "Equipamento Provedor", "ID Equipamento", "Especime", "Qualificador", "Nota", "Arquivo"]
MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
         "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}


# ------------------------------------------------------------------
# ControlLab
# ------------------------------------------------------------------
def _num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return None
    s = str(v).strip()
    s = s.replace(",", "") if "." in s else s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _data_envio_controllab(nome_envio, data_ava):
    """'Abr/2026' → 1º/abr/2026. Rodada especial ('Esp/2022') usa o mês da avaliação."""
    m = re.match(r"([A-Za-zçÇ]{3})\w*/(\d{4})", str(nome_envio or ""))
    if m and m.group(1).lower() in MESES:
        return date(int(m.group(2)), MESES[m.group(1).lower()], 1)
    try:
        d = datetime.strptime(str(data_ava)[:10], "%Y-%m-%d").date()
        return date(d.year, d.month, 1)
    except ValueError:
        return None


def _bloco_sa(linha, tipo, campo):
    """Os campos S1..S8 descrevem o sistema analítico (fabricante, método, equipamento...)."""
    for k in range(1, 9):
        if linha.get(f"S{k}_SA") == tipo:
            v = linha.get(f"S{k}_{campo}")
            return None if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
    return None


def ler_controllab(csv_bytes, nome_arquivo=""):
    """CSV de avaliações do ControlLab (separador '|', UTF-8) → linhas quantitativas do(s) lab(s)."""
    try:
        texto = csv_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = csv_bytes.decode("latin-1")
    if not texto.strip():
        return []
    df = pd.read_csv(io.StringIO(texto), sep="|", dtype=str, quotechar='"')
    if "PART" not in df.columns:
        return []
    df = df[df["PART"].isin(PARTICIPANTES_CONTROLLAB)]
    linhas = []
    for _, r in df.iterrows():
        rl, vd = _num(r.get("VALOR")), _num(r.get("MEDIA"))
        if rl is None or vd is None:
            continue   # qualitativos, não realizados ou sem média do grupo
        exp = _num(r.get("EXPOENTE")) or 0
        if exp:
            rl, vd = rl * 10 ** exp, vd * 10 ** exp
        sinal = str(r.get("SINAL") or "").strip()
        # faixa aceita ("9.1 a 13.7"): limites de aceitação, na mesma escala do resultado
        faixa = [_num(x) for x in re.split(r"\s+a\s+", str(r.get("RESULTADO_ACEITO") or "").strip())]
        lim_inf, lim_sup = (faixa if len(faixa) == 2 and None not in faixa else (None, None))
        if exp and lim_inf is not None:
            lim_inf, lim_sup = lim_inf * 10 ** exp, lim_sup * 10 ** exp
        linhas.append({
            "Provedor": "ControlLab", "Programa": r.get("MODULO"), "Rodada": r.get("NOME_ENVIO"),
            "Data Envio": _data_envio_controllab(r.get("NOME_ENVIO"), r.get("DATA_AVA")),
            "Teste Provedor": r.get("ANALITO"), "Unidade": r.get("UNIDADE"), "Grupo": r.get("NOME_GRUPO_AVA"),
            "Sistema": r.get("SISTEMA"), "Equipamento Provedor": _bloco_sa(r, "EQU", "NOME_SA"),
            "ID Equipamento": _bloco_sa(r, "ID EQU", "OUTRO_SA"),
            "Especime": r.get("NOME_ITEM"), "Num": int(_num(r.get("NUM_ITEM")) or 0),
            "Qualificador": sinal if sinal in ("<", ">", "<=", ">=") else "",
            "RL": rl, "VD": vd, "DP Grupo": _num(r.get("DP")), "N Labs": _num(r.get("QTD_DADOS")),
            "SDI": _num(r.get("INDICE_DESVIO")), "Nota": r.get("AVA"),
            "Data Avaliação": r.get("DATA_AVA"), "Arquivo": nome_arquivo,
            "Lim Inf": lim_inf, "Lim Sup": lim_sup,
        })
    return linhas


# ------------------------------------------------------------------
# CAP
# ------------------------------------------------------------------
def ler_cap(pdf_bytes, nome_arquivo="", pdftotext="pdftotext"):
    cab, linhas, _ = ep_cap.ler_pdf(pdf_bytes, nome_arquivo, pdftotext=pdftotext)
    saida = []
    for r in linhas:
        if r.get("RL") is None:
            continue   # linha só com a estatística do grupo, sem o resultado do laboratório
        saida.append({
            "Provedor": "CAP", "Programa": r["Programa"], "Rodada": r["Kit ID"], "Data Envio": r["Data Envio"],
            "Teste Provedor": r["Teste CAP"], "Unidade": r["Unidade"], "Grupo": r["Grupo"], "Sistema": "",
            "Equipamento Provedor": None, "ID Equipamento": None, "Especime": r["Especime"], "Num": r["Num"],
            "Qualificador": r["Qualificador"], "RL": r["RL"], "VD": r["VD"], "DP Grupo": r["DP Grupo"],
            "N Labs": r["N Labs"], "SDI": r["SDI"], "Nota": r["Nota"],
            "Data Avaliação": r["Data Avaliação"].isoformat() if r.get("Data Avaliação") else "",
            "Arquivo": nome_arquivo, "Lim Inf": r.get("Lim Inf"), "Lim Sup": r.get("Lim Sup"),
        })
    return saida, cab


# ------------------------------------------------------------------
# Arquivos de entrada
# ------------------------------------------------------------------
def hash_bytes(conteudo):
    return hashlib.sha1(conteudo).hexdigest()


def expande_entradas(entradas):
    """(nome, bytes) de PDFs/CSVs, abrindo .zip (inclusive zip dentro de zip) um arquivo por vez."""
    for nome, conteudo in entradas:
        baixo = nome.lower()
        if baixo.endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
                membros = ((Path(n).name, z.read(n)) for n in z.namelist() if not n.endswith("/"))
                yield from expande_entradas(membros)
        elif baixo.endswith((".pdf", ".csv")):
            yield nome, conteudo


def entradas_de_caminhos(caminhos):
    for c in caminhos:
        p = Path(c)
        arquivos = [p] if p.is_file() else sorted(x for x in p.rglob("*") if x.is_file())
        for a in arquivos:
            if a.suffix.lower() == ".zip":
                with zipfile.ZipFile(a) as z:   # lê direto do disco, sem carregar o zip inteiro
                    for n in z.namelist():
                        if not n.endswith("/"):
                            yield Path(n).name, z.read(n)
            elif a.suffix.lower() in (".pdf", ".csv"):
                yield a.name, a.read_bytes()


def processa_entradas(entradas, hashes_conhecidos=frozenset(), pdftotext="pdftotext", progresso=None,
                      lote=100):
    """Lê os arquivos novos (hash ainda não processado), em lotes pra não segurar tudo na memória.
    Devolve (linhas, registro_arquivos, ignorados). progresso(n_processados) é chamado a cada lote."""
    ignorados = 0
    vistos = set(hashes_conhecidos)

    def um(item):
        nome, conteudo, h = item
        try:
            if nome.lower().endswith(".csv"):
                linhas, info = ler_controllab(conteudo, nome), {"Provedor": "ControlLab"}
            else:
                linhas, cab = ler_cap(conteudo, nome, pdftotext=pdftotext)
                info = {"Provedor": "CAP", "Programa": cab.get("Programa"), "Ano": cab["Data Envio"].year
                        if cab.get("Data Envio") else None}
        except Exception as e:   # arquivo corrompido/inesperado não interrompe os demais
            return [], {"Arquivo": nome, "Hash": h, "Linhas": 0,
                        "Processado em": datetime.now().isoformat(timespec="seconds"),
                        "Erro": f"{type(e).__name__}: {e}"[:200]}, None
        # só guarda os bytes de quem tem resultado (é o que vai para a pasta de originais)
        return linhas, {"Arquivo": nome, "Hash": h, "Linhas": len(linhas),
                        "Processado em": datetime.now().isoformat(timespec="seconds"), **info}, \
            (conteudo if linhas else None)

    linhas, registro, pendentes = [], [], []

    def roda_lote(ex):
        for ls, reg, conteudo in ex.map(um, pendentes):
            linhas.extend(ls)
            registro.append({**reg, "_bytes": conteudo})
        pendentes.clear()
        if progresso:
            progresso(len(registro))

    with ThreadPoolExecutor(max_workers=8) as ex:
        for nome, conteudo in expande_entradas(entradas):
            h = hash_bytes(conteudo)
            if h in vistos:
                ignorados += 1
                continue
            vistos.add(h)
            pendentes.append((nome, conteudo, h))
            if len(pendentes) >= lote:
                roda_lote(ex)
        if pendentes:
            roda_lote(ex)
    return linhas, registro, ignorados


# ------------------------------------------------------------------
# Base consolidada
# ------------------------------------------------------------------
# "NA" é o mnemônico do sódio: o pandas, por padrão, lê "NA" (e "N/A", "NULL"...) como vazio
SO_VAZIO_E_NA = {"keep_default_na": False, "na_values": [""]}


def _le_base(fonte):
    base = pd.read_csv(fonte, dtype={c: str for c in TEXTO}, encoding="utf-8-sig", **SO_VAZIO_E_NA)
    base["Data Envio"] = pd.to_datetime(base["Data Envio"], errors="coerce").dt.date
    for c in TEXTO:
        base[c] = base[c].fillna("") if c in base.columns else ""
    return base


def carrega_base():
    if not ARQ_RESULTADOS.exists():
        return pd.DataFrame(columns=COLUNAS)
    return _le_base(ARQ_RESULTADOS)


def carrega_registro():
    if not ARQ_ARQUIVOS.exists():
        return pd.DataFrame(columns=["Arquivo", "Hash", "Linhas", "Processado em", "Provedor", "Programa", "Ano"])
    return pd.read_csv(ARQ_ARQUIVOS, dtype=str, encoding="utf-8-sig")


def consolida(base, linhas_novas):
    """Junta à base, sem duplicar: para a mesma amostra, fica a avaliação mais recente."""
    novas = pd.DataFrame(linhas_novas, columns=COLUNAS)
    for c in TEXTO:
        novas[c] = novas[c].fillna("").astype(str)
    tudo = pd.concat([base, novas], ignore_index=True) if len(base) else novas
    tudo["Data Avaliação"] = tudo["Data Avaliação"].fillna("").astype(str)
    tudo = tudo.sort_values("Data Avaliação", kind="stable")
    return tudo.drop_duplicates(CHAVE, keep="last").sort_values(
        ["Provedor", "Data Envio", "Programa", "Teste Provedor", "Num"], na_position="last").reset_index(drop=True)


def grava(linhas_novas, registro_novo, arquivar_originais=False):
    """Grava as linhas na base e registra os arquivos. Os originais já ficam nas pastas de entrada (rede e
    APP/EP_entrada); com arquivar_originais=True guarda também uma cópia por ano/programa em dados_ep/originais."""
    PASTA_DADOS.mkdir(exist_ok=True)
    base = consolida(carrega_base(), linhas_novas)
    base.to_csv(ARQ_RESULTADOS, index=False, encoding="utf-8-sig")
    reg = carrega_registro()
    novos = pd.DataFrame([{k: v for k, v in r.items() if k != "_bytes"} for r in registro_novo])
    pd.concat([reg, novos], ignore_index=True).to_csv(ARQ_ARQUIVOS, index=False, encoding="utf-8-sig")
    if arquivar_originais:
        for r in registro_novo:
            if not r.get("_bytes") or not r.get("Linhas"):
                continue   # só guarda o que tem resultado quantitativo
            destino = PASTA_ORIGINAIS / str(r.get("Provedor") or "outros") / str(r.get("Ano") or "") / \
                str(r.get("Programa") or "")
            destino.mkdir(parents=True, exist_ok=True)
            (destino / r["Arquivo"]).write_bytes(r["_bytes"])
    ARQ_PACOTE.write_bytes(monta_pacote(base, carrega_rodadas()))
    return base


COLUNAS_RODADAS = ["Provedor", "Programa", "Rodada", "Mneumonico", "Sistema", "Equipamento", "Modo Viés",
                   "Fórmula Sigma"]


def _le_rodadas(fonte):
    df = pd.read_csv(fonte, dtype=str, encoding="utf-8-sig", **SO_VAZIO_E_NA).fillna("")
    for c in COLUNAS_RODADAS:   # arquivos gravados antes de existir a coluna "Fórmula Sigma"
        if c not in df.columns:
            df[c] = ""
    return df[COLUNAS_RODADAS]


def carrega_rodadas():
    if not ARQ_RODADAS.exists():
        return pd.DataFrame(columns=COLUNAS_RODADAS)
    return _le_rodadas(ARQ_RODADAS)


def grava_rodadas(df):
    PASTA_DADOS.mkdir(exist_ok=True)
    df.to_csv(ARQ_RODADAS, index=False, encoding="utf-8-sig")
    base = carrega_base()
    if len(base):   # no app online não há base local (ela vem do pacote enviado): não refaz o pacote
        ARQ_PACOTE.write_bytes(monta_pacote(base, df))


# ------------------------------------------------------------------
# Pacote de EP (.zip) — leva a base pronta para o app online, que não lê PDF nem guarda arquivos
# ------------------------------------------------------------------
ARQ_DEPARA = Path(__file__).parent / "reference_data" / "ep_depara_testes.xlsx"


def monta_pacote(base, rodadas=None, depara_bytes=None):
    """Zip com a base de EP já processada, as escolhas de equipamento/viés/fórmula e o de-para de
    testes (assim o app online não depende do de-para estar no GitHub)."""
    if depara_bytes is None and ARQ_DEPARA.exists():
        depara_bytes = ARQ_DEPARA.read_bytes()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("ep_resultados.csv", base.to_csv(index=False).encode("utf-8-sig"))
        if rodadas is not None and len(rodadas):
            z.writestr("ep_rodadas.csv", rodadas.to_csv(index=False).encode("utf-8-sig"))
        if depara_bytes:
            z.writestr("ep_depara_testes.xlsx", depara_bytes)
        datas = pd.to_datetime(base["Data Envio"], errors="coerce").dropna()
        periodo = f"rodadas de {datas.min():%m/%Y} a {datas.max():%m/%Y}" if len(datas) else "sem rodadas"
        z.writestr("LEIA-ME.txt", (
            f"Pacote da base de EP gerado em {datetime.now():%d/%m/%Y %H:%M}.\n"
            f"{len(base)} resultado(s) ({', '.join(f'{p}: {n}' for p, n in base['Provedor'].value_counts().items())}), "
            f"{periodo}.\n"
            f"{0 if rodadas is None else len(rodadas)} escolha(s) de equipamento/viés/fórmula.\n"
            f"De-para de testes: {'incluído' if depara_bytes else 'NÃO incluído'}.\n\n"
            "Para usar: no app, barra lateral → \"Base de EP (.zip)\". Não descompacte.\n").encode("utf-8"))
    return buf.getvalue()


def le_pacote(conteudo):
    """Pacote gerado por monta_pacote → (base, rodadas ou None, bytes do de-para ou None)."""
    with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
        nomes = {Path(n).name: n for n in z.namelist()}
        if "ep_resultados.csv" not in nomes:
            raise ValueError("este .zip não é um pacote de EP (falta o ep_resultados.csv) — use o pacote "
                             "gerado na aba EP (\"Baixar pacote de EP\").")
        base = _le_base(io.BytesIO(z.read(nomes["ep_resultados.csv"])))
        rodadas = (_le_rodadas(io.BytesIO(z.read(nomes["ep_rodadas.csv"])))
                   if "ep_rodadas.csv" in nomes else None)
        depara = z.read(nomes["ep_depara_testes.xlsx"]) if "ep_depara_testes.xlsx" in nomes else None
    return base, rodadas, depara


# ------------------------------------------------------------------
# Pastas de entrada (atualização automática)
# ------------------------------------------------------------------
def carrega_config():
    config = {"pastas": [], "automatico": False}
    if ARQ_CONFIG.exists():
        try:
            config.update(json.loads(ARQ_CONFIG.read_text(encoding="utf-8")))
        except (ValueError, OSError):
            pass
    return config


def grava_config(pastas, automatico):
    PASTA_DADOS.mkdir(exist_ok=True)
    ARQ_CONFIG.write_text(json.dumps({"pastas": pastas, "automatico": bool(automatico)},
                                     ensure_ascii=False, indent=2), encoding="utf-8")


def _assinatura(p):
    info = p.stat()
    return f"{p.resolve()}|{info.st_size}|{int(info.st_mtime)}"


def arquivos_pendentes(pastas):
    """Arquivos das pastas ainda não vistos (caminho + tamanho + data) — não precisa reler o que já
    foi lido. Ignora a própria pasta dados_ep (onde ficam os originais arquivados)."""
    vistos = set()
    if ARQ_ORIGENS.exists():
        vistos = set(pd.read_csv(ARQ_ORIGENS, dtype=str, encoding="utf-8-sig")["Assinatura"])
    dados = PASTA_DADOS.resolve()
    pendentes, inexistentes = [], []
    for pasta in pastas:
        raiz = Path(pasta)
        if not raiz.exists():
            inexistentes.append(pasta)
            continue
        for p in ([raiz] if raiz.is_file() else raiz.rglob("*")):
            if not p.is_file() or p.suffix.lower() not in (".pdf", ".csv", ".zip"):
                continue
            if dados == p.resolve() or dados in p.resolve().parents:
                continue
            assinatura = _assinatura(p)
            if assinatura not in vistos:
                pendentes.append((p, assinatura))
    return pendentes, inexistentes


def sincroniza_pastas(pastas, pdftotext="pdftotext", progresso=None):
    """Importa só o que é novo nas pastas e grava direto na base. Devolve um resumo."""
    pendentes, inexistentes = arquivos_pendentes(pastas)
    resumo = {"verificados": len(pendentes), "novos": 0, "ignorados": 0, "linhas": 0,
              "erros": [], "inexistentes": inexistentes}
    if not pendentes:
        return resumo
    falharam = set()

    def entradas():
        for p, _ in pendentes:
            try:
                yield from entradas_de_caminhos([p])
            except Exception as e:   # ex.: arquivo do OneDrive ainda não baixado, zip corrompido
                falharam.add(p)
                resumo["erros"].append(f"{p.name}: {type(e).__name__}: {e}"[:200])

    hashes = set(carrega_registro()["Hash"])
    linhas, registro, ignorados = processa_entradas(entradas(), hashes, pdftotext=pdftotext,
                                                    progresso=progresso)
    if registro:
        grava(linhas, registro)
    lidos = pd.DataFrame([{"Assinatura": a, "Arquivo": p.name,
                           "Lido em": datetime.now().isoformat(timespec="seconds")}
                          for p, a in pendentes if p not in falharam])
    if len(lidos):
        anteriores = (pd.read_csv(ARQ_ORIGENS, dtype=str, encoding="utf-8-sig")
                      if ARQ_ORIGENS.exists() else pd.DataFrame())
        pd.concat([anteriores, lidos], ignore_index=True).to_csv(ARQ_ORIGENS, index=False, encoding="utf-8-sig")
    resumo.update(novos=len(registro), ignorados=ignorados, linhas=len(linhas))
    resumo["erros"] += [f"{r['Arquivo']}: {r['Erro']}" for r in registro if r.get("Erro")]
    registra_log(resumo)
    return resumo


def refaz_provedor(provedor, caminhos, pdftotext="pdftotext", progresso=None):
    """Apaga da base e do registro tudo o que veio do provedor e reimporta os caminhos — para quando
    o leitor melhora (os arquivos já registrados seriam ignorados). Não mexe nas escolhas de
    equipamento/viés/fórmula nem no outro provedor."""
    base = carrega_base()
    reg = carrega_registro()
    PASTA_DADOS.mkdir(exist_ok=True)
    base[base["Provedor"] != provedor].to_csv(ARQ_RESULTADOS, index=False, encoding="utf-8-sig")
    reg[reg.get("Provedor", pd.Series("", index=reg.index)).fillna("") != provedor].to_csv(
        ARQ_ARQUIVOS, index=False, encoding="utf-8-sig")
    hashes = set(carrega_registro()["Hash"])
    linhas, registro, ignorados = processa_entradas(entradas_de_caminhos(caminhos), hashes,
                                                    pdftotext=pdftotext, progresso=progresso)
    grava(linhas, registro)
    resumo = {"verificados": len(registro) + ignorados, "novos": len(registro), "ignorados": ignorados,
              "linhas": len(linhas), "erros": [f"{r['Arquivo']}: {r['Erro']}" for r in registro if r.get("Erro")],
              "inexistentes": []}
    registra_log(resumo)
    return resumo


def registra_log(resumo):
    PASTA_DADOS.mkdir(exist_ok=True)
    with ARQ_LOG.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat(timespec='seconds')}  {resumo['verificados']} arquivo(s) "
                f"verificado(s), {resumo['novos']} novo(s), {resumo['ignorados']} repetido(s), "
                f"{resumo['linhas']} resultado(s), {len(resumo['erros'])} erro(s)\n")
        for e in resumo["erros"]:
            f.write(f"    erro: {e}\n")


def _tem_modo_tabela(exe):
    """O leitor do CAP precisa do pdftotext do xpdf (opção -table). O do poppler (Linux, ex.:
    Streamlit Cloud) não tem essa opção e leria os PDFs em silêncio sem extrair nada."""
    try:
        r = subprocess.run([exe, "-h"], capture_output=True, text=True, timeout=10)
        return "-table" in (r.stdout + r.stderr)
    except (OSError, subprocess.SubprocessError):
        return False


def localiza_pdftotext():
    candidatos = [shutil.which("pdftotext"),
                  str(Path(__file__).parent / "bin" / "pdftotext"),   # binário do xpdf junto do app
                  r"C:\Program Files\Git\mingw64\bin\pdftotext.exe",
                  r"C:\Program Files (x86)\Git\mingw64\bin\pdftotext.exe"]
    for c in candidatos:
        if c and Path(c).exists() and _tem_modo_tabela(c):
            return c
    return None


if __name__ == "__main__":
    # modo robô: sem argumentos varre as pastas configuradas; com argumentos importa esses caminhos
    if len(sys.argv) >= 2 and sys.argv[1] == "--pacote":
        # python ep_base.py --pacote [destino.zip]: gera o pacote da base atual (para o app online)
        destino = Path(sys.argv[2] if len(sys.argv) > 2 else f"pacote_ep_{datetime.now():%Y-%m-%d}.zip")
        destino.write_bytes(monta_pacote(carrega_base(), carrega_rodadas()))
        print(f"Pacote gerado: {destino.resolve()}")
        sys.exit(0)
    exe = localiza_pdftotext()
    if not exe:
        sys.exit("pdftotext não encontrado (vem com o Git for Windows).")
    if len(sys.argv) >= 4 and sys.argv[1] == "--refazer":
        # python ep_base.py --refazer CAP <pastas ou zips>: reimporta o provedor com o leitor atual
        r = refaz_provedor(sys.argv[2], sys.argv[3:], pdftotext=exe)
        print(f"{sys.argv[2]} refeito: {r['novos']} arquivo(s) lido(s), {r['ignorados']} repetido(s), "
              f"{r['linhas']} resultado(s); {len(r['erros'])} erro(s). Base: {len(carrega_base())} linhas.")
        sys.exit(0)
    pastas = sys.argv[1:] or carrega_config()["pastas"]
    if not pastas:
        print(__doc__)
        sys.exit("Nenhuma pasta configurada: informe na aba EP do app ou passe o caminho no comando.")
    r = sincroniza_pastas(pastas, pdftotext=exe)
    print(f"{r['verificados']} arquivo(s) verificado(s), {r['novos']} novo(s), {r['ignorados']} repetido(s), "
          f"{r['linhas']} resultado(s) importado(s).")
    for e in r["erros"]:
        print("  erro:", e)
    for p in r["inexistentes"]:
        print("  pasta não encontrada:", p)
