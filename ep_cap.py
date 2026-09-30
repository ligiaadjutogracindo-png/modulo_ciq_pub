"""
Leitura dos relatórios "Original Evaluation" do CAP (Ensaio de Proficiência).

Usa o pdftotext (xpdf, vem junto com o Git for Windows) no modo -table, que mantém
cada espécime na mesma linha dos seus valores. O modo -layout desalinha o nome do
espécime dos valores em alguns relatórios e gravaria o resultado no espécime errado.

PDFs gerados pelo "Actuate" declaram um CMap de codificação que o pdftotext não
conhece e saem com o texto deslocado (e sem dígitos). O texto dentro deles já está
em Unicode, então basta trocar a codificação da fonte por Identity-H — sem OCR.
"""
import re
import subprocess
import tempfile
import zlib
from datetime import datetime
from pathlib import Path

MESES_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]

_ACTUATE_FONTE = re.compile(rb"/Encoding \d+ 0 R /Subtype/Type0/ToUnicode/Identity-H")


def _objetos(pdf):
    return {int(m.group(1)): m.start() for m in re.finditer(rb"(?<![\d])(\d+) 0 obj", pdf)}


def _corpo(pdf, pos):
    ini = pdf.index(b"obj", pos) + 3
    return pdf[ini:pdf.index(b"endobj", ini)].strip()


def _cmap_codigo_cid(corpo):
    """CMap de codificação (stream) → {código Unicode: CID}."""
    s = corpo.index(b"stream") + 6
    while corpo[s:s + 1] in b"\r\n":
        s += 1
    txt = zlib.decompress(corpo[s:corpo.rindex(b"endstream")]).decode("latin-1")
    mapa = {}
    for bloco in re.findall(r"begincidrange(.*?)endcidrange", txt, re.S):
        for a, b, cid in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(\d+)", bloco):
            for k, cod in enumerate(range(int(a, 16), int(b, 16) + 1)):
                mapa[cod] = int(cid) + k
    for bloco in re.findall(r"begincidchar(.*?)endcidchar", txt, re.S):
        for a, cid in re.findall(r"<([0-9A-Fa-f]+)>\s*(\d+)", bloco):
            mapa[int(a, 16)] = int(cid)
    return mapa


def _larguras(array_w):
    """Array /W do PDF ("3[250 260] 21 28 469 ...") → {CID: largura}."""
    toks = re.findall(rb"\[|\]|-?\d+(?:\.\d+)?", array_w)
    larg, i = {}, 0
    toks = toks[1:-1] if toks and toks[0] == b"[" else toks
    while i < len(toks):
        c = int(toks[i])
        if toks[i + 1] == b"[":
            j = i + 2
            while toks[j] != b"]":
                larg[c] = float(toks[j])
                c += 1
                j += 1
            i = j + 1
        else:
            for cid in range(c, int(toks[i + 1]) + 1):
                larg[cid] = float(toks[i + 2])
            i += 3
    return larg


def corrige_actuate(pdf_bytes):
    """PDF "Actuate": o texto está em Unicode, mas a fonte declara um CMap que o pdftotext
    não conhece. Grava (em cópia, na memória) uma atualização incremental com a fonte em
    Identity-H e a tabela de larguras reindexada pelos códigos Unicode — sem as larguras
    certas o pdftotext junta colunas ("0.28130" em vez de "0.281  30")."""
    if b"Registry(Actuate)" not in pdf_bytes or not _ACTUATE_FONTE.search(pdf_bytes):
        return pdf_bytes, False
    try:
        objs = _objetos(pdf_bytes)
        novos = {}
        for n, pos in objs.items():
            corpo = _corpo(pdf_bytes, pos)
            if not _ACTUATE_FONTE.search(corpo):
                continue
            enc = int(re.search(rb"/Encoding (\d+) 0 R", corpo).group(1))
            desc = int(re.search(rb"/DescendantFonts\s*\[\s*(\d+) 0 R", corpo).group(1))
            corpo_desc = _corpo(pdf_bytes, objs[desc])
            ref_w = re.search(rb"/W (\d+) 0 R", corpo_desc)
            codigo_cid = _cmap_codigo_cid(_corpo(pdf_bytes, objs[enc]))
            novos[n] = re.sub(rb"/Encoding \d+ 0 R\s*", b"/Encoding/Identity-H",
                              corpo.replace(b"/ToUnicode/Identity-H", b""))
            if ref_w:
                larg = _larguras(_corpo(pdf_bytes, objs[int(ref_w.group(1))]))
                pares = sorted((cod, larg[cid]) for cod, cid in codigo_cid.items() if cid in larg)
                novos[int(ref_w.group(1))] = b"[" + b" ".join(
                    b"%d[%s]" % (cod, (b"%g" % w)) for cod, w in pares) + b"]"
        if not novos:
            raise ValueError("fonte Actuate não encontrada")
        trailer = pdf_bytes[pdf_bytes.rindex(b"trailer"):]
        root = re.search(rb"/Root\s+(\d+ \d+ R)", trailer).group(1)
        info = re.search(rb"/Info\s+(\d+ \d+ R)", trailer)
        tamanho = max(int(re.search(rb"/Size\s+(\d+)", trailer).group(1)), max(objs) + 1)
        prev = int(re.search(rb"startxref\s+(\d+)", trailer).group(1))

        saida = bytearray(pdf_bytes if pdf_bytes.endswith(b"\n") else pdf_bytes + b"\n")
        offsets = {}
        for n, corpo in sorted(novos.items()):
            offsets[n] = len(saida)
            saida += b"%d 0 obj\n%s\nendobj\n" % (n, corpo)
        inicio_xref = len(saida)
        saida += b"xref\n"
        for n in sorted(offsets):
            saida += b"%d 1\n%010d 00000 n \n" % (n, offsets[n])
        saida += b"trailer\n<</Size %d /Root %s%s /Prev %d>>\nstartxref\n%d\n%%%%EOF\n" % (
            tamanho, root, (b" /Info " + info.group(1)) if info else b"", prev, inicio_xref)
        return bytes(saida), True
    except Exception:
        # formato inesperado: pelo menos troca a codificação (texto sai, colunas podem colar)
        novo = _ACTUATE_FONTE.sub(
            lambda m: b"/Encoding/Identity-H/Subtype/Type0".ljust(len(m.group(0))), pdf_bytes)
        return novo, True


def extrai_texto(pdf_bytes, pdftotext="pdftotext"):
    """Texto do PDF em modo -table (UTF-8). Recebe os bytes do arquivo."""
    pdf_bytes, _ = corrige_actuate(pdf_bytes)
    with tempfile.TemporaryDirectory() as tmp:
        entrada = Path(tmp) / "rel.pdf"
        saida = Path(tmp) / "rel.txt"
        entrada.write_bytes(pdf_bytes)
        subprocess.run([pdftotext, "-table", "-enc", "UTF-8", "-q", str(entrada), str(saida)],
                       capture_output=True, check=False)
        return saida.read_text(encoding="utf-8", errors="replace") if saida.exists() else ""


# ------------------------------------------------------------------
# Cabeçalho
# ------------------------------------------------------------------
_DATA = re.compile(r"\d{1,2}/\d{1,2}/\d{4}")
_PROGRAMA = re.compile(
    r"E\s*V\s*A\s*L\s*U\s*A\s*T\s*I\s*O\s*N[ \t]{2,}([A-Z0-9]{1,8}(?:-[A-Z0-9]{1,3})?)[ \t]+(\d{4})[ \t]+(.+)")


def data_br(texto_data):
    """'6/3/2019' (formato americano M/D/AAAA) → date(2019, 6, 3)."""
    return datetime.strptime(texto_data, "%m/%d/%Y").date()


def data_extenso(d):
    """date → '03/jun/2019' (evita confusão entre dia e mês)."""
    return f"{d.day:02d}/{MESES_PT[d.month - 1]}/{d.year}" if d else None


def _data_do_rotulo(linhas, rotulo):
    """Data ao lado do rótulo ("Kit Mailed: 3/11/2019") ou logo abaixo dele, na mesma coluna."""
    rot = re.compile(r"\s+".join(rotulo.split()) + r"\s*:?")
    for i, linha in enumerate(linhas):
        m = rot.search(linha)
        if not m:
            continue
        mesma = _DATA.match(linha[m.end():].strip())
        if mesma:
            return mesma.group(0)
        for prox in linhas[i + 1:i + 4]:
            datas = [(abs(d.start() - m.start()), d.group(0)) for d in _DATA.finditer(prox)]
            if datas:
                return min(datas)[1]
    return None


def ler_cabecalho(texto):
    linhas = texto.split("\n")
    m = _PROGRAMA.search(texto)
    kit_mailed = _data_do_rotulo(linhas, "Kit Mailed")
    avaliacao = _data_do_rotulo(linhas, "Original Evaluation")
    kit_id = re.search(r"Kit\s*(?:ID|#)\s*:\s*(\d{6,})", texto)
    if not kit_id:
        # página 1 do formato novo: "Kit #:" numa linha e o número embaixo
        kit_id = re.search(r"Kit\s*#\s*:[^\n]*\n\s*(\d{6,})", texto)
    return {
        "Programa": m.group(1) if m else None,
        "Ano Programa": int(m.group(2)) if m else None,
        "Nome Programa": re.sub(r"\s+", " ", m.group(3)).strip() if m else None,
        "Kit ID": kit_id.group(1) if kit_id else None,
        "Data Envio": data_br(kit_mailed) if kit_mailed else None,
        "Data Avaliação": data_br(avaliacao) if avaliacao else None,
    }


# ------------------------------------------------------------------
# Resultados quantitativos (relatórios com S.D.I)
# ------------------------------------------------------------------
_ESPECIME = re.compile(r"(?<!\S)([A-Z][A-Z0-9]{0,7})-(\d{1,3})(?!\S)")
_NUM = re.compile(r"^[<>]?=?[+-]?\d+(?:\.\d+)?$")
_INTEIRO = re.compile(r"^\d+$")
_NOTAS = ("Acceptable", "Unacceptable", "Not Graded", "Not Evaluated", "See Note")
_RODAPE = "The College of American Pathologists recommends"
_UNIDADE = re.compile(r"(/|%|^x ?10|^(Seconds|sec|ratio|Ratio|INR|Index|fL|pg|g|Units|U|mm|mmHg|mIU|IU)\b)")
_CABECALHO_TABELA = re.compile(r"Specimen|Peer Group|Unit of Measure|Your Result|Evaluation|Limits of|^Methods?$")
# cabeçalho/rodapé de página e notas: nunca são rótulo de teste
_LIXO_PAGINA = re.compile(
    r"^ORIGINAL\b|E\s*V\s*A\s*L\s*U\s*A\s*T\s*I?\s*O\s*N|CAP\s+Number|Institution\s*:|Attention\s*:|"
    r"City\s*/\s*State|Kit\s+Mailed|Original\s+Evaluation|Reviewed\s+By|LEGEND|COPIED\s+TO|"
    r"KIT\s+INFORMATION|Exception\s+Reason|^x:\s|^P\s*=\s*Based|^\d{4}$")


def _eh_grupo_comparacao(rotulo):
    """Rótulos que parecem teste novo mas são outro grupo de comparação do mesmo teste."""
    r = rotulo.lower()
    return (r.startswith("all ")                       # All Methods, All Unextracted Methods...
            or r.startswith(("method/", "unextracted method"))  # Method/Instrument/Result...
            or r in {"method mean", "instrument mean", "peer group mean", "group mean"})


def _num(tok):
    return float(tok.lstrip("<>=").lstrip("+"))


def _parse_valores(tokens):
    """Tokens depois do espécime → resultado (RL), média (VD), DP, nº labs, SDI, limites, nota.

    Nem toda linha traz o resultado do laboratório: em alguns programas o grupo do
    equipamento mostra só Média/DP/Labs/SDI e o resultado aparece no bloco "All Methods".
    Para distinguir, usa o nº de labs (sempre inteiro), que vem logo depois do DP:
    posição 3 → há resultado; posição 2 → linha só de estatística do grupo.
    """
    nums, i = [], 0
    while i < len(tokens):
        tok = tokens[i]
        if _NUM.match(tok):
            nums.append(tok)
        elif len(nums) == 1 and re.fullmatch(r"[A-Z]", tok):
            pass  # letra de método logo após o resultado (formato alergia: "0.88 P")
        else:
            break
        i += 1
    if len(nums) < 3:
        return None
    resto = " ".join(tokens[i:])
    nota = next((n for n in _NOTAS if resto.startswith(n)), None)

    k = next((j for j in range(2, len(nums)) if _INTEIRO.match(nums[j])
              and (j + 1 == len(nums) or "." in nums[j + 1] or nums[j + 1][0] in "+-")), None)
    if k == 3:
        qual = re.match(r"^[<>]=?", nums[0])
        rl, qualificador, estat = _num(nums[0]), (qual.group(0) if qual else ""), nums[1:]
    elif k == 2:
        rl, qualificador, estat = None, "", nums      # grupo sem o resultado do laboratório
    else:
        return None
    campos = {"Qualificador": qualificador, "RL": rl, "VD": _num(estat[0]),
              "DP Grupo": _num(estat[1]), "N Labs": int(estat[2]),
              "SDI": None, "Lim Inf": None, "Lim Sup": None, "Nota": nota}
    extra = estat[3:]
    if len(extra) >= 1 and ("." in extra[0] or extra[0][0] in "+-"):
        campos["SDI"] = _num(extra[0])
        extra = extra[1:]
    if len(extra) >= 2:
        campos["Lim Inf"], campos["Lim Sup"] = _num(extra[0]), _num(extra[1])
    return campos


def _acha_especime(linha):
    """Código do espécime na linha. Nome de teste também pode ter esse formato ("IGF-1"),
    então vale o código seguido de valores; sem valores, só um código já na coluna do
    espécime (longe da margem, onde ficam os rótulos)."""
    candidatos = list(_ESPECIME.finditer(linha))
    for m in candidatos:
        valores = _parse_valores(linha[m.end():].split())
        if valores:
            return m, valores
    candidatos = [m for m in candidatos if m.start() >= 15]
    return (candidatos[-1], None) if candidatos else (None, None)


def _limpa_unidade(u):
    """Tira o código do programa que às vezes gruda na unidade ("% HG-B 2025")."""
    return re.sub(r"\s+[A-Z0-9]{1,8}-[A-Z]\s+\d{4}.*$", "", u).strip() if u else u


def ler_resultados(texto):
    """Uma linha por espécime quantitativo, com teste/unidade/método/grupo do bloco."""
    if "S.D.I" not in texto:
        return []
    resultados = []
    bloco = []            # linhas de resultado do teste atual
    rotulos = []          # textos da coluna da esquerda do teste atual, em ordem
    pendentes = []        # rótulos em linhas sem espécime, antes do próximo resultado
    ultimo = None         # (prefixo, número) do último espécime
    formato_b = False     # alergia: "Teste  unidade" numa linha, método na linha do espécime
    teste_b = unidade_b = None

    def fecha_bloco():
        if not bloco:
            return
        rot = [r for r in rotulos if r]
        teste = rot[0] if rot else None
        extra = rot[1:]
        unidade = None
        if extra and _UNIDADE.search(extra[0]):
            unidade = extra.pop(0)
        elif len(extra) >= 2 and _UNIDADE.search(extra[1]) and re.match(r"\(|.*[a-z]", extra[0]):
            # nome do teste quebrado em duas linhas ("Alanine aminotransferase" / "(ALT/SGPT)");
            # linha só em maiúsculas é equipamento/método, não continuação do nome
            teste = f"{teste} {extra.pop(0)}"
            unidade = extra.pop(0)
        metodo = extra[0] if len(extra) >= 2 else None
        grupo = extra[-1] if extra else None
        for r in bloco:
            r.update({"Teste CAP": teste, "Unidade": _limpa_unidade(unidade), "Método": metodo,
                      "Grupo": r.pop("_grupo_proprio", None) or grupo})
        # linhas sem valores (ex.: "<12 ... [28]") só servem pra manter os rótulos no lugar
        resultados.extend(r for r in bloco if not r.pop("_sem_valores", False))

    cabecalho_b_acima = False   # a linha "Test -- Unit Of Measure" veio logo antes
    for linha in texto.split("\n"):
        if not linha.strip():
            continue
        if "Test" in linha and "Unit Of Measure" in linha:
            fecha_bloco()
            bloco, rotulos, pendentes, ultimo = [], [], [], None
            formato_b, cabecalho_b_acima = True, True
            continue
        if "Specimen" in linha and ("Result" in linha or "Mean" in linha):
            # cabeçalho de tabela: só continua no formato B se veio logo após "Unit Of Measure"
            # (o mesmo relatório pode ter tabelas dos dois formatos)
            if not cabecalho_b_acima:
                formato_b = False
            cabecalho_b_acima = False
            continue
        cabecalho_b_acima = False
        if _RODAPE in linha or re.search(r"\bPage\s+\d+\s+of\s+\d+", linha):
            continue
        if _LIXO_PAGINA.search(linha.strip()):
            continue
        m, valores = _acha_especime(linha)
        esquerda = re.sub(r"\s+", " ", linha[:m.start()] if m else linha).strip()

        if formato_b:
            if valores is None:
                # "Common Ragweed      IU/mL" (ou só "ANA, titer"), na coluna 0 → teste e
                # unidade do(s) próximo(s) espécime(s); linhas recuadas são método/equipamento.
                # Sem unidade, só aceita nome com minúscula (tudo maiúsculo é equipamento).
                partes = re.split(r"\s{2,}", linha.strip())
                if not m and not linha[:1].isspace() and not _CABECALHO_TABELA.search(partes[0]):
                    tem_unidade = len(partes) == 2 and _UNIDADE.search(partes[1])
                    if tem_unidade or re.search(r"[a-z]", partes[0]):
                        teste_b = partes[0]
                        unidade_b = _limpa_unidade(partes[1]) if tem_unidade else None
                continue
            resultados.append({"Especime": m.group(0), "Prefixo": m.group(1), "Num": int(m.group(2)),
                               "Teste CAP": teste_b, "Unidade": unidade_b, "Método": None,
                               "Grupo": esquerda or None, **valores})
            continue

        if valores is None and not m:
            # linha sem espécime: rótulo solto (fim do bloco atual ou começo do próximo)
            if esquerda and len(esquerda) < 60 and not _CABECALHO_TABELA.search(esquerda):
                pendentes.append(esquerda)
            continue
        if valores is None:
            # espécime sem valores numéricos (abaixo do limite, não realizado...): não vira
            # resultado, mas participa do bloco pra não deslocar os rótulos do teste
            valores = {"_sem_valores": True}

        chave = (m.group(1), int(m.group(2)))
        novo_bloco = ultimo is None or chave[0] != ultimo[0] or chave[1] <= ultimo[1]
        if novo_bloco and bloco and _eh_grupo_comparacao(esquerda):
            # mesmo teste, outro grupo de comparação ("All Methods"): herda teste e unidade
            rotulos.extend(pendentes)
            base = [r for r in rotulos if r][:2]
            fecha_bloco()
            bloco, rotulos, pendentes = [], base, []
            linha_rot = None
            grupo_proprio = esquerda
        elif novo_bloco:
            if esquerda:
                # a 1ª linha do teste novo já traz o nome: os rótulos soltos são o fim do
                # bloco anterior (método/equipamento de um teste com poucos espécimes)
                rotulos.extend(pendentes)
                pendentes = []
            fecha_bloco()
            bloco, rotulos, pendentes = [], list(pendentes), []
            linha_rot, grupo_proprio = esquerda, None
        else:
            rotulos.extend(pendentes)
            pendentes = []
            linha_rot = esquerda
            grupo_proprio = bloco[-1].get("_grupo_proprio") if bloco else None
        if linha_rot is not None:
            rotulos.append(linha_rot)
        registro = {"Especime": m.group(0), "Prefixo": m.group(1), "Num": int(m.group(2)), **valores}
        if grupo_proprio:
            registro["_grupo_proprio"] = grupo_proprio
        bloco.append(registro)
        ultimo = chave

    fecha_bloco()
    return resultados


def ler_pdf(pdf_bytes, nome_arquivo="", pdftotext="pdftotext"):
    """Cabeçalho + resultados de um PDF. Cada resultado já vem com os dados do cabeçalho."""
    texto = extrai_texto(pdf_bytes, pdftotext=pdftotext)
    cab = ler_cabecalho(texto)
    linhas = ler_resultados(texto)
    for r in linhas:
        r.update(cab)
        r["Arquivo"] = nome_arquivo
    return cab, linhas, texto
