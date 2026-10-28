"""
Leitura do CIQ exportado do Shift (planilha de CQ mensal por equipamento / analito / lote / nível).

Devolve os registros no MESMO formato do leitor do Infinity (parse_zip, no app.py), para que todo o
resto do app — níveis, CV, Bias, Erro Total, Sigma, gráficos — funcione igual para as duas fontes.

Quais colunas do Shift viram o quê (decidido em 28/out/2026):
    Média observada  ← XMBreakPoint
    CV observado (%) ← CVBreakPoint (= DPBreakPoint ÷ XMBreakPoint × 100; é o "CV Mensal" do Shift)
    Valor-alvo       ← XM Uso
    DP alvo          ← DP Uso (guardado para consulta; ainda não entra em nenhum cálculo)
    N                ← Quantidade de pontos validados
As colunas XMMensal / DPMensal NÃO são usadas: não batem com o CV Mensal do próprio Shift.

Formato do arquivo: CSV com ";" , latin-1 (ou UTF-8), um mês por arquivo (colunas Ano e Mes).
Os números vêm misturados: vírgula decimal em CV/XM/DP do mês ("7,118") e ponto decimal nas colunas
do fabricante e de uso ("3.5"). Não há separador de milhar em nenhuma delas.
"""
import io

import pandas as pd

# Colunas que identificam um export do Shift (para o app reconhecer o arquivo sozinho)
COLUNAS_ASSINATURA = {"XMBreakPoint", "CVBreakPoint", "XM Uso", "Analito", "Nivel"}

MESES_EXTENSO = {"janeiro": 1, "fevereiro": 2, "marco": 3, "março": 3, "abril": 4, "maio": 5,
                 "junho": 6, "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10,
                 "novembro": 11, "dezembro": 12}


def decodifica(conteudo):
    """Bytes do arquivo → texto. O Shift exporta em latin-1; aceita UTF-8 também."""
    try:
        return conteudo.decode("utf-8-sig")
    except UnicodeDecodeError:
        return conteudo.decode("latin-1")


def eh_shift(texto):
    """True quando o cabeçalho (1ª linha não vazia — o Shift começa o arquivo com uma linha em
    branco) tem as colunas típicas do Shift."""
    cabecalho = next((l for l in texto.splitlines() if l.strip()), "")
    colunas = {c.strip().strip('"') for c in cabecalho.split(";")}
    return COLUNAS_ASSINATURA <= colunas


def numero(v):
    """'7,118' → 7.118 · '3.5' → 3.5 · '' → None.
    Com vírgula, a vírgula é o decimal (e um ponto seria milhar); sem vírgula, o ponto é o decimal."""
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "-"):
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def mes_numero(texto):
    """'Setembro' → 9 (também aceita '9' ou '09')."""
    t = str(texto or "").strip().lower()
    if t.isdigit():
        return int(t)
    return MESES_EXTENSO.get(t)


def ler_shift(texto, nome_arquivo=""):
    """Texto do CSV do Shift → (registros, avisos).
    registros: lista de dicts com as mesmas chaves do leitor do Infinity, mais:
        "Fonte CIQ" = "Shift", "DP alvo" (DP Uso) e "Sem BreakPoint" (True quando o Shift
        não trouxe XM/CV do BreakPoint — a linha fica sem análise, só com os dados brutos).
    avisos: lista de (arquivo, motivo) para linhas que não puderam ser lidas."""
    df = pd.read_csv(io.StringIO(texto), sep=";", dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]
    registros, avisos = [], []
    for i, row in df.iterrows():
        teste = row.get("Analito", "").strip()
        if not teste:
            continue
        ano, mes = numero(row.get("Ano")), mes_numero(row.get("Mes"))
        if not ano or not mes:
            avisos.append((nome_arquivo, f"linha {i + 2}: mês/ano não identificado "
                                         f"({row.get('Mes')!r}/{row.get('Ano')!r})"))
            continue
        media, cv = numero(row.get("XMBreakPoint")), numero(row.get("CVBreakPoint"))
        registros.append({
            "Arquivo": nome_arquivo,
            "Fonte CIQ": "Shift",
            "Módulo/Arquivo": row.get("Modulo", "").strip(),
            "Ano": int(ano), "Mês": mes,
            "Equipamento": row.get("Equipamento", "").strip(),   # alguns nomes vêm com espaço no fim
            "Teste": teste,
            "Nível": row.get("Nivel", "").strip(),                 # "Nível 1" → nível 1 pelo número final
            "Número de lote": row.get("Lote", "").strip(),
            "N": numero(row.get("Quantidade de pontos validados")),
            "Config. valor alvo": numero(row.get("XM Uso")),
            "DP alvo": numero(row.get("DP Uso")),
            "Média": media,
            "CV (%)": cv,
            "Sem BreakPoint": media is None or cv is None,
        })
    return registros, avisos
