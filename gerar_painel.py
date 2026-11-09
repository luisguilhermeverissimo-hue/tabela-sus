"""Gera o painel HTML das ações de reajuste da tabela do SUS no TRF-1.

Uso (na pasta CONSOLIDAÇÃO MUNICÍPIOS TABELA SUS):
    py -3 painel/gerar_painel.py

Lê, da pasta principal:
  - cnpj-municipios-<uf>.md   lista de municípios e CNPJs de cada estado do Nordeste
  - reajuste-sus-trf1-<uf>.xlsx   resultado da pesquisa no TRF-1 (abas Resumo, Municípios, Processos)

Um estado sem planilha aparece no painel como "pesquisa pendente". Quando a planilha
de um novo estado for salva na pasta, basta rodar o script de novo.

Usa, desta pasta (painel/):
  - modelo-painel.html          o layout do painel
  - logo-monteiro.svg           logo horizontal do manual de marca
  - malha-ibge-nordeste.json    malha municipal do IBGE (baixada de novo se faltar)

Grava: painel-tabela-sus-nordeste.html na pasta principal.
"""
import datetime as dt
import glob
import json
import math
import os
import re
import sys
import unicodedata
import urllib.request

try:
    import openpyxl
except ImportError:
    sys.exit("Falta o pacote openpyxl. Instale com: py -3 -m pip install openpyxl")

AQUI = os.path.dirname(os.path.abspath(__file__))
PASTA = os.path.dirname(AQUI)
SAIDA = os.path.join(PASTA, "painel-tabela-sus-nordeste.html")

ESTADOS = [  # ordem de exibição: norte para sul
    ("MA", "Maranhão"), ("PI", "Piauí"), ("CE", "Ceará"), ("RN", "Rio Grande do Norte"),
    ("PB", "Paraíba"), ("PE", "Pernambuco"), ("AL", "Alagoas"), ("SE", "Sergipe"), ("BA", "Bahia"),
]
FERNANDO_DE_NORONHA = "2605459"  # distrito estadual: sem CNPJ municipal, fora do mapa


def chave(txt):
    return unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode().lower().strip()


# ---------------------------------------------------------------- municípios (CNPJ)
def ler_lista_cnpj(uf):
    caminho = os.path.join(PASTA, f"cnpj-municipios-{uf.lower()}.md")
    if not os.path.exists(caminho):
        return None
    linhas = []
    for ln in open(caminho, encoding="utf-8"):
        m = re.match(r"\|\s*\d+\s*\|\s*(.+?)\s*\|\s*(\d{7})\s*\|\s*([\d./-]{18})\s*\|", ln)
        if m:
            linhas.append({"nome": m.group(1), "ibge": m.group(2), "cnpj": m.group(3)})
    return linhas


# ---------------------------------------------------------------- planilhas do TRF-1
def situacao(txt):
    t = chave(txt or "")
    if t.startswith("com"):
        return "com"
    if t.startswith("sem"):
        return "sem"
    if t.startswith("verificar"):
        return "verificar"
    return "verificar"  # rótulo desconhecido: pede conferência em vez de esconder


def fase(mov):
    s = chave(mov or "")
    if "suspens" in s or "sobrest" in s:
        return "Suspenso (Tema 1.305 do STJ)"
    if "grau de recurso" in s:
        return "Remetido ao TRF-1 em recurso"
    if "embargos" in s:
        return "Embargos de declaração"
    if "apela" in s or "contrarraz" in s:
        return "Apelação e contrarrazões"
    if s.startswith("conclusos"):
        return "Concluso ao juiz"
    if "arquivado" in s or "cancelada a distribui" in s:
        return "Arquivado ou cancelado"
    if "distribui" in s or "emenda" in s:
        return "Início do processo"
    return "Outros andamentos"


def data_iso(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(v or ""))
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None


def linhas_aba(wb, nome):
    alvo = chave(nome)
    for ws in wb.worksheets:
        if chave(ws.title) == alvo:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return []
            cab = [chave(c or "") for c in rows[0]]
            return [dict(zip(cab, r)) for r in rows[1:] if any(c is not None for c in r)]
    return []


def campo(linha, *nomes):
    for n in nomes:
        n = chave(n)
        for k, v in linha.items():
            if k == n or k.startswith(n):
                return v
    return None


def ler_planilha(uf):
    caminho = os.path.join(PASTA, f"reajuste-sus-trf1-{uf.lower()}.xlsx")
    if not os.path.exists(caminho):
        return None
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)

    resumo = []
    for ws in wb.worksheets:
        if chave(ws.title) == "resumo":
            resumo = [str(r[0]) for r in ws.iter_rows(values_only=True) if r and r[0] is not None]
    data_pesquisa = next((m.group(1) for r in resumo for m in [re.search(r"(\d{2}/\d{2}/\d{4})", str(r))] if m), None)
    i = next((k for k, r in enumerate(resumo) if chave(r) == "metodologia"), None)
    metodologia = [re.sub(r"^\d+\.\s*", "", str(r)) for r in resumo[i + 1:]] if i is not None else []

    municipios = {}
    for r in linhas_aba(wb, "Municípios"):
        ibge = str(campo(r, "Código IBGE") or "").strip()
        if not ibge.isdigit():
            continue
        obs = campo(r, "Observação")
        municipios[ibge] = {
            "sit": situacao(campo(r, "Situação")),
            "obs": obs.strip() if isinstance(obs, str) else "",
            "cnpj": campo(r, "CNPJ"),
            "nome": campo(r, "Município"),
        }

    processos = []
    for r in linhas_aba(wb, "Processos"):
        papel = chave(campo(r, "Papel do município") or "")
        if papel != "autor":
            continue
        processos.append({
            "cnpj": campo(r, "CNPJ"),
            "mun": campo(r, "Município"),
            "num": str(campo(r, "Processo") or "").strip(),
            "vara": str(campo(r, "Vara") or "").strip(),
            "aut": data_iso(campo(r, "Autuação")),
            "mov": str(campo(r, "Último movimento") or "").strip(),
        })
    wb.close()
    return {"data": data_pesquisa, "metodologia": metodologia, "municipios": municipios,
            "processos": processos, "arquivo": os.path.basename(caminho)}


# ---------------------------------------------------------------- mapa
def carregar_malha():
    caminho = os.path.join(AQUI, "malha-ibge-nordeste.json")
    if os.path.exists(caminho):
        return json.load(open(caminho, encoding="utf-8"))
    print("Baixando a malha municipal do IBGE...")
    base = "https://servicodados.ibge.gov.br/api/v3/malhas/estados/{uf}?formato=application/vnd.geo%2Bjson&qualidade=minima{extra}"
    malha = {"municipios": {}, "estados": {}}
    for uf, _ in ESTADOS:
        with urllib.request.urlopen(base.format(uf=uf, extra="&intrarregiao=municipio"), timeout=120) as r:
            for f in json.load(r)["features"]:
                malha["municipios"][f["properties"]["codarea"]] = f["geometry"]
        with urllib.request.urlopen(base.format(uf=uf, extra=""), timeout=120) as r:
            malha["estados"][uf] = json.load(r)["features"][0]["geometry"]
    json.dump(malha, open(caminho, "w", encoding="utf-8"), separators=(",", ":"))
    return malha


def poligonos(geom):
    if geom["type"] == "Polygon":
        return [geom["coordinates"]]
    if geom["type"] == "MultiPolygon":
        return geom["coordinates"]
    return []


def projetar(malha, largura=1000):
    geoms = {k: v for k, v in malha["municipios"].items() if k != FERNANDO_DE_NORONHA}
    pts = [p for g in geoms.values() for poly in poligonos(g) for anel in poly for p in anel]
    lon0, lon1 = min(p[0] for p in pts), max(p[0] for p in pts)
    lat0, lat1 = min(p[1] for p in pts), max(p[1] for p in pts)
    k = math.cos(math.radians((lat0 + lat1) / 2))
    escala = largura / ((lon1 - lon0) * k)
    altura = (lat1 - lat0) * escala

    def caminho(g):
        partes = []
        for poly in poligonos(g):
            for anel in poly:
                xy = [((x - lon0) * k * escala, (lat1 - y) * escala) for x, y in anel]
                partes.append("M" + "L".join(f"{x:.1f},{y:.1f}" for x, y in xy) + "Z")
        return "".join(partes)

    return {
        "w": largura, "h": round(altura, 1),
        "municipios": {c: caminho(g) for c, g in geoms.items()},
        "estados": {uf: caminho(g) for uf, g in malha["estados"].items()},
    }


# ---------------------------------------------------------------- logo
def logo_inline():
    svg = open(os.path.join(AQUI, "logo-monteiro.svg"), encoding="utf-8").read()
    svg = re.sub(r"<\?xml[^>]*\?>", "", svg)
    svg = re.sub(r"<defs>.*?</defs>", "", svg, flags=re.S)
    svg = re.sub(r'\sid="[^"]*"', "", svg)
    svg = svg.replace('class="cls-1"', 'class="logo-texto"').replace('class="cls-2"', 'class="logo-simbolo"')
    svg = svg.replace("<svg ", '<svg class="logo" role="img" aria-label="Monteiro e Monteiro Advogados Associados" ', 1)
    if "<script" in svg.lower():
        sys.exit("O logo contém script; use outro arquivo.")
    return svg.strip()


# ---------------------------------------------------------------- montagem
def main():
    estados, municipios, processos, avisos = [], [], [], []
    for uf, nome in ESTADOS:
        lista = ler_lista_cnpj(uf)
        if lista is None:
            avisos.append(f"{uf}: falta cnpj-municipios-{uf.lower()}.md; estado fora do painel.")
            continue
        pesq = ler_planilha(uf)
        por_cnpj = {}
        if pesq:
            for p in pesq["processos"]:
                por_cnpj.setdefault(p["cnpj"], []).append(p)
        for m in lista:
            r = pesq["municipios"].get(m["ibge"]) if pesq else None
            if pesq and r is None:
                avisos.append(f"{uf}: {m['nome']} ({m['ibge']}) não está na planilha; marcado como pendente.")
            sit = r["sit"] if r else "pendente"
            procs = por_cnpj.pop(m["cnpj"], [])
            if sit == "com" and not procs:
                avisos.append(f"{uf}: {m['nome']} está 'com ação' mas sem processo na aba Processos.")
            municipios.append({
                "ibge": m["ibge"], "nome": m["nome"], "uf": uf, "cnpj": m["cnpj"],
                "sit": sit, "obs": r["obs"] if r else "", "n": len(procs),
            })
            for p in procs:
                processos.append({"ibge": m["ibge"], "uf": uf, "num": p["num"], "vara": p["vara"],
                                  "aut": p["aut"], "mov": p["mov"], "fase": fase(p["mov"])})
        if pesq:
            for cnpj, sobra in por_cnpj.items():
                avisos.append(f"{uf}: {len(sobra)} processo(s) com CNPJ {cnpj} fora da lista de municípios; ignorados.")
        estados.append({
            "uf": uf, "nome": nome, "total": len(lista), "pesquisado": bool(pesq),
            "data": pesq["data"] if pesq else None,
            "metodologia": pesq["metodologia"] if pesq else [],
            "arquivo": pesq["arquivo"] if pesq else None,
        })

    dados = {
        "gerado": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "estados": estados, "municipios": municipios, "processos": processos,
        "mapa": projetar(carregar_malha()),
    }
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    modelo = open(os.path.join(AQUI, "modelo-painel.html"), encoding="utf-8").read()
    html = modelo.replace("/*__DADOS__*/null", js).replace("<!--__LOGO__-->", logo_inline())
    open(SAIDA, "w", encoding="utf-8").write(html)

    pesq = [e["uf"] for e in estados if e["pesquisado"]]
    pend = [e["uf"] for e in estados if not e["pesquisado"]]
    cont = {s: sum(1 for m in municipios if m["sit"] == s) for s in ("com", "sem", "verificar", "pendente")}
    print(f"Painel gravado em {SAIDA} ({len(html) / 1024:.0f} KB)")
    print(f"Estados pesquisados: {', '.join(pesq) or 'nenhum'} | pendentes: {', '.join(pend) or 'nenhum'}")
    print(f"Municípios: {len(municipios)} | com ação {cont['com']} | sem ação {cont['sem']} | "
          f"verificar {cont['verificar']} | pendentes {cont['pendente']} | processos {len(processos)}")
    for a in avisos:
        print("Aviso:", a)


if __name__ == "__main__":
    main()
