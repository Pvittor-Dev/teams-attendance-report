#!/usr/bin/env python3
"""Consolida os relatórios de presença do Teams (CSV) em Excel + painel HTML.

Uso:
    python consolidar_presenca.py                 (usa a pasta ./presencas)
    python consolidar_presenca.py --min-minutos 10

Entrada : CSVs baixados do Teams na pasta ./presencas
Opcional: ./esperados.txt  -> um participante por linha ("Nome" ou "Nome;email")
Saída   : ./saida/Presenca_Consolidada.xlsx  e  ./saida/painel.html
"""
import argparse
import base64
import csv
import io
import json
import re
import sys
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent


# ---------- leitura ----------
def ler_texto(caminho: Path) -> str:
    dados = caminho.read_bytes()
    if dados[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return dados.decode("utf-16")
    return dados.decode("utf-8-sig", errors="replace")


def norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def parse_duracao(txt: str) -> int:
    """'1h 5m 3s' / '1 h 5 min 3 s' -> segundos."""
    txt = (txt or "").lower()
    h = re.search(r"(\d+)\s*h", txt)
    m = re.search(r"(\d+)\s*m(?:in)?(?![a-z])", txt)
    s = re.search(r"(\d+)\s*s", txt)
    return (int(h.group(1)) * 3600 if h else 0) + (int(m.group(1)) * 60 if m else 0) + (int(s.group(1)) if s else 0)


FORMATOS_DATA = [
    "%d/%m/%Y %H:%M:%S", "%d/%m/%Y, %H:%M:%S", "%d/%m/%Y %H:%M",
    "%m/%d/%y, %I:%M:%S %p", "%m/%d/%Y, %I:%M:%S %p", "%Y-%m-%d %H:%M:%S",
]


def parse_data(txt: str):
    txt = (txt or "").strip()
    for f in FORMATOS_DATA:
        try:
            return datetime.strptime(txt, f)
        except ValueError:
            pass
    return None


def hora(txt: str) -> str:
    m = re.search(r"(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?", txt or "")
    if not m:
        return ""
    h, mi, ap = int(m.group(1)), m.group(2), m.group(3)
    if ap:
        h = h % 12 + (12 if ap.lower() == "pm" else 0)
    return f"{h:02d}:{mi}"


FORMATOS_CARIMBO = ["%m/%d/%y, %H:%M:%S", "%d/%m/%y, %H:%M:%S", "%m/%d/%Y, %H:%M:%S", "%d/%m/%Y, %H:%M:%S",
                    "%m/%d/%y, %I:%M:%S %p", "%d/%m/%Y %H:%M:%S"]


def ler_eventos(linhas, caminho: Path):
    """Relatório 'Nome Completo | Atividade | Carimbo de data/hora' (Ingressou / Saiu)."""
    from datetime import date as _date
    eventos = []
    for l in linhas[1:]:
        if len(l) >= 3 and l[0].strip():
            eventos.append((re.sub(r"\s+", " ", l[0]).strip(), norm(l[1]), l[2].strip()))
    if not eventos:
        raise ValueError("arquivo sem eventos")
    # m/d/aa ou d/m/aa? escolhe o formato cuja data fica mais perto da data do arquivo
    ref = datetime.fromtimestamp(caminho.stat().st_mtime).date()
    melhor = None
    for f in FORMATOS_CARIMBO:
        try:
            d = datetime.strptime(eventos[0][2], f)
        except ValueError:
            continue
        dist = abs((d.date() - ref).days)
        if melhor is None or dist < melhor[0]:
            melhor = (dist, f)
    if not melhor:
        raise ValueError("formato de data/hora não reconhecido")
    fmt = melhor[1]
    evs = [(n, a, datetime.strptime(t, fmt)) for n, a, t in eventos]
    fim = max(t for _, _, t in evs)
    dia = min(t for _, _, t in evs).date()
    sess = defaultdict(list)  # chave -> [[entrada, saida|None], ...]
    nomes = {}
    for n, a, t in sorted(evs, key=lambda e: e[2]):
        k = norm(n)
        nomes.setdefault(k, n)
        if a.startswith(("ingress", "joined", "entrou")):
            if not sess[k] or sess[k][-1][1] is not None:
                sess[k].append([t, None])
        elif a.startswith(("saiu", "left")):
            if sess[k] and sess[k][-1][1] is None:
                sess[k][-1][1] = t
    pessoas = {}
    for k, ss in sess.items():
        seg = sum(((b or fim) - a).total_seconds() for a, b in ss)
        ultima = ss[-1][1] or fim
        pessoas[k] = dict(nome=nomes[k], email="", entrada=ss[0][0].strftime("%H:%M"), saida=ultima.strftime("%H:%M"),
                          seg=int(seg), papel="")
    return dia, pessoas


def ler_relatorio(caminho: Path):
    linhas = list(csv.reader(io.StringIO(ler_texto(caminho)), delimiter="\t"))
    if linhas and [norm(x) for x in linhas[0][:2]] == ["nome completo", "atividade"]:
        return ler_eventos(linhas, caminho)
    data = None
    for l in linhas:
        if len(l) > 1 and norm(l[0]) in ("hora de inicio da reuniao", "meeting start time", "start time", "hora de inicio"):
            data = parse_data(l[1])
            break
    if not data:
        m = re.search(r"(\d{4})-(\d{2})-(\d{2})", caminho.name)
        data = datetime(*map(int, m.groups())) if m else datetime.fromtimestamp(caminho.stat().st_mtime)

    # cabeçalho da tabela de participantes
    ini = None
    for i, l in enumerate(linhas):
        c = [norm(x) for x in l]
        if c and c[0] in ("nome", "name") and any(x.startswith(("duracao", "in-meeting", "duration")) for x in c):
            ini = i
            break
    if ini is None:
        raise ValueError("não achei a tabela de participantes")
    cab = [norm(x) for x in linhas[ini]]

    def col(*chaves):
        for i, c in enumerate(cab):
            if any(c.startswith(k) for k in chaves):
                return i
        return None

    i_nome, i_ent = col("nome", "name"), col("hora de entrada", "first join")
    i_sai, i_dur = col("hora de saida", "last leave"), col("duracao", "in-meeting", "duration")
    i_mail, i_papel = col("email"), col("funcao", "role")

    pessoas = {}
    for l in linhas[ini + 1:]:
        if not l or not any(x.strip() for x in l):
            break  # fim da seção
        g = lambda i: l[i].strip() if i is not None and i < len(l) else ""
        nome, mail = g(i_nome), g(i_mail).lower()
        chave = mail or norm(nome)
        dur = parse_duracao(g(i_dur))
        ent, sai = hora(g(i_ent)), hora(g(i_sai))
        if chave in pessoas:  # entrou mais de uma vez
            p = pessoas[chave]
            p["seg"] += dur
            p["entrada"] = min(filter(None, [p["entrada"], ent]), default="")
            p["saida"] = max(filter(None, [p["saida"], sai]), default="")
        else:
            pessoas[chave] = dict(nome=nome, email=mail, entrada=ent, saida=sai, seg=dur, papel=g(i_papel))
    return data.date(), pessoas


def mapa_equipes():
    """equipes.csv -> {nome normalizado: supervisor}"""
    arq = BASE / "equipes.csv"
    mapa = {}
    if arq.exists():
        for l in arq.read_text(encoding="utf-8-sig").splitlines():
            c = [x.strip() for x in l.split(";")]
            if l.startswith("#") or len(c) < 4 or c[0].lower() == "nome":
                continue
            cod = c[4] if len(c) > 4 else ""
            k = norm(c[0])
            if k in mapa and cod and cod not in mapa[k][3].split(" / "):
                cod = mapa[k][3] + " / " + cod  # mesma pessoa em dois cadastros
            mapa[k] = (c[0], c[1].lower(), c[2], cod)
    return mapa


def ler_esperados():
    """esperados.txt + todos os nomes do equipes.csv."""
    lista = []
    arq = BASE / "esperados.txt"
    if arq.exists():
        for l in arq.read_text(encoding="utf-8-sig").splitlines():
            l = l.strip()
            if not l or l.startswith("#"):
                continue
            nome, _, mail = l.partition(";")
            lista.append((nome.strip(), mail.strip().lower()))
    ja = {norm(n) for n, _ in lista}
    for k, (nome, mail, _, _c) in mapa_equipes().items():
        if k not in ja:
            lista.append((nome, mail))
    return lista


# ---------- consolidação ----------
def ler_apelidos():
    """apelidos.csv: 'nome no Teams;nome oficial' (ou IGNORAR no lugar do oficial)."""
    arq = BASE / "apelidos.csv"
    mapa = {}
    if arq.exists():
        for l in arq.read_text(encoding="utf-8-sig").splitlines():
            c = [x.strip() for x in l.split(";")]
            if l.startswith("#") or len(c) < 2 or not c[0] or not c[1]:
                continue
            mapa[norm(c[0])] = c[1]
    return mapa


def tokens(nome):
    return {t for t in re.split(r"[^a-z0-9]+", norm(nome)) if t}


def resolver_nomes(vistos, cadastro, aliases):
    """vistos: {chave: dados}. Devolve ({chave_vista: chave_cadastro|None(ignorar)}, relatório)."""
    apelidos = ler_apelidos()
    res, rel = {}, []
    ocupados = set()

    def oficial(txt):
        return aliases.get(norm(txt))

    # 1) apelido manual e nome/e-mail exato
    for k, p in vistos.items():
        ap = apelidos.get(norm(p["nome"]))
        if ap:
            if ap.upper() == "IGNORAR":
                res[k] = None
                rel.append((p["nome"], "", "ignorado"))
            elif oficial(ap):
                res[k] = oficial(ap)
                ocupados.add(res[k])
                rel.append((p["nome"], cadastro[res[k]]["nome"], "apelido"))
            continue
        for cand in (p["email"], norm(p["nome"]), k):
            if cand in aliases:
                res[k] = aliases[cand]
                ocupados.add(res[k])
                rel.append((p["nome"], cadastro[res[k]]["nome"], "exato"))
                break
    # 2) só uma pessoa da lista contém todas as palavras do nome usado no Teams
    esperados = {c: tokens(v["nome"]) for c, v in cadastro.items() if v["esperado"]}
    pend = {k: p for k, p in vistos.items() if k not in res}
    ambiguos = {}
    for k, p in pend.items():
        t = tokens(p["nome"])
        cands = [c for c, tk in esperados.items() if t and t <= tk]
        if len(cands) == 1:
            res[k] = cands[0]
            ocupados.add(cands[0])
            rel.append((p["nome"], cadastro[cands[0]]["nome"], "único"))
        elif len(cands) > 1:
            ambiguos[k] = cands
    # 3) ambíguos: tira quem já foi reconhecido por outro nome
    for k, cands in ambiguos.items():
        livres = [c for c in cands if c not in ocupados]
        if len(livres) == 1:
            res[k] = livres[0]
            ocupados.add(livres[0])
            rel.append((vistos[k]["nome"], cadastro[livres[0]]["nome"], "inferido (confira)"))
        else:
            rel.append((vistos[k]["nome"], " | ".join(cadastro[c]["nome"] for c in cands), "AMBÍGUO"))
    return res, rel


def consolidar(pasta: Path, min_seg: int):
    dias = {}
    for arq in sorted(pasta.glob("*.csv")):
        try:
            d, pessoas = ler_relatorio(arq)
        except Exception as e:
            print(f"  ! {arq.name}: {e} (ignorado)")
            continue
        if d in dias:  # duas reuniões no mesmo dia: junta
            for k, p in pessoas.items():
                if k in dias[d]:
                    dias[d][k]["seg"] += p["seg"]
                else:
                    dias[d][k] = p
        else:
            dias[d] = pessoas
        print(f"  ok {arq.name}: {d.strftime('%d/%m/%Y')} - {len(pessoas)} participantes")
    if not dias:
        sys.exit("Nenhum CSV válido em " + str(pasta))

    datas = sorted(dias)
    # cadastro de pessoas: esperados primeiro, depois os que apareceram
    cadastro = {}  # chave -> dict(nome, email, esperado)
    aliases = {}   # email/nome normalizado -> chave canônica
    for nome, mail in ler_esperados():
        k = mail or norm(nome)
        cadastro[k] = dict(nome=nome, email=mail, esperado=True)
        aliases[norm(nome)] = k
        if mail:
            aliases[mail] = k
    tem_lista = bool(cadastro)

    vistos = {}
    for d in datas:
        for k, p in dias[d].items():
            vistos.setdefault(k, p)
    res, rel = resolver_nomes(vistos, cadastro, aliases)

    registros = defaultdict(dict)  # chave -> data -> dados
    for d in datas:
        for k, p in dias[d].items():
            if p["seg"] < min_seg:
                continue
            dest = res.get(k, k)
            if k in res and res[k] is None:
                continue  # marcado como IGNORAR
            if dest not in cadastro:
                cadastro[dest] = dict(nome=p["nome"], email=p["email"], esperado=False)
                if k not in res:
                    rel.append((p["nome"], "", "extra (fora da lista)"))
                    res[k] = dest
            if d in registros[dest]:  # dois nomes da mesma pessoa no mesmo dia: junta
                a = registros[dest][d]
                a["seg"] += p["seg"]
                a["entrada"] = min(filter(None, [a["entrada"], p["entrada"]]), default="")
                a["saida"] = max(filter(None, [a["saida"], p["saida"]]), default="")
            else:
                registros[dest][d] = dict(p)

    saida_dir = BASE / "saida"
    saida_dir.mkdir(exist_ok=True)
    ambig = {r[0] for r in rel if r[2] == "AMBÍGUO"}
    rel = [r for r in rel if not (r[2].startswith("extra") and r[0] in ambig)]
    with open(saida_dir / "conferencia_nomes.csv", "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["Nome no Teams", "Reconhecido como", "Como"])
        w.writerows(sorted(set(rel), key=lambda r: (r[2], r[0].lower())))
    contagem = defaultdict(int)
    for r in set(rel):
        contagem[r[2]] += 1
    print("  nomes: " + ", ".join(f"{v} {k}" for k, v in sorted(contagem.items())))

    eqmap = mapa_equipes()
    pessoas = []
    for k, c in cadastro.items():
        reg = registros.get(k, {})
        presentes = len(reg)
        entradas = [int(h[:2]) * 60 + int(h[3:]) for h in (r["entrada"] for r in reg.values()) if h]
        media_ent = f"{sum(entradas)//len(entradas)//60:02d}:{sum(entradas)//len(entradas)%60:02d}" if entradas else ""
        pessoas.append(dict(
            nome=c["nome"], email=c["email"], equipe=(eqmap.get(norm(c["nome"])) or ("", "", "", ""))[2].title(), cod=(eqmap.get(norm(c["nome"])) or ("", "", "", ""))[3], esperado=c["esperado"] or not tem_lista,
            presentes=presentes, total=len(datas), pct=round(100 * presentes / len(datas), 1),
            min_medio=round(sum(r["seg"] for r in reg.values()) / 60 / presentes, 1) if presentes else 0,
            entrada_media=media_ent,
            dias={d.isoformat(): (dict(ent=r["entrada"], sai=r["saida"], min=round(r["seg"] / 60, 1)) if (r := reg.get(d)) else None) for d in datas},
        ))
    pessoas.sort(key=lambda p: (-p["pct"], p["nome"].lower()))
    return datas, pessoas, tem_lista


# ---------- Excel ----------
VERDE, VERMELHO, CINZA = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFC7CE"), PatternFill("solid", fgColor="1F3A5F")
BRANCO = Font(color="FFFFFF", bold=True)


def cabecalho(ws, titulos):
    ws.append(titulos)
    for c in ws[1]:
        c.fill, c.font, c.alignment = CINZA, BRANCO, Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def gerar_excel(datas, pessoas, destino: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    cabecalho(ws, ["Cód.", "Nome", "E-mail", "Dias presente", "Dias de reunião", "% presença", "Tempo médio (min)", "Entrada média", "Esperado?"])
    for p in pessoas:
        ws.append([p.get("cod", ""), p["nome"], p["email"], p["presentes"], p["total"], p["pct"] / 100, p["min_medio"], p["entrada_media"], "Sim" if p["esperado"] else "Não"])
        ws.cell(ws.max_row, 6).number_format = "0.0%"
    for i, w in enumerate([10, 32, 34, 14, 16, 12, 18, 14, 11], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = ws.dimensions

    ws2 = wb.create_sheet("Chamada")
    cabecalho(ws2, ["Cód.", "Nome"] + [d.strftime("%d/%m") for d in datas] + ["Presenças", "% presença"])
    for p in pessoas:
        linha = [p.get("cod", ""), p["nome"]] + ["P" if p["dias"][d.isoformat()] else "F" for d in datas] + [p["presentes"], p["pct"] / 100]
        ws2.append(linha)
        r = ws2.max_row
        for j in range(3, 3 + len(datas)):
            c = ws2.cell(r, j)
            c.fill, c.alignment = (VERDE if c.value == "P" else VERMELHO), Alignment(horizontal="center")
        ws2.cell(r, 4 + len(datas)).number_format = "0.0%"
    ws2.column_dimensions["A"].width = 10
    ws2.column_dimensions["B"].width = 32
    ws2.freeze_panes = "C2"

    ws3 = wb.create_sheet("Detalhe")
    cabecalho(ws3, ["Data", "Cód.", "Nome", "E-mail", "Entrada", "Saída", "Minutos na reunião"])
    for d in datas:
        for p in pessoas:
            x = p["dias"][d.isoformat()]
            if x:
                ws3.append([d, p.get("cod", ""), p["nome"], p["email"], x["ent"], x["sai"], x["min"]])
                ws3.cell(ws3.max_row, 1).number_format = "dd/mm/yyyy"
    for i, w in enumerate([12, 10, 32, 34, 10, 10, 20], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.auto_filter.ref = ws3.dimensions
    wb.save(destino)


# ---------- painel ----------
def gerar_painel(datas, pessoas, destino: Path, titulo=None):
    dados = dict(
        titulo=titulo,
        gerado=datetime.now().strftime("%d/%m/%Y %H:%M"),
        datas=[d.isoformat() for d in datas],
        pessoas=pessoas,
    )
    html = (BASE / "painel_template.html").read_text(encoding="utf-8")
    logo = BASE / "logo.png"
    uri = "data:image/png;base64," + base64.b64encode(logo.read_bytes()).decode() if logo.exists() else ""
    html = html.replace("__LOGO__", uri).replace("__DADOS__", json.dumps(dados, ensure_ascii=False))
    destino.write_text(html, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pasta", default=str(BASE / "presencas"))
    ap.add_argument("--min-minutos", type=float, default=1, help="tempo mínimo para contar como presente (padrão 1)")
    a = ap.parse_args()
    pasta = Path(a.pasta)
    pasta.mkdir(exist_ok=True)
    saida = BASE / "saida"
    saida.mkdir(exist_ok=True)

    print("Lendo relatórios...")
    datas, pessoas, _ = consolidar(pasta, int(a.min_minutos * 60))
    gerar_excel(datas, pessoas, saida / "Presenca_Consolidada.xlsx")
    gerar_painel(datas, pessoas, saida / "painel.html")
    print(f"\nPronto: {len(datas)} dias, {len(pessoas)} pessoas.\n  {saida / 'Presenca_Consolidada.xlsx'}\n  {saida / 'painel.html'}")


if __name__ == "__main__":
    main()
