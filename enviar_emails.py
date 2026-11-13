#!/usr/bin/env python3
"""Gera os painéis (PDF) e envia pelo Outlook, em DOIS e-mails por destinatário:
  1) "Reunião do dia"  -> só a última reunião (quem veio, quem faltou)
  2) "Média do mês"    -> mês corrente até a última reunião
Destinatários: cada supervisor (só a equipe dele) + o painel GERAL para EMAIL_GERAL.

Uso:
    python enviar_emails.py                      -> abre RASCUNHOS no Outlook (nada é enviado)
    python enviar_emails.py --enviar             -> envia de verdade
    python enviar_emails.py --tipo dia           -> só o e-mail do dia (ou: mes)
    python enviar_emails.py --apenas supervisor1 | --apenas geral
    python enviar_emails.py --teste-para voce@empresa.com   -> manda tudo para você, com [TESTE]
    python enviar_emails.py --so-pdf             -> só gera os PDFs em saida/equipes/
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

import consolidar_presenca as cp

BASE = cp.BASE
EMAIL_GERAL = os.environ.get("PRESENCA_EMAIL_GERAL", "")  # quem recebe o painel geral (opcional)
EDGE = next((p for p in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                         r"C:\Program Files\Microsoft\Edge\Application\msedge.exe") if Path(p).exists()), None)
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
ESTILO = "font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1b2430"
RODAPE = '<p style="color:#6b7686;font-size:12px">E-mail gerado automaticamente a partir do relatório de presença do Teams.</p>'


# ---------- equipes ----------
def ler_equipes():
    arq = BASE / "equipes.csv"
    if not arq.exists():
        sys.exit("Falta o arquivo equipes.csv (Nome;Email;Supervisor;EmailSupervisor).")
    equipes = {}  # (email_sup, supervisor) -> [(nome, email)]
    for l in arq.read_text(encoding="utf-8-sig").splitlines():
        l = l.strip()
        if not l or l.startswith("#"):
            continue
        c = [x.strip() for x in l.split(";")]
        if len(c) < 4 or c[0].lower() == "nome":
            continue
        equipes.setdefault((c[3].lower(), c[2]), []).append((c[0], c[1].lower()))
    return equipes


def pertence(p, membros):
    nomes = {cp.norm(n) for n, _ in membros}
    mails = {m for _, m in membros if m}
    return cp.norm(p["nome"]) in nomes or (p["email"] and p["email"] in mails)


def montar_equipe(pessoas, membros, datas):
    equipe = [p for p in pessoas if pertence(p, membros)]
    vistos = {cp.norm(p["nome"]) for p in equipe} | {p["email"] for p in equipe if p["email"]}
    for nome, mail in membros:  # quem nunca apareceu entra com 0%
        if cp.norm(nome) not in vistos and mail not in vistos:
            equipe.append(dict(nome=nome, email=mail, esperado=True, equipe="", presentes=0, total=len(datas), pct=0.0,
                               min_medio=0, entrada_media="", dias={d.isoformat(): None for d in datas}))
    return equipe


def recorte(pessoas, datas):
    """Recalcula presença só para as datas escolhidas."""
    iso = [d.isoformat() for d in datas]
    out = []
    for p in pessoas:
        dias = {d: p["dias"][d] for d in iso}
        reg = [x for x in dias.values() if x]
        ents = [int(x["ent"][:2]) * 60 + int(x["ent"][3:]) for x in reg if x["ent"]]
        em = sum(ents) // len(ents) if ents else None
        out.append({**p, "dias": dias, "presentes": len(reg), "total": len(datas),
                    "pct": round(100 * len(reg) / len(datas), 1),
                    "min_medio": round(sum(x["min"] for x in reg) / len(reg), 1) if reg else 0,
                    "entrada_media": f"{em // 60:02d}:{em % 60:02d}" if em is not None else ""})
    return sorted(out, key=lambda p: (-p["pct"], p["nome"].lower()))


# ---------- PDF ----------
def html_para_pdf(html: Path, pdf: Path):
    if not EDGE:
        sys.exit("Microsoft Edge não encontrado (necessário para gerar o PDF).")
    with tempfile.TemporaryDirectory() as perfil:
        subprocess.run([EDGE, "--headless=new", "--disable-gpu", f"--user-data-dir={perfil}",
                        "--no-pdf-header-footer", f"--print-to-pdf={pdf}", "--virtual-time-budget=4000",
                        html.as_uri()], check=True, capture_output=True, timeout=90)
    if not pdf.exists():
        sys.exit(f"Falhou ao gerar {pdf.name}")


# ---------- corpos de e-mail ----------
def rot(p):
    return f"{p['cod']} · {p['nome']}" if p.get("cod") else p["nome"]


def nota_fora(fora):
    if not fora:
        return ""
    return (f'<p style="color:#6b7686;font-size:13px">Também entraram {len(fora)} participante(s) fora da lista de vendedores '
            "(supervisores, gerência, TI e convidados).</p>")


def lista(itens):
    return "<ul>" + ("".join(f"<li>{x}</li>" for x in itens) or "<li>—</li>") + "</ul>"


def saudacao(sup, texto):
    nome = ", " + sup.split()[0] if sup else ""
    return f"<p>Olá{nome}!</p><p>{texto}</p>"


def corpo_dia(sup, equipe, d):
    iso = d.isoformat()
    fora = [p for p in equipe if p.get("esperado") is False and p["dias"][iso]]
    equipe = [p for p in equipe if p.get("esperado") is not False]
    pres = sorted((p for p in equipe if p["dias"][iso]), key=lambda p: p["dias"][iso]["ent"] or "99")
    aus = sorted((p for p in equipe if not p["dias"][iso]), key=lambda p: p["nome"])
    pct = 100 * len(pres) / len(equipe) if equipe else 0
    quem = "da sua equipe" if sup else "(visão geral, todas as equipes)"
    intro = f"Segue a presença na reunião diária de <b>{d.strftime('%d/%m/%Y')}</b> {quem}. O painel está no PDF anexo."
    itens_pres = [f"{rot(p)} — entrou às {p['dias'][iso]['ent']}" if p["dias"][iso]["ent"] else rot(p) for p in pres]
    return f"""<div style="{ESTILO}">{saudacao(sup, intro)}
<table style="border-collapse:collapse;font-size:14px"><tr><td style="padding:4px 16px 4px 0">Presentes</td><td><b>{len(pres)} de {len(equipe)} ({pct:.0f}%)</b></td></tr></table>
<p><b>Ausentes ({len(aus)}):</b></p>{lista(rot(p) for p in aus)}
<p><b>Presentes ({len(pres)}):</b></p>{lista(itens_pres)}
{nota_fora(fora)}
{RODAPE}</div>"""


def corpo_mes(sup, equipe, datas, resumo=None):
    equipe = [p for p in equipe if p.get("esperado") is not False]
    media = sum(p["pct"] for p in equipe) / len(equipe) if equipe else 0
    por_dia = [(d, sum(1 for p in equipe if p["dias"][d.isoformat()])) for d in datas]
    melhor, pior = max(por_dia, key=lambda x: x[1]), min(por_dia, key=lambda x: x[1])
    baixos = [p for p in equipe if p["pct"] < 60]
    mes = f"{MESES[datas[-1].month - 1]}/{datas[-1].year}"
    quem = "da sua equipe" if sup else "(visão geral, todas as equipes)"
    menores = sorted(equipe, key=lambda p: (p["pct"], p["nome"]))[: (len(equipe) if sup else 10)]
    intro = (f"Segue a <b>média de presença do mês</b> ({mes}, de {datas[0].strftime('%d/%m')} a {datas[-1].strftime('%d/%m')}"
             f" — {len(datas)} reuniões) {quem}. O painel está no PDF anexo.")
    titulo_lista = "Presença por pessoa (da menor para a maior)" if sup else "10 menores presenças"
    itens = [f"{rot(p)} — {p['pct']}% ({p['presentes']}/{p['total']})" for p in menores]
    return f"""<div style="{ESTILO}">{saudacao(sup, intro)}
<table style="border-collapse:collapse;font-size:14px">
<tr><td style="padding:4px 16px 4px 0">Frequência média</td><td><b>{media:.1f}%</b></td></tr>
<tr><td style="padding:4px 16px 4px 0">Melhor dia</td><td><b>{melhor[0].strftime('%d/%m')}</b> ({melhor[1]} presentes)</td></tr>
<tr><td style="padding:4px 16px 4px 0">Pior dia</td><td><b>{pior[0].strftime('%d/%m')}</b> ({pior[1]} presentes)</td></tr></table>
{resumo or ""}
<p><b>{titulo_lista}:</b></p>{lista(itens)}
<p><b>Abaixo de 60%:</b> {str(len(baixos)) + " pessoa(s)" if baixos else "ninguém"}</p>
{RODAPE}</div>"""


def resumo_supervisores(pacotes):
    ordem = sorted(pacotes, key=lambda x: -sum(p["pct"] for p in x[1]) / max(len(x[1]), 1))
    linhas = "".join(
        f"<tr><td style='padding:3px 16px 3px 0'>{sup}</td><td style='padding:3px 16px 3px 0'>{len(eq)} pessoas</td>"
        f"<td><b>{sum(p['pct'] for p in eq) / max(len(eq), 1):.1f}%</b></td></tr>" for sup, eq in ordem)
    if not linhas:
        return ""
    return f"<p><b>Frequência por supervisor:</b></p><table style='border-collapse:collapse;font-size:14px'>{linhas}</table>"


def aviso_geral(texto):
    """E-mail de alerta só para o dono (EMAIL_GERAL)."""
    print("AVISO:", texto)
    if not EMAIL_GERAL:
        return
    import win32com.client
    m = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
    m.To, m.Subject = EMAIL_GERAL, "⚠ Presença da reunião diária — nada foi enviado"
    m.HTMLBody = f'<div style="{ESTILO}"><p>{texto}</p>{RODAPE}</div>'
    m.Send()


# ---------- principal ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pasta", default=str(BASE / "presencas"))
    ap.add_argument("--min-minutos", type=float, default=1)
    ap.add_argument("--tipo", choices=["dia", "mes", "ambos"], default="ambos")
    ap.add_argument("--enviar", action="store_true", help="envia de verdade (sem isso, só abre rascunhos)")
    ap.add_argument("--so-pdf", action="store_true")
    ap.add_argument("--apenas", help="só este supervisor (parte do e-mail, ex.: parte do e-mail) ou 'geral'")
    ap.add_argument("--teste-para", help="manda tudo para este e-mail em vez dos destinatários reais ([TESTE] no assunto)")
    ap.add_argument("--sem-geral", action="store_true")
    ap.add_argument("--auto", action="store_true",
                    help="modo agendado: só envia se o CSV de hoje estiver na pasta e se ainda não foi enviado hoje")
    a = ap.parse_args()

    equipes = ler_equipes()
    if a.apenas and a.apenas.lower() != "geral":
        equipes = {k: v for k, v in equipes.items() if a.apenas.lower() in k[0]}
    geral_ok = bool(EMAIL_GERAL) and not a.sem_geral and (not a.apenas or a.apenas.lower() == "geral")
    if a.apenas and a.apenas.lower() == "geral":
        equipes = {}

    print("Lendo relatórios...")
    datas, pessoas, _ = cp.consolidar(Path(a.pasta), int(a.min_minutos * 60))
    ultimo = datas[-1]
    estado = BASE / "saida" / "enviados.json"
    enviados = json.loads(estado.read_text(encoding="utf-8")) if estado.exists() else []
    if a.auto:
        a.enviar = True
        if ultimo != date.today():
            aviso_geral(f"Não encontrei o CSV de presença de hoje ({date.today().strftime('%d/%m/%Y')}) na pasta presencas. "
                        f"O último é de {ultimo.strftime('%d/%m/%Y')}. Nada foi enviado aos supervisores. "
                        "Salve o CSV na pasta e rode enviar_hoje.bat.")
            return
        if ultimo.isoformat() in enviados:
            print(f"Já enviado em {ultimo.strftime('%d/%m/%Y')}; nada a fazer.")
            return
    do_mes = [d for d in datas if (d.year, d.month) == (ultimo.year, ultimo.month)]
    nome_mes = f"{MESES[ultimo.month - 1]}/{ultimo.year}"
    recortes = {"dia": ([ultimo], f"Reunião de {ultimo.strftime('%d/%m/%Y')}"),
                "mes": (do_mes, f"Média do mês — {nome_mes}")}
    tipos = ["dia", "mes"] if a.tipo == "ambos" else [a.tipo]

    saida = BASE / "saida" / "equipes"
    saida.mkdir(parents=True, exist_ok=True)
    base_equipes = {k: montar_equipe(pessoas, m, datas) for k, m in equipes.items()}
    emails = []  # (assunto, destino, corpo, pdf, rotulo)

    for tipo in tipos:
        ds, titulo = recortes[tipo]
        if tipo == "dia":
            assunto = f"Presença na reunião diária — {ultimo.strftime('%d/%m/%Y')}"
        else:
            assunto = f"Média de presença do mês — {nome_mes} (até {ultimo.strftime('%d/%m')})"
        pac = []
        for (mail_sup, sup), eq_full in base_equipes.items():
            eq = recorte(eq_full, ds)
            slug = mail_sup.split("@")[0]
            html, pdf = saida / f"{tipo}_{slug}.html", saida / f"Presenca_{tipo}_{slug}.pdf"
            cp.gerar_painel(ds, eq, html, titulo=f"{titulo} — equipe de {sup}")
            html_para_pdf(html, pdf)
            html.unlink()
            print(f"  PDF: {pdf.name} ({len(eq)} pessoas)")
            pac.append((sup, eq))
            corpo = corpo_dia(sup, eq, ds[-1]) if tipo == "dia" else corpo_mes(sup, eq, ds)
            emails.append((assunto, mail_sup, corpo, pdf, sup))
        if geral_ok:
            eq = recorte(pessoas, ds)
            html, pdf = saida / f"{tipo}_geral.html", saida / f"Presenca_{tipo}_GERAL.pdf"
            cp.gerar_painel(ds, eq, html, titulo=f"{titulo} — visão geral")
            html_para_pdf(html, pdf)
            html.unlink()
            print(f"  PDF: {pdf.name} ({len(eq)} pessoas)")
            corpo = corpo_dia(None, eq, ds[-1]) if tipo == "dia" else corpo_mes(None, eq, ds, resumo_supervisores(pac))
            emails.append((assunto, EMAIL_GERAL, corpo, pdf, "GERAL"))

    if a.so_pdf:
        return
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    for assunto, destino, corpo, pdf, rotulo in emails:
        m = outlook.CreateItem(0)
        real = a.teste_para or destino
        m.To, m.Subject, m.HTMLBody = real, ("[TESTE] " if a.teste_para else "") + assunto, corpo
        m.Attachments.Add(str(pdf))
        if a.enviar:
            m.Send()
            print(f"  enviado -> {rotulo} <{real}>: {assunto}")
        else:
            m.Display(False)
            print(f"  rascunho aberto -> {rotulo} <{real}>: {assunto}")
    if a.enviar and not a.teste_para and not a.apenas and a.tipo == "ambos":
        estado.write_text(json.dumps(sorted(set(enviados + [ultimo.isoformat()]))), encoding="utf-8")
    if not a.enviar:
        print("\nNada foi enviado. Confira os rascunhos no Outlook; para enviar direto use --enviar.")


if __name__ == "__main__":
    main()
