# -*- coding: utf-8 -*-
"""ONB_92 – Compilazione dei modelli di onboarding dalla scheda dati (ONB_91).

Uso:  python ONB_92_compila_modelli.py <ONB_91_compilata.xlsx> <cartella_destinazione>

- Copia tutti i modelli .docx del kit nella cartella di destinazione (stessa struttura).
- Sostituisce i campi «CODICE» valorizzati nella scheda dati; i campi vuoti restano evidenziati.
- I campi «DA COMPILARE» non sono toccati.
- Elaborazione locale: nessun dato è trasmesso all'esterno.
Requisiti: python 3, python-docx, openpyxl.
"""
import os, re, sys, shutil, datetime
import docx
from docx.enum.text import WD_COLOR_INDEX
from openpyxl import load_workbook

KIT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PH = re.compile(r"«([A-Z]{3}_[A-Z0-9_]+)»")


def load(path):
    ws = load_workbook(path, data_only=True)["Dati"]
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] and row[3] not in (None, ""):
            v = row[3]
            if isinstance(v, (datetime.date, datetime.datetime)):
                v = v.strftime("%d/%m/%Y")
            out[str(row[0]).strip()] = str(v)
    return out


def runs(d):
    blocks = [d]
    for s in d.sections:
        blocks += [s.header, s.footer, s.first_page_header, s.first_page_footer]
    for b in blocks:
        for p in b.paragraphs:
            yield from p.runs
        for t in b.tables:
            for row in t.rows:
                for c in row.cells:
                    for p in c.paragraphs:
                        yield from p.runs


def fill(src, dst, data):
    d = docx.Document(src)
    n = 0
    for r in runs(d):
        m = PH.fullmatch(r.text.strip())
        if m and m.group(1) in data:
            r.text = data[m.group(1)]
            r.font.highlight_color = None
            n += 1
        elif PH.search(r.text):
            new = PH.sub(lambda x: data.get(x.group(1), x.group(0)), r.text)
            if new != r.text:
                r.text = new; n += 1
    d.save(dst)
    return n


def main():
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(1)
    data = load(sys.argv[1]); out = sys.argv[2]
    tot = 0
    for root, _, files in os.walk(KIT):
        if "40_Strumenti" in root:
            continue
        for f in files:
            if f.endswith(".docx"):
                rel = os.path.relpath(os.path.join(root, f), KIT)
                dst = os.path.join(out, rel); os.makedirs(os.path.dirname(dst), exist_ok=True)
                tot += fill(os.path.join(root, f), dst, data)
            elif f.endswith(".xlsx") and not f.startswith("ONB_91"):
                rel = os.path.relpath(os.path.join(root, f), KIT)
                dst = os.path.join(out, rel); os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(os.path.join(root, f), dst)
    print(f"Campi sostituiti: {tot}. Campi valorizzati nella scheda: {len(data)}. Output: {out}")


if __name__ == "__main__":
    main()
