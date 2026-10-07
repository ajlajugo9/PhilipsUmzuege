"""
Erzeugt Philips_Umzüge_Rechner.html aus den Original-Excel-Dateien.

- Gegenstandskataloge (RE-Werte) aus "m³ rechner - Firma.xlsx" / "m³ rechner - Privat.xlsx"
- Preise / Faktoren aus dem Kostenteil der Rechner (mit Prüfung der erwarteten Formeln)
- Kostenvoranschlag-Vorlage (alle ZIP-Teile unverändert) + Logo aus "Kostenvoranschlag Beispiel.xlsx"

Aufruf:  python build_app.py [Ordner mit den Excel-Dateien]
Standard: der Ordner oberhalb des Projektordners (Desktop).
"""
import base64
import json
import re
import sys
import zipfile
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
SRC_DIR = Path(sys.argv[1]) if len(sys.argv) > 1 else PROJECT.parent

FIRMA = SRC_DIR / "m³ rechner - Firma.xlsx"
PRIVAT = SRC_DIR / "m³ rechner - Privat.xlsx"
TEMPLATE_CANDIDATES = ["Kostenvoranschlag Beispiel.xlsx"]

warnings = []


def warn(msg):
    warnings.append(msg)
    print("WARNUNG:", msg)


def find_template():
    for name in TEMPLATE_CANDIDATES:
        for d in (SRC_DIR, PROJECT):
            p = d / name
            if p.exists():
                return p
    raise SystemExit("Kostenvoranschlag-Vorlage nicht gefunden: " + ", ".join(TEMPLATE_CANDIDATES))


CATEGORY_NAMES = {
    "WOHNZIMMER": "Wohnzimmer",
    "ESSZIMMER": "Esszimmer",
    "SCHLAFZIMMER": "Schlafzimmer",
    "ARBEITSZIMMER": "Arbeitszimmer",
    "KINDERZIMMER / STUDIO": "Kinderzimmer / Studio",
    "DIELE / BAD": "Diele / Bad",
    "KÜCHE": "Küche",
    "KELLER / SPEICHER / GARTEN": "Keller / Speicher / Garten",
}


def read_block(ws, mode, qty_col, name_col, re_col, row_from, row_to, start_cat, catalog):
    """Liest einen Gegenstandsblock. Kopfzeilen (GROSSBUCHSTABEN) wechseln die Kategorie."""
    cat = start_cat
    for r in range(row_from, row_to + 1):
        name = ws[f"{name_col}{r}"].value
        re_val = ws[f"{re_col}{r}"].value
        if not isinstance(name, str) or not name.strip():
            continue
        name = name.strip()
        if name in CATEGORY_NAMES:
            cat = CATEGORY_NAMES[name]
            continue
        if name.upper() in ("ÜBERTRAG", "GESAMTSUMME", "GEGENSTAND"):
            continue
        if not isinstance(re_val, (int, float)):
            warn(f"{mode}: '{name}' ({name_col}{r}) hat keinen RE-Wert in der Excel-Datei -> nicht übernommen")
            continue
        # Prüfen, ob die Excel-Zeile ihren Wert auch wirklich in die Summe übernimmt
        sum_col = chr(ord(re_col) + 1)
        f = ws[f"{sum_col}{r}"].value
        expected = {f"={re_col}{r}*{qty_col}{r}", f"={qty_col}{r}*{re_col}{r}"}
        if f not in expected:
            warn(f"{mode}: '{name}' ({sum_col}{r}) enthält '{f}' statt Formel {re_col}{r}*{qty_col}{r} "
                 f"(Excel zählt diesen Gegenstand nicht). Im Rechner mit RE {re_val} übernommen.")
        catalog.append({
            "id": f"{mode}-{name_col}{r}",
            "name": name,
            "re": re_val,
            "cat": cat,
        })


def expect_formula(ws, ref, expected):
    v = ws[ref].value
    if v != expected:
        warn(f"{ws.parent.properties.title or ''} {ref}: erwartet '{expected}', gefunden '{v}'")


def num(ws, ref, blank_is_zero=False):
    v = ws[ref].value
    if v is None and blank_is_zero:
        warn(f"{ref} ist leer -> wie in Excel als 0 übernommen")
        return 0
    if not isinstance(v, (int, float)):
        raise SystemExit(f"Zahl erwartet in {ref}, gefunden {v!r}")
    return v


def read_costs(ws, mode, r_base, r_extra_from, r_extra_to, r_helfer, r_ent):
    """Kostenteil. r_base = Zeile 'Wohnfläche m²' (Grundpreis)."""
    rb = r_base
    m3 = f"E{rb}"
    expect_formula(ws, f"H{rb}", f"={m3}*F{rb}+130")
    expect_formula(ws, f"H{rb+1}", f"=C{rb+1}*D{rb+1}")
    expect_formula(ws, f"H{rb+2}", f"={m3}*C{rb+2}*D{rb+2}")
    expect_formula(ws, f"D{rb+3}", f"=C{rb+3}/10")
    expect_formula(ws, f"H{rb+3}", f"=D{rb+3}*{m3}")
    expect_formula(ws, f"D{rb+4}", f"=C{rb+4}/10")
    expect_formula(ws, f"H{rb+4}", f"=D{rb+4}*{m3}")
    expect_formula(ws, f"H{rb+5}", f"=IF({m3}>21,C{rb+5}*D{rb+5}*1.5,C{rb+5}*D{rb+5})")

    base = {
        "pricePerM3": num(ws, f"F{rb}"),
        "baseFee": 130,
        "transporterPrice": num(ws, f"D{rb+1}"),
        "transporterDefault": num(ws, f"C{rb+1}"),
        "floorFactor": num(ws, f"D{rb+2}"),
        "walkDivisor": 10,
        "kmPrice": num(ws, f"D{rb+5}"),
        "kmThresholdM3": 21,
        "kmFactor": 1.5,
    }

    # Welche Zeilen zählen zu "Helfer benötigt" (=(H..+H..)/46/8)
    hf = ws[f"H{r_helfer}"].value
    m = re.match(r"^=\((.*)\)/46/8$", hf or "")
    if not m:
        raise SystemExit(f"Helfer-Formel unerwartet: {hf}")
    labor_rows = {int(x) for x in re.findall(r"H(\d+)", m.group(1))}

    extras = []
    for r in range(r_extra_from, r_extra_to + 1):
        label = ws[f"B{r}"].value
        price = num(ws, f"D{r}")
        qty_cell = ws[f"C{r}"].value
        f = ws[f"H{r}"].value
        if f != f"=C{r}*D{r}":
            warn(f"{mode}: Kostenzeile '{label}' H{r} = '{f}' statt =C{r}*D{r}. Im Rechner korrekt als Menge × Preis berechnet.")
        e = {"id": f"{mode}-x{r}", "label": label.strip(), "price": price, "labor": r in labor_rows}
        if isinstance(qty_cell, str) and qty_cell.startswith("="):
            mm = re.match(r"^=A(\d+)/(\d+(?:\.\d+)?)$", qty_cell)
            if not mm:
                raise SystemExit(f"Unbekannte Mengenformel {qty_cell} in C{r}")
            e["autoFrom"] = f"{mode}-B{mm.group(1)}"
            e["autoDivisor"] = float(mm.group(2))
        extras.append(e)

    labor_base = {
        "base": rb in labor_rows,
        "floors": (rb + 2) in labor_rows,
        "walkLoad": (rb + 3) in labor_rows,
        "walkUnload": (rb + 4) in labor_rows,
        "transporter": (rb + 1) in labor_rows,
        "km": (rb + 5) in labor_rows,
    }

    # Entsorgung
    re_ = r_ent
    expect_formula(ws, f"E{re_+2}", f"=(C{re_+2}*E{re_+1})/1000")
    expect_formula(ws, f"C{re_+12}", f"=E{re_+1}/10")
    expect_formula(ws, f"E{re_+16}", f"=C{re_+16}")
    expect_formula(ws, f"E{re_+19}", f"=E{re_+7}+E{re_+10}+E{re_+14}+E{re_+16}")
    disposal = {
        "kgPerM3": num(ws, f"C{re_+2}", True),
        "woodShare": num(ws, f"C{re_+3}", True),
        "bulkShare": num(ws, f"C{re_+4}", True),
        "woodPerT": num(ws, f"C{re_+5}", True),
        "bulkPerT": num(ws, f"C{re_+6}", True),
        "transporterCount": num(ws, f"C{re_+9}", True),
        "transporterPrice": num(ws, f"D{re_+9}", True),
        "staffDivisor": 10,
        "staffDayRate": num(ws, f"C{re_+13}", True),
        "material": num(ws, f"C{re_+16}", True),
    }
    return base, extras, labor_base, disposal


def excel_sample(path, mode, catalog, extras, r_base, ref_m3, ref_netto, ref_disp):
    """Eingaben und von Excel gespeicherte Ergebnisse der Datei – Referenzfall für die Selbsttests."""
    wsv = openpyxl.load_workbook(path, data_only=True).active
    items = {}
    for it in catalog:
        col, row = it["id"].split("-")[1][0], it["id"].split("-")[1][1:]
        q = wsv[f"{chr(ord(col) - 1)}{row}"].value
        if isinstance(q, (int, float)) and q:
            items[it["id"]] = q
    ex = {}
    for e in extras:
        if not e.get("autoFrom"):
            q = wsv[f"C{e['id'].split('-x')[1]}"].value
            ex[e["id"]] = q if isinstance(q, (int, float)) else 0
    keys = ["transporter", "floors", "walkLoad", "walkUnload", "km"]
    params = {k: (wsv[f"C{r_base + 1 + i}"].value or 0) for i, k in enumerate(keys)}
    res = {"m3": wsv[ref_m3].value, "netto": wsv[ref_netto].value, "disposal": wsv[ref_disp].value}
    if any(not isinstance(v, (int, float)) for v in res.values()):
        warn(f"{mode}: keine gespeicherten Excel-Ergebnisse gefunden – Referenztest entfällt")
        return None
    return {"items": items, "extras": ex, "params": params, "excel": res}


def build_firma():
    ws = openpyxl.load_workbook(FIRMA).active
    catalog = []
    read_block(ws, "firma", "A", "B", "C", 2, 47, "Arbeitszimmer", catalog)
    expect_formula(ws, "D48", "=SUM(D2:D47)")
    expect_formula(ws, "H54", "=C54/10")
    expect_formula(ws, "H57", "=H54/18")
    base, extras, labor_base, disposal = read_costs(ws, "firma", 70, 76, 93, 56, 97)
    sample = excel_sample(FIRMA, "firma", catalog, extras, 70, "H54", "H94", "E116")
    return {"label": "Firma", "catalog": catalog, "base": base, "extras": extras,
            "laborBase": labor_base, "disposal": disposal, "excelSample": sample}


def build_privat():
    ws = openpyxl.load_workbook(PRIVAT).active
    catalog = []
    # Blöcke exakt wie die Summen C126..C132 im Privat-Rechner
    expect_formula(ws, "C126", "=SUM(D3:D37)")
    expect_formula(ws, "C127", "=SUM(I3:I35)")
    expect_formula(ws, "C128", "=SUM(D45:D89)")
    expect_formula(ws, "C129", "=SUM(H44:H89)")
    expect_formula(ws, "C130", "=SUM(D94:D120)")
    expect_formula(ws, "C131", "=SUM(I94:I120)")
    expect_formula(ws, "C132", "=SUM(M3:M56)")
    expect_formula(ws, "H126", "=C134/10")
    read_block(ws, "privat", "A", "B", "C", 3, 37, "Wohnzimmer", catalog)
    read_block(ws, "privat", "E", "F", "H", 3, 35, "Wohnzimmer", catalog)
    read_block(ws, "privat", "A", "B", "C", 44, 89, "Schlafzimmer", catalog)
    read_block(ws, "privat", "E", "F", "G", 44, 89, "Esszimmer", catalog)
    read_block(ws, "privat", "A", "B", "C", 94, 120, "Küche", catalog)
    read_block(ws, "privat", "E", "F", "H", 94, 120, "Keller / Speicher / Garten", catalog)
    read_block(ws, "privat", "J", "K", "L", 3, 56, "Sonstiges", catalog)
    base, extras, labor_base, disposal = read_costs(ws, "privat", 142, 148, 166, 128, 170)
    sample = excel_sample(PRIVAT, "privat", catalog, extras, 142, "H126", "H167", "E189")
    return {"label": "Privat", "catalog": catalog, "base": base, "extras": extras,
            "laborBase": labor_base, "disposal": disposal, "excelSample": sample}


# Kundenzellen der Vorlage: Der Beispiel-Kostenvoranschlag enthält echte Kundendaten, die nicht in die
# (veröffentlichte) App gehören. Die App überschreibt diese Zellen beim Export ohnehin.
CUSTOMER_CELLS = {
    "A10": "Muster GmbH",
    "A11": "Musterstr. 1",
    "A12": "12345 Musterstadt",
    "A13": "",
    "A14": "",
    "E10": "Nach: Beispielstr. 2",
    "E11": "12345 Musterstadt",
    "E12": "",
    "E14": "Umzugstermin: 01.01.2026",
    "E15": "Beginn: 08:00",
}


def xml_text(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def sanitize_template(parts):
    """Ersetzt die Texte der Kundenzellen in sharedStrings.xml und entfernt Personennamen aus den Metadaten."""
    sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    sst = parts["xl/sharedStrings.xml"].decode("utf-8")
    sis = list(re.finditer(r"<si>(.*?)</si>", sst, re.S))
    new_text = {}
    for ref, text in CUSTOMER_CELLS.items():
        m = re.search(rf'<c r="{ref}"[^>]*t="s"[^>]*><v>(\d+)</v>', sheet)
        if m:
            new_text[int(m.group(1))] = text
    for idx in sorted(new_text, reverse=True):
        si = sis[idx]
        runs = re.findall(r"<r>.*?</r>", si.group(1), re.S)
        if runs and new_text[idx].startswith("Nach:"):
            # Rich Text "Nach:" (fett) + Adresse: Formatierung behalten, nur Adresse tauschen
            last = re.sub(r"<t[^>]*>.*?</t>", f'<t xml:space="preserve">{xml_text(new_text[idx][5:])}</t>', runs[-1], flags=re.S)
            body = "".join(runs[:-1]) + last
        else:
            body = f'<t xml:space="preserve">{xml_text(new_text[idx])}</t>'
        sst = sst[:si.start(1)] + body + sst[si.end(1):]
    parts["xl/sharedStrings.xml"] = sst.encode("utf-8")

    core = parts["docProps/core.xml"].decode("utf-8")
    core = re.sub(r"<dc:creator>.*?</dc:creator>", "<dc:creator>Philips Umzüge</dc:creator>", core)
    core = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "<cp:lastModifiedBy>Philips Umzüge</cp:lastModifiedBy>", core)
    parts["docProps/core.xml"] = core.encode("utf-8")

    # Lokalen Speicherpfad (x15ac:absPath) der Original-Datei entfernen
    wb = parts["xl/workbook.xml"].decode("utf-8")
    wb = re.sub(r"<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>).)*absPath.*?</mc:AlternateContent>", "", wb, flags=re.S)
    if "absPath" in wb:
        raise SystemExit("Lokaler Pfad (absPath) konnte nicht aus workbook.xml entfernt werden")
    parts["xl/workbook.xml"] = wb.encode("utf-8")


def build_template(path):
    z = zipfile.ZipFile(path)
    raw = {info.filename: z.read(info.filename) for info in z.infolist()}
    sanitize_template(raw)
    parts = {}
    for name, data in raw.items():
        parts[name] = base64.b64encode(data).decode("ascii")
    logo = parts.get("xl/media/image1.png")
    if not logo:
        raise SystemExit("Logo xl/media/image1.png nicht in der Vorlage gefunden")
    return parts, logo


def main():
    tpl_path = find_template()
    print("Vorlage:", tpl_path)
    data = {
        "modes": {"firma": build_firma(), "privat": build_privat()},
        "quote": {"transporterPrice": 140, "vatRate": 0.19},
        "templateName": tpl_path.name,
        "buildWarnings": warnings,
    }
    parts, logo = build_template(tpl_path)

    html = (HERE / "app_template.html").read_text(encoding="utf-8")
    for key, val in {
        "/*__APP_DATA__*/null": json.dumps(data, ensure_ascii=False),
        "/*__TEMPLATE_PARTS__*/null": json.dumps(parts),
        "__LOGO_DATA_URI__": "data:image/png;base64," + logo,
    }.items():
        if key not in html:
            raise SystemExit("Platzhalter fehlt im Template: " + key)
        html = html.replace(key, val)

    out = PROJECT / "Philips_Umzüge_Rechner.html"
    out.write_text(html, encoding="utf-8")
    for m in ("firma", "privat"):
        md = data["modes"][m]
        print(f"{m}: {len(md['catalog'])} Gegenstände, {len(md['extras'])} Zusatzleistungen")
    print("Geschrieben:", out, f"({out.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
