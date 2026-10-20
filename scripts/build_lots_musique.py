#!/usr/bin/env python3
"""Génère l'Excel des 250 lots musique (RM45–60, en stock, livrables en Malaisie).

Usage :
    python scripts/build_lots_musique.py --research <dossier_json> --out <fichier.xlsx> [--data-dir data]

Le dossier de recherche contient des tableaux JSON d'objets avec les clés de FIELDS
(vinyl_*.json, cd*.json, funko*.json, books*.json, other*.json). Les doublons (même lien)
sont fusionnés, les prix hors fourchette écartés, puis les quotas par catégorie appliqués.
Ce qui dépasse les quotas part dans l'onglet « Réserve » (remplaçants si rupture).
"""
import argparse
import csv
import glob
import json
import os
import re
from urllib.parse import urlparse

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

FIELDS = [
    "category", "name", "price_myr", "url", "store", "store_country",
    "stock_evidence", "shipping_note", "notes",
]
PRICE_MIN, PRICE_MAX = 45.0, 60.0
TOTAL = 250
# Ordre = ordre d'affichage. Les Funko prennent tout ce qui reste après les autres quotas.
QUOTAS = {"Vinyle": 88, "CD": 62, "Funko Pop": None, "Livre": 13, "Autre": 12}
CATEGORY_ALIASES = {
    "vinyl": "Vinyle", "vinyle": "Vinyle", "cd": "CD", "funko": "Funko Pop",
    "funko pop": "Funko Pop", "livre": "Livre", "book": "Livre", "autre": "Autre", "other": "Autre",
}

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
CAT_FILLS = {
    "Vinyle": "E2EFDA", "CD": "DDEBF7", "Funko Pop": "FCE4D6", "Livre": "FFF2CC", "Autre": "EDE2F6",
}
THIN = Side(style="thin", color="BFBFBF")


def norm_category(value):
    return CATEGORY_ALIASES.get(str(value or "").strip().lower(), str(value or "").strip())


def norm_url(url):
    p = urlparse(url.strip())
    return f"{p.netloc.lower().removeprefix('www.')}{p.path.rstrip('/')}"


def norm_name(name):
    return re.sub(r"[^a-z0-9]+", " ", name.lower()).strip()


def load(research_dir):
    rows, seen_urls, seen_names = [], set(), set()
    for path in sorted(glob.glob(os.path.join(research_dir, "*.json"))):
        with open(path, encoding="utf-8") as f:
            items = json.load(f)
        for it in items:
            row = {k: it.get(k, "") for k in FIELDS}
            row["category"] = norm_category(row["category"])
            try:
                row["price_myr"] = round(float(row["price_myr"]), 2)
            except (TypeError, ValueError):
                continue
            if not (PRICE_MIN <= row["price_myr"] <= PRICE_MAX) or not row["url"]:
                continue
            key_url = norm_url(row["url"])
            key_name = (row["category"], norm_name(row["name"]))
            if key_url in seen_urls or key_name in seen_names:
                continue
            seen_urls.add(key_url)
            seen_names.add(key_name)
            rows.append(row)
    return rows


# Vinyles : artistes connus du grand public en priorité (ce sont des lots).
FAMOUS_VINYL = [
    "elton john", "u2", "beatles", "simon and garfunkel", "nirvana", "clash", "coldplay",
    "neil diamond", "barry manilow", "jackson browne", "rod stewart", "lionel richie", "commodores",
    "donna summer", "human league", "yes –", "bob seger", "steve winwood", "adam and the ants",
    "olivia newton-john", "roger taylor", "agnetha", "frank sinatra", "smokey robinson",
    "temptations", "stylistics", "three degrees", "barbra streisand", "pat benatar", "chicago",
    "manhattan transfer", "tlc", "en vogue", "soul ii soul", "salt ‘n’ pepa", "bobby brown",
    "thompson twins", "basement jaxx", "dead or alive", "ziggy marley", "brand new heavies",
    "the shamen", "bomb the bass", "t’pau", "crystal waters", "la toya jackson", "teenage fanclub",
    "bad brains", "los lobos", "zee avi", "fall, the", "wedding present", "manfred mann",
    "sharifah aini", "bay city rollers", "sheena easton", "kris kristofferson", "connie francis",
    "count basie", "chet atkins", "polnareff", "demis roussos", "gilbert o'sullivan", "gap band",
    "matt bianco", "johnny hates jazz", "dennis deyoung", "yvonne elliman", "terry jacks",
    "boz scaggs", "bachman", "nolans", "rick springfield", "joan armatrading", "tomita",
    "motoharu sano", "mal waldron", "buddy rich", "sidney bechet", "liberace", "martin denny",
    "teddy pendergrass", "jon anderson", "paul mauriat", "henry mancini", "herbie mann",
    "guru josh", "altern 8", "gorilla biscuits", "kids from fame", "melanie", "ray conniff",
]
# « Autre » : sélection variée, objets les plus parlants pour un lot musique d'abord.
AUTRE_PREFERRED = [
    "kurt cobain", "record display frame", "snoopy headphones", "blackpink polaroid",
    "jack white", "johnny cash", "k-pop demon hunters 4 inch", "rose plante music box",
    "beyonce", "behringer", "snoopy blue cap", "bob dylan pop art",
    "fanlight emblem", "cassette head cleaner", "frank sinatra pop art", "dua lipa",
    "paul stanley", "twice", "stray kids", "mitski \"the land", "godflesh",
]


def autre_rank(r):
    name = r["name"].lower()
    for i, k in enumerate(AUTRE_PREFERRED):
        if k in name:
            return i
    return len(AUTRE_PREFERRED)


# Mots qui signalent un article moins pertinent (relevé par la recherche dans les notes).
WEAK_MARKERS = ["[secondary pick]", "lower prior", "text-led", "coming-soon"]


def priority(r):
    """0 = à prendre en premier."""
    text = f"{r['name']} {r['notes']}".lower()
    if any(m in text for m in WEAK_MARKERS):
        return 2
    if r["category"] == "Vinyle":
        return 0 if any(a in r["name"].lower() for a in FAMOUS_VINYL) else 1
    if r["category"] == "Funko Pop":
        # Pop! standard d'abord ; Vinyl Gold, Bitty, crossovers ensuite
        return 1 if any(m in text for m in ["vinyl gold", "bitty", "crossover", "tinytan"]) else 0
    return 0


def drop_duplicates(rows):
    """Même produit vendu par 2 boutiques (marqué DUPLICATE) : garde la boutique qui en vend le plus."""
    per_store = {}
    for r in rows:
        per_store[(r["category"], r["store"])] = per_store.get((r["category"], r["store"]), 0) + 1
    best = {}
    for r in rows:
        if "DUPLICATE" not in r["notes"]:
            continue
        m = re.search(r"same figure \(([^)]*)\)", r["notes"])
        key = (r["category"], m.group(1) if m else r["name"])
        if key not in best or per_store[(r["category"], r["store"])] > per_store[(r["category"], best[key]["store"])]:
            best[key] = r
    keep = {id(r) for r in best.values()}
    return [r for r in rows if "DUPLICATE" not in r["notes"] or id(r) in keep]


def select(rows):
    """Applique les quotas ; artistes connus puis boutiques malaisiennes en priorité."""
    rows = drop_duplicates(rows)

    def rank(r):
        return (priority(r), 0 if str(r["store_country"]).lower() == "malaysia" else 1,
                r["store"], r["name"])

    by_cat = {c: sorted([r for r in rows if r["category"] == c], key=rank) for c in QUOTAS}
    by_cat["Autre"].sort(key=autre_rank)
    chosen = {}
    for cat, quota in QUOTAS.items():
        if quota is not None:
            chosen[cat] = by_cat[cat][:quota]
    # Funko : le reste, au maximum
    chosen["Funko Pop"] = by_cat["Funko Pop"][: TOTAL - sum(len(v) for v in chosen.values())]
    # Manque éventuel : on complète avec les « Autres », puis livres, puis vinyles, puis CD en réserve
    for cat in ["Autre", "Livre", "Vinyle", "CD", "Funko Pop"]:
        missing = TOTAL - sum(len(v) for v in chosen.values())
        if missing <= 0:
            break
        extra = by_cat[cat][len(chosen[cat]):][:missing]
        chosen[cat] = chosen[cat] + extra
    selected = [r for cat in QUOTAS for r in chosen[cat]]
    picked = {id(r) for r in selected}
    reserve = [r for cat in QUOTAS for r in by_cat[cat] if id(r) not in picked]
    return selected, reserve


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


HEADERS = ["N°", "Catégorie", "Produit", "Prix (RM)", "Lien", "Boutique", "Pays boutique",
           "Livraison Malaisie", "Stock vérifié", "Notes"]
WIDTHS = [6, 12, 52, 11, 50, 26, 15, 34, 34, 44]


def write_items(ws, items, table_name):
    ws.append(HEADERS)
    style_header(ws, 1, len(HEADERS))
    for i, r in enumerate(items, start=1):
        ws.append([i, r["category"], r["name"], r["price_myr"], r["url"], r["store"],
                   r["store_country"], r["shipping_note"], r["stock_evidence"], r["notes"]])
        row = ws.max_row
        fill = PatternFill("solid", fgColor=CAT_FILLS.get(r["category"], "FFFFFF"))
        for c in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = Font(name=FONT, size=10)
            cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
            cell.alignment = Alignment(vertical="top", wrap_text=c in (3, 8, 9, 10))
        ws.cell(row=row, column=2).fill = fill
        ws.cell(row=row, column=4).number_format = '"RM "#,##0.00'
        link = ws.cell(row=row, column=5)
        link.hyperlink = r["url"]
        link.font = Font(name=FONT, size=10, color="0563C1", underline="single")
    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2"
    if items:
        ref = f"A1:{get_column_letter(len(HEADERS))}{ws.max_row}"
        tab = Table(displayName=table_name, ref=ref)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        ws.add_table(tab)


def write_summary(ws, n_lots, n_reserve, date):
    ws["A1"] = "Lots musique – 250 lots entre RM45 et RM60"
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    ws["A2"] = f"Prix et stock relevés le {date} directement sur les boutiques (lien dans l'onglet Lots)."
    ws["A2"].font = Font(name=FONT, italic=True, size=10)
    head = ["Catégorie", "Objectif (%)", "Nb lots", "% réel", "Prix moyen (RM)", "Budget (RM)", "En réserve"]
    for c, h in enumerate(head, start=1):
        ws.cell(row=4, column=c, value=h)
    style_header(ws, 4, len(head))
    targets = {"Vinyle": 0.35, "CD": 0.25, "Funko Pop": None, "Livre": 0.05, "Autre": None}
    last = n_lots + 1
    r = 5
    for cat, target in targets.items():
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=target if target is not None else "reste")
        ws.cell(row=r, column=3, value=f'=COUNTIF(Lots!$B$2:$B${last},A{r})')
        ws.cell(row=r, column=4, value=f"=IF($C$10=0,0,C{r}/$C$10)")
        ws.cell(row=r, column=5, value=f'=IFERROR(AVERAGEIF(Lots!$B$2:$B${last},A{r},Lots!$D$2:$D${last}),0)')
        ws.cell(row=r, column=6, value=f'=SUMIF(Lots!$B$2:$B${last},A{r},Lots!$D$2:$D${last})')
        ws.cell(row=r, column=7, value=f'=COUNTIF(Réserve!$B$2:$B${n_reserve + 1},A{r})')
        r += 1
    ws.cell(row=r, column=1, value="Total")
    ws.cell(row=r, column=3, value="=SUM(C5:C9)")
    ws.cell(row=r, column=4, value="=SUM(D5:D9)")
    ws.cell(row=r, column=5, value=f"=IFERROR(AVERAGE(Lots!$D$2:$D${last}),0)")
    ws.cell(row=r, column=6, value="=SUM(F5:F9)")
    ws.cell(row=r, column=7, value="=SUM(G5:G9)")
    for row in ws.iter_rows(min_row=5, max_row=r, max_col=len(head)):
        for cell in row:
            cell.font = Font(name=FONT, bold=cell.row == r)
            cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
        row[1].number_format = "0%"
        row[3].number_format = "0.0%"
        row[4].number_format = '"RM "#,##0.00'
        row[5].number_format = '"RM "#,##0.00'
    ws.cell(row=r + 1, column=1, value="Prix min / max")
    ws.cell(row=r + 1, column=3, value=f"=MIN(Lots!$D$2:$D${last})")
    ws.cell(row=r + 1, column=4, value=f"=MAX(Lots!$D$2:$D${last})")
    for c in (3, 4):
        ws.cell(row=r + 1, column=c).number_format = '"RM "#,##0.00'
    for c in (1, 3, 4):
        ws.cell(row=r + 1, column=c).font = Font(name=FONT, italic=True)

    notes = [
        "Méthode",
        "• Prix = prix article affiché sur la boutique, hors frais de port. Fourchette appliquée : RM45 ≤ prix ≤ RM60.",
        "• Stock : disponibilité lue dans les données produit de la boutique (variante « disponible »), pas de précommande.",
        "• Boutiques malaisiennes en priorité ; les boutiques étrangères (ex. Singapour) indiquent la livraison et la conversion dans « Notes ».",
        "• CD : uniquement des albums cultes / classiques. Funko Pop : figurines de musiciens (Pop! Rocks, albums…).",
        "• Vinyles : quasi tous d'occasion (le neuf commence vers RM70 en Malaisie). Mudah.my = particuliers : confirmer avec le vendeur avant de payer.",
        "• Funko Pop : aucune boutique malaisienne sous RM60 (RM65–70 en MY) → boutiques étrangères, surtout Fugitive Toys (USA). Frais de port en plus : grouper la commande.",
        "• Onglet « Réserve » : produits valides en plus, pour remplacer un article épuisé.",
        "• Les stocks bougent vite : revérifier la disponibilité au moment de commander.",
    ]
    for i, t in enumerate(notes):
        cell = ws.cell(row=r + 3 + i, column=1, value=t)
        cell.font = Font(name=FONT, bold=i == 0, size=10)
    for col, w in zip("ABCDEFG", [16, 14, 10, 10, 16, 14, 12]):
        ws.column_dimensions[col].width = w


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--research", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--data-dir")
    ap.add_argument("--date", default="20/10/2026")
    args = ap.parse_args()

    rows = load(args.research)
    selected, reserve = select(rows)

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Résumé"
    write_items(wb.create_sheet("Lots"), selected, "Lots")
    write_items(wb.create_sheet("Réserve"), reserve, "Reserve")
    write_summary(ws_sum, len(selected), len(reserve), args.date)
    wb.save(args.out)

    if args.data_dir:
        os.makedirs(args.data_dir, exist_ok=True)
        with open(os.path.join(args.data_dir, "lots_musique.json"), "w", encoding="utf-8") as f:
            json.dump({"lots": selected, "reserve": reserve}, f, ensure_ascii=False, indent=1)
        with open(os.path.join(args.data_dir, "lots_musique.csv"), "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=FIELDS)
            w.writeheader()
            w.writerows(selected)

    counts = {c: sum(1 for r in selected if r["category"] == c) for c in QUOTAS}
    print(f"{len(rows)} produits valides -> {len(selected)} lots, {len(reserve)} en réserve : {counts}")


if __name__ == "__main__":
    main()
