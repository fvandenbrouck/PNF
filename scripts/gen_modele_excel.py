#!/usr/bin/env python3
"""Génère le modèle Excel de dépôt groupé des propositions au PNF.

Le modèle reprend les champs du formulaire Grist (page « Formulaire »), avec les listes de choix
lues dans le document au moment de la génération. À régénérer si le formulaire ou les tables de
référence changent.

Usage (depuis la racine du dépôt) :
    python3 scripts/gen_modele_excel.py

Lecture seule sur Grist (via gen_formulaire_md.py). Dépendance : openpyxl.
Sortie : modele_propositions_PNF_2027-2028.xlsx à la racine du dépôt.

Structure de la feuille « Propositions » (contrat avec l'import) :
    ligne 1 : groupes de questions      ligne 2 : intitulés
    ligne 3 : aide à la saisie          ligne 4 : clés techniques (masquée)
    lignes 5 et suivantes : une proposition par ligne
Les clés sont les identifiants de colonnes de Propositions_PNF ; chaque orientation du schéma
directeur a sa colonne, de clé « Orientation_SD:<valeur> ».
"""
import datetime
import json
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.dont_write_bytecode = True
import gen_formulaire_md as g  # noqa: E402  (lit le formulaire et les tables de référence)

OUT = g.ROOT / "modele_propositions_PNF_2027-2028.xlsx"
NB_LIGNES = 200
FIRST = 5
LAST = FIRST + NB_LIGNES - 1

NAVY, WHITE, GREY = "000091", "FFFFFF", "F6F6F6"
BANDS = {"desc": "000091", "arr": "6A4C93", "sd": "18753C"}
BAND_TITLES = {"desc": "Description de l'action proposée",
               "arr": "Conformité à l'arrêté du 1er août 2023",
               "sd": "Principes et orientations du schéma directeur (facultatif) : mettre « x »"}
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WIDTHS = {"Email_contact": 30, "Structure_porteuse": 42, "Titre": 45, "Objectifs": 60, "Contexte": 45,
          "Action_SD": 70, "Public_cible": 50, "Public_autre": 30, "Lieu": 28, "Periode": 26,
          "Modalite": 30, "Format": 34}


# ----------------------------------------------------------------------------- lecture du formulaire
def walk(node, paragraphs, fields):
    if node["type"] == "Paragraph":
        paragraphs.append(node.get("text", ""))
    elif node["type"] == "Field" and node["leaf"] in g.FIELDS:
        fields.append(g.FIELDS[node["leaf"]])
    for child in node.get("children", []):
        walk(child, paragraphs, fields)


PARAGRAPHS, FIELDS = [], []
walk(g.LAYOUT, PARAGRAPHS, FIELDS)


def choices_of(f):
    ftype = f["type"]
    if ftype.startswith("Choice"):
        return json.loads(f["cwo"] or "{}").get("choices", [])
    if ftype == "Ref:Structures":
        return [f"{c} — {lib}" for c, lib in g.STRUCTURES]
    if ftype == "Ref:Modalites":
        return g.MODALITES
    if ftype == "Ref:Formats":
        return g.FORMATS
    if ftype == "Ref:Actions_axes":
        return [lib for _, lib in g.ACTIONS]
    if ftype.startswith("Ref:"):
        raise SystemExit(f"Type de référence non géré : {ftype}")
    return []


def columns():
    """Une entrée par colonne du classeur, dans l'ordre du formulaire."""
    cols = []
    for f in FIELDS:
        w = json.loads(f["cwo"] or "{}")
        label = g.clean(w.get("question") or f["label"])
        required = bool(w.get("formRequired"))
        key, ftype = f["colId"], f["type"]
        group = "arr" if key.startswith("Arr_") else "sd" if key == "Orientation_SD" else "desc"
        base = {"key": key, "label": label, "required": required, "group": group, "type": ftype, "choices": []}
        if key == "Orientation_SD":
            for value in choices_of(f):
                cols.append({**base, "key": f"{key}:{value}", "label": value, "choices": ["x"],
                             "hint": "« x » ou vide", "width": 18})
            continue
        if key == "Public_cible":
            base["hint"] = "Plusieurs publics possibles, séparés par un point-virgule (;). Liste dans la Notice."
        elif key == "Public_autre":
            base["label"] = "Précision si le public est « Autres »"
            base["hint"] = "Texte libre"
        else:
            base["choices"] = choices_of(f)
            if base["choices"]:
                base["hint"] = "Choisir dans la liste"
            elif ftype == "Int":
                base["hint"] = "Nombre entier"
            elif ftype == "Numeric":
                base["hint"] = "Nombre (heures)"
            else:
                base["hint"] = "Texte libre"
        base["hint"] += " — obligatoire" if required else " — facultatif"
        base["width"] = WIDTHS.get(key, 24)
        cols.append(base)
    return cols


# ----------------------------------------------------------------------------- texte de la notice
def plain(md):
    md = re.sub(r"<[^>]+>", "", md)
    md = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", md)
    md = md.replace("**", "").replace("*", "")
    return re.sub(r"^#+\s*", "", md).strip()


def notice_lines():
    """(style, texte, lien) ; les consignes sont reprises des textes du formulaire."""
    intro = next(p for p in PARAGRAPHS if "Qui peut proposer" in p)
    edito = re.search(r"\[éditorial\]\((https?://[^)]+)\)", intro)
    limite = re.search(r"avant le (\d{1,2}(?:er)? \S+ \d{4})", intro)
    if not edito or not limite:
        sys.exit("Lien vers l'éditorial ou date limite introuvable dans le texte du formulaire.")

    out = [("title", "Propositions d'actions de formation au PNF 2027-2028 — dépôt groupé", None),
           ("text", "Ce classeur s'adresse aux porteurs qui ont de nombreuses propositions à déposer. "
                    "Le formulaire en ligne reste la voie normale de dépôt.", None),
           ("head", "Mode d'emploi", None),
           ("bullet", "Saisissez une proposition par ligne dans la feuille « Propositions », à partir de la ligne 5.", None),
           ("bullet", "Les colonnes marquées d'un astérisque (*) sont obligatoires.", None),
           ("bullet", "Utilisez les menus déroulants quand ils existent. Si vous collez des valeurs, vérifiez "
                      "qu'elles figurent exactement dans la liste.", None),
           ("bullet", "Public cible : indiquez les publics séparés par un point-virgule (;), de préférence avec les "
                      "libellés ci-dessous. Un public non reconnu sera classé dans « Autres ».", None),
           ("bullet", "Ne modifiez pas les intitulés, n'ajoutez pas et ne déplacez pas de colonnes ou de feuilles.", None),
           ("bullet", f"Transmettez le fichier à la personne qui vous a communiqué ce modèle, avant le "
                      f"{limite.group(1)}, délai de rigueur.", None),
           ("bullet", "Ne déposez pas aussi dans le formulaire en ligne une proposition qui figure dans ce fichier.", None),
           ("link", "Lire l'éditorial du PNF 2027-2028 (PDF)", edito.group(1))]

    for block in intro.split("\n"):
        text = block.strip()
        if not text or text.startswith("# "):
            continue
        # consignes propres au formulaire en ligne : remplacées par le mode d'emploi ci-dessus
        if "Avant de commencer le dépôt" in text or "Déposer la proposition dans le formulaire" in text:
            continue
        if "Pour limiter les données personnelles" in text or "Protection de vos données" in text:
            continue
        bullet = text.startswith("- ")
        out.append(("bullet" if bullet else "text", plain(text[2:] if bullet else text), None))

    def block(title, marker):
        p = next((p for p in PARAGRAPHS if marker in p), None)
        if p is None:
            return
        out.append(("head", title, None))
        p = re.sub(r"<details><summary>.*?</summary>", "", p, flags=re.S)
        p = p.replace("Les sept questions ci-dessous", "Les questions « arrêté » de la feuille « Propositions »")
        p = p.replace("Cochez ceux auxquelles", "Marquez d'un « x » ceux auxquels")
        for line in p.split("\n"):
            line = line.strip()
            if not plain(line):
                continue
            bullet = line.startswith("- ")
            out.append(("bullet" if bullet else "text", plain(line[2:] if bullet else line), None))

    block("Consignes pour le titre", "Consignes pour le titre")
    block("Consignes pour les objectifs", "Consignes pour rédiger les objectifs")
    block("Conformité à l'arrêté du 1er août 2023", "Toute action de formation proposée au PNF")
    block("Principes et orientations du schéma directeur", "Ce qui est encouragé")

    out.append(("head", "Publics proposés", None))
    current = None
    grouped = {}
    for cat, lib in g.PUBLICS:
        grouped.setdefault(cat or "Sans catégorie", []).append(lib)
    for current, libs in grouped.items():
        out.append(("bullet", f"{current} : {' ; '.join(libs)}", None))

    start = intro.find("**Protection de vos données personnelles.**")
    if start < 0:
        sys.exit("Mention sur les données personnelles introuvable dans le texte du formulaire.")
    out.append(("head", "Protection de vos données personnelles", None))
    for line in intro[start:].replace("**Protection de vos données personnelles.**", "").split("\n"):
        if plain(line):
            out.append(("text", plain(line), None))
    return out


def write_notice(ws):
    ws.title = "Notice"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 120
    for row, (style, text, link) in enumerate(notice_lines(), 1):
        cell = ws.cell(row, 1, ("•  " + text) if style == "bullet" else text)
        cell.alignment = Alignment(wrap_text=True, vertical="top", indent=1 if style == "bullet" else 0)
        lines = sum(max(1, -(-len(part) // 125)) for part in cell.value.split("\n"))
        height = 16 * lines + 3
        if style == "title":
            cell.font = Font(bold=True, size=15, color=NAVY)
            height = 28
        elif style == "head":
            cell.font = Font(bold=True, size=12, color=NAVY)
            cell.alignment = Alignment(vertical="bottom")
            height = 30
        elif style == "link":
            cell.hyperlink = link
            cell.font = Font(bold=True, color="0000EE", underline="single")
        ws.row_dimensions[row].height = height
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ----------------------------------------------------------------------------- classeur
def build():
    cols = columns()
    wb = Workbook()
    write_notice(wb.active)
    ws = wb.create_sheet("Propositions")
    lists = wb.create_sheet("Listes")

    list_refs = {}

    def list_ref(values):
        key = tuple(values)
        if key not in list_refs:
            letter = get_column_letter(len(list_refs) + 1)
            lists.column_dimensions[letter].width = 40
            for i, v in enumerate(values, 1):
                lists[f"{letter}{i}"] = v
            list_refs[key] = f"Listes!${letter}$1:${letter}${len(values)}"
        return list_refs[key]

    for i, col in enumerate(cols, 1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = col["width"]
        color = BANDS[col["group"]]

        head = ws.cell(2, i, col["label"] + (" *" if col["required"] else ""))
        head.font = Font(bold=True, color=WHITE)
        head.fill = PatternFill("solid", fgColor=color)
        hint = ws.cell(3, i, col["hint"])
        hint.font = Font(italic=True, size=9, color="555555")
        hint.fill = PatternFill("solid", fgColor=GREY)
        ws.cell(4, i, col["key"])
        for cell in (head, hint):
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = BORDER

        numeric = col["type"] in ("Int", "Numeric")
        for row in range(FIRST, LAST + 1):
            cell = ws.cell(row, i)
            cell.alignment = Alignment(wrap_text=True, vertical="top",
                                       horizontal="center" if col["group"] == "sd" else None)
            cell.border = BORDER
            if not numeric:
                # format Texte : une saisie commençant par « - » n'est pas prise pour une formule
                cell.number_format = "@"

        zone = f"{letter}{FIRST}:{letter}{LAST}"
        dv = None
        if col["choices"]:
            dv = DataValidation(type="list", formula1=list_ref(col["choices"]), allow_blank=True)
            dv.error = "Choisissez une valeur dans la liste déroulante."
        elif col["type"] == "Int":
            dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", allow_blank=True)
            dv.error = "Saisissez un nombre entier supérieur ou égal à 1."
        elif col["type"] == "Numeric":
            dv = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True)
            dv.error = "Saisissez un nombre d'heures supérieur à 0."
        if dv is not None:
            dv.errorTitle = "Valeur non acceptée"
            dv.showErrorMessage = True
            ws.add_data_validation(dv)
            dv.add(zone)

    # bandeau des groupes de questions (ligne 1)
    start = 0
    for i in range(1, len(cols) + 1):
        if i == len(cols) or cols[i]["group"] != cols[start]["group"]:
            group = cols[start]["group"]
            # le premier groupe est scindé après les colonnes figées, pour que le titre reste lisible
            first = start + 1
            if group == "desc":
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
                first = 4
            cell = ws.cell(1, 1 if group == "desc" else first, BAND_TITLES[group])
            cell.font = Font(bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=BANDS[group])
            if first <= i:
                if group == "desc":
                    ws.cell(1, first).fill = PatternFill("solid", fgColor=BANDS[group])
                ws.merge_cells(start_row=1, start_column=first, end_row=1, end_column=i)
            start = i

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 48
    ws.row_dimensions[3].height = 40
    ws.row_dimensions[4].hidden = True
    ws.freeze_panes = "D5"
    lists.sheet_state = "hidden"
    wb.properties.title = "Modèle de dépôt groupé — PNF 2027-2028"
    wb.properties.description = f"Généré le {datetime.date.today().isoformat()}"
    wb.properties.creator = "PNF"
    wb.save(OUT)
    return cols


if __name__ == "__main__":
    cols = build()
    print(f"{OUT.name} : {len(cols)} colonnes, {NB_LIGNES} lignes de saisie")
