#!/usr/bin/env python3
"""Génère des copies Markdown du formulaire Grist de recueil PNF (page « Formulaire »).

Deux versions, à partir du formulaire tel qu'il est enregistré dans le document :
  - relecture : pour les collaborateurs (repères techniques : type de réponse, colonne, bandeau) ;
  - porteurs  : fiche de préparation sans repères techniques, avec un espace « Votre réponse »
                sous chaque question, pour préparer la saisie avant d'ouvrir le formulaire.

Les listes déroulantes (structures, actions du schéma directeur, modalités, publics, choix)
sont développées à partir des tables de référence du document.

Usage (depuis la racine du dépôt) :
    python3 scripts/gen_formulaire_md.py            # les deux versions
    python3 scripts/gen_formulaire_md.py porteurs   # une seule version

Lecture seule sur Grist (GET et requêtes SQL uniquement). Identifiants lus dans ./.env.
Sorties à la racine du dépôt :
    formulaire_PNF_2027-2028_relecture.md
    fiche_preparation_PNF_2027-2028.md
"""
import datetime
import json
import re
import sys
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = {"relecture": ROOT / "formulaire_PNF_2027-2028_relecture.md",
       "porteurs": ROOT / "fiche_preparation_PNF_2027-2028.md"}

env = {}
for line in (ROOT / ".env").read_text().splitlines():
    line = line.strip()
    if line and not line.startswith("#") and "=" in line:
        k, v = line.split("=", 1)
        env[k] = v.strip().strip('"').strip("'")
BASE = f"{env['GRIST_SERVER']}/api/docs/{env['GRIST_DOC_ID']}"
HDR = {"Authorization": f"Bearer {env['GRIST_API_KEY']}"}


def get(path):
    return json.loads(urllib.request.urlopen(urllib.request.Request(BASE + path, headers=HDR)).read())


def sql(q):
    return [r["fields"] for r in get("/sql?" + urllib.parse.urlencode({"q": q}))["records"]]


def recs(table):
    return get(f"/tables/{table}/records")["records"]


def clean(s):
    return re.sub(r"\s+", " ", (s or "").replace("\n", " ")).strip()


# ----------------------------------------------------------------------------- lecture du formulaire
sections = sql("select s.id, s.layoutSpec from _grist_Views_section s where s.parentKey='form'")
if len(sections) != 1:
    sys.exit(f"{len(sections)} formulaires trouvés dans le document : adapter le script.")
SID = sections[0]["id"]
LAYOUT = json.loads(sections[0]["layoutSpec"])
FIELDS = {r["fid"]: r for r in sql(
    "select f.id fid, c.colId, c.label, c.type, c.widgetOptions cwo "
    f"from _grist_Views_section_field f join _grist_Tables_column c on f.colRef=c.id where f.parentId={SID}")}

# ----------------------------------------------------------------------------- tables de référence
STRUCTURES = [(r["fields"]["code"], clean(r["fields"]["libelle"])) for r in recs("Structures")]
AXES = {r["id"]: (clean(r["fields"]["libelle"]), clean(r["fields"]["description"])) for r in recs("Axes")}
ACTIONS = [(r["fields"]["id_axe"], clean(r["fields"]["libelle"])) for r in recs("Actions_axes")]
MODALITES = [clean(r["fields"]["libelle"]) for r in recs("Modalites")]
FORMATS = [clean(r["fields"]["libelle"]) for r in recs("Formats")]
PUBLICS = [(clean(r["fields"]["categorie"]), clean(r["fields"]["libelle"])) for r in recs("Publics")]
DATE_LIMITE = next((r["fields"]["valeur"] for r in recs("Parametres_recueil")
                    if r["fields"]["cle"] == "date_limite"), "")


def date_fr(iso):
    try:
        return datetime.date.fromisoformat(iso).strftime("%d/%m/%Y")
    except ValueError:
        return iso


def expand(ftype, cwo):
    """Valeurs proposées par un champ, sous forme de lignes Markdown."""
    if ftype.startswith("Choice"):
        return ["- " + c for c in json.loads(cwo or "{}").get("choices", [])]
    if ftype == "Ref:Structures":
        return [f"- {c} — {lib}" for c, lib in STRUCTURES]
    if ftype == "Ref:Modalites":
        return ["- " + m for m in MODALITES]
    if ftype == "Ref:Formats":
        return ["- " + m for m in FORMATS]
    if ftype == "Ref:Actions_axes":
        lines, cur = [], None
        for ax, lib in ACTIONS:
            if ax != cur:
                cur = ax
                lines.append(f"- **{AXES[ax][0]} — {AXES[ax][1]}**")
            lines.append("  - " + lib)
        return lines
    if ftype == "RefList:Publics":
        lines, cur = [], object()
        for cat, lib in PUBLICS:
            if cat != cur:
                cur = cat
                lines.append(f"- **{cat or 'Sans catégorie'}**")
            lines.append("  - " + lib)
        return lines
    if ftype.startswith("Ref"):
        raise SystemExit(f"Type de référence non géré : {ftype}")
    return []


def count(ftype, lines):
    nested = ftype in ("Ref:Actions_axes", "RefList:Publics")
    return sum(1 for x in lines if x.startswith("  - ")) if nested else len(lines)


KIND_TEXT = {"Text": "Réponse libre", "Numeric": "Nombre", "Int": "Nombre entier", "Date": "Date"}


def kind_of(ftype, cwo):
    n_choices = len(json.loads(cwo or "{}").get("choices", [])) if ftype == "Choice" else 99
    if ftype in ("ChoiceList", "RefList:Publics"):
        return "Cases à cocher (plusieurs choix possibles)", "Cochez une ou plusieurs réponses dans la liste."
    if ftype == "Choice" and n_choices <= 3:
        return "Choix unique (boutons ou liste)", "Choisissez une seule réponse."
    if ftype.startswith("Choice") or ftype.startswith("Ref:"):
        return "Liste déroulante (un seul choix)", "Choisissez une seule réponse dans la liste."
    return KIND_TEXT.get(ftype, ftype), "Réponse à saisir."


# ----------------------------------------------------------------------------- rendu
def render(mode):
    out, counter, req_flags, unknown = [], [0], [], set()

    def field(fid):
        if fid not in FIELDS:   # champ supprimé (colonne retirée) mais encore cité dans la mise en page du formulaire
            print(f"Avertissement : le champ {fid} cité dans la mise en page n'existe plus ; ignoré.", file=sys.stderr)
            return
        f = FIELDS[fid]
        ftype = f["type"]
        counter[0] += 1
        required = bool(json.loads(f["cwo"] or "{}").get("formRequired"))
        req_flags.append((counter[0], required))
        kind, hint = kind_of(ftype, f["cwo"])
        star = " \\*" if required else ""
        out.append(f"#### {counter[0]}. {clean(f['label'])}{star}\n")
        if mode == "relecture":
            out.append(f"*Type de réponse : {kind}{' — obligatoire' if required else ' — facultatif'}*"
                       f" — colonne `{f['colId']}`\n")
        else:
            out.append(f"*{hint}{' Champ obligatoire.' if required else ' Champ facultatif.'}*\n")
        lines = expand(ftype, f["cwo"])
        if lines:
            out.append(f"Valeurs proposées ({count(ftype, lines)}) :\n")
            out.extend(lines)
            out.append("")
        if mode == "porteurs":
            out.append("**Votre réponse :** …\n")

    def walk(node):
        t = node["type"]
        if t == "Paragraph":
            out.append(node.get("text", "") + "\n")
        elif t == "Field":
            field(node["leaf"])
        elif t == "Section":
            out.append("\n---\n")
            for ch in node["children"]:
                walk(ch)
        elif t == "Submit":
            out.append("\n---\n\n*[Bouton d'envoi du formulaire]*\n" if mode == "relecture" else "")
        else:
            if t not in ("Layout", "Columns"):
                unknown.add(t)
            for ch in node.get("children", []):
                walk(ch)

    walk(LAYOUT)
    today = datetime.date.today().isoformat()
    if mode == "relecture":
        head = (f"<!-- Copie de relecture du formulaire de recueil PNF 2027-2028 — générée depuis le document Grist "
                f"le {today} (scripts/gen_formulaire_md.py).\n     Les listes déroulantes sont développées. "
                f"Reflète l'état du formulaire au moment de la génération. -->\n\n"
                "> **Document de relecture** — copie du formulaire de dépôt des propositions d'actions de formation "
                "au PNF 2027-2028, avec les listes de choix développées. Les mentions « Type de réponse » et "
                "« colonne » sont des repères techniques destinés aux relecteurs ; elles n'apparaissent pas dans le "
                "formulaire. Un astérisque (\\*) signale un champ obligatoire.\n\n")
        body = "\n".join(out)
    else:
        limite = f" avant le {date_fr(DATE_LIMITE)}" if DATE_LIMITE else ""
        head = (f"<!-- Fiche de préparation générée le {today} (scripts/gen_formulaire_md.py). -->\n\n"
                "> **Fiche de préparation.** Ce document reprend le contenu du formulaire en ligne pour vous permettre "
                "de préparer vos réponses à l'avance. Il ne constitue pas un dépôt : votre proposition doit être "
                f"saisie dans le formulaire en ligne{limite}. Les champs marqués d'un astérisque (\\*) y sont "
                "obligatoires.\n\n")
        body = "\n".join(out)
    return head + body, counter[0], sum(1 for _, r in req_flags if r), unknown


if __name__ == "__main__":
    modes = sys.argv[1:] or ["relecture", "porteurs"]
    for m in modes:
        if m not in OUT:
            sys.exit(f"Mode inconnu : {m} (relecture | porteurs)")
        text, n, n_req, unknown = render(m)
        OUT[m].write_text(text)
        print(f"{m:10s} → {OUT[m].name} : {n} champs ({n_req} obligatoires)"
              + (f" ; éléments non gérés : {sorted(unknown)}" if unknown else ""))
