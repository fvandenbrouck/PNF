#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Contrôle des règles d'accès (ACL) et des accès du document Grist du PNF.

Lecture seule : le script n'écrit jamais dans Grist.

Deux usages :

  python3 scripts/controle_regles_acces.py --init
      Enregistre l'état actuel comme RÉFÉRENCE VALIDÉE (à faire une fois les
      règles relues, puis après chaque changement voulu : nouvelle table,
      publication du formulaire, etc.).

  python3 scripts/controle_regles_acces.py
      Compare l'état actuel à la référence et vérifie des garde-fous.
      Code de sortie : 0 = conforme, 1 = écart ou garde-fou violé, 2 = erreur.

Ce qui est comparé : règles par table et par colonnes (ordre inclus : l'ordre
des règles change leur effet), règles par défaut, liste des tables, rôles du
document (propriétaires, éditeurs, lecteurs).

Garde-fous vérifiés, indépendamment de la référence :
  - chaque table a au moins une règle et une règle de repli (colonnes *, sans condition) ;
  - aucune règle sans condition n'accorde de droit d'écriture (création,
    modification, suppression) à tous, sauf exceptions déclarées (--autoriser-creation-publique) ;
  - la modification du schéma est refusée aux non-propriétaires ;
  - aucun accès « tout le monde » au niveau du document ;
  - aucun compte n'a gagné un rôle plus élevé que dans la référence.

La référence est stockée hors du dépôt (par défaut dans
~/grist-backups/PNF/acl_reference.json, permissions 600). Les adresses
électroniques n'y figurent qu'en empreinte (SHA-256) ; les rapports les affichent masquées.

Configuration : GRIST_API_KEY, GRIST_DOC_ID, GRIST_SERVER (fichier .env du projet).
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_REFERENCE = Path.home() / "grist-backups" / "PNF" / "acl_reference.json"
ROLE_RANK = {"": 0, "none": 0, "viewers": 1, "editors": 2, "owners": 3}


# ----------------------------------------------------------------------------- configuration
def load_env(path: Path) -> None:
    """Lit un fichier .env simple (sans dépendance externe) sans écraser l'environnement."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def api_get(server: str, api_key: str, path: str) -> Any:
    req = urllib.request.Request(
        server + path, headers={"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Erreur HTTP {e.code} sur {path}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Erreur réseau sur {path} : {e}") from e


def records(server: str, key: str, doc: str, table: str) -> list[dict[str, Any]]:
    return api_get(server, key, f"/api/docs/{doc}/tables/{table}/records")["records"]


# ----------------------------------------------------------------------------- lecture de l'état
def mask(email: str) -> str:
    local, _, domain = email.partition("@")
    return f"{local[:2]}***@{domain}" if domain else email[:2] + "***"


def digest(email: str) -> str:
    return hashlib.sha256(email.strip().lower().encode("utf-8")).hexdigest()[:16]


def read_state(server: str, key: str, doc: str) -> dict[str, Any]:
    doc_q = urllib.parse.quote(doc, safe="")
    tables = {r["id"]: r["fields"]["tableId"] for r in records(server, key, doc_q, "_grist_Tables")}
    resources = {r["id"]: r["fields"] for r in records(server, key, doc_q, "_grist_ACLResources")}
    acl = records(server, key, doc_q, "_grist_ACLRules")

    grouped: dict[str, list[tuple[int, dict[str, str]]]] = {}
    for r in acl:
        f = r["fields"]
        res = resources.get(f["resource"], {})
        table = res.get("tableId") or "(défaut)"
        cols = ",".join(sorted(c for c in (res.get("colIds") or "").split(",") if c)) or "*"
        label = f"{table} [{cols}]"
        grouped.setdefault(label, []).append(
            (f.get("rulePos") or 0, {"formule": f.get("aclFormula") or "", "droits": f.get("permissionsText") or ""})
        )
    rules = {label: [x[1] for x in sorted(items, key=lambda it: it[0])] for label, items in sorted(grouped.items())}

    access = api_get(server, key, f"/api/docs/{doc_q}/access")
    people: dict[str, dict[str, Any]] = {}
    for u in access.get("users", []):
        email = (u.get("email") or "").strip()
        if not email:
            continue
        people[digest(email)] = {
            "masque": mask(email),
            "role": u.get("access") or "",
            "role_herite": u.get("parentAccess") or "",
            "public": email.lower().startswith(("everyone@", "anon@")),   # accès « tout le monde » / anonyme
        }

    return {
        "date": dt.datetime.now().isoformat(timespec="seconds"),
        "tables": sorted(tables.values()),
        "regles": rules,
        "roles": people,
        "role_maximal_herite": access.get("maxInheritedRole") or "",
    }


# ----------------------------------------------------------------------------- analyse
def parse_permissions(text: str) -> tuple[set[str], set[str]]:
    """'+R-CUD' -> ({'R'}, {'C','U','D'})."""
    granted: set[str] = set()
    denied: set[str] = set()
    sign = "+"
    for ch in text or "":
        if ch in "+-":
            sign = ch
        elif ch.isalpha():
            (granted if sign == "+" else denied).add(ch)
    return granted, denied


def effective_role(info: dict[str, Any]) -> str:
    own = info.get("role") or ""
    inherited = info.get("role_herite") or ""
    return own if ROLE_RANK.get(own, 0) >= ROLE_RANK.get(inherited, 0) else inherited


def guard_rails(state: dict[str, Any], allowed_public_create: set[str]) -> list[str]:
    problems: list[str] = []
    rules = state["regles"]

    by_table: dict[str, list[tuple[str, list[dict[str, str]]]]] = {}
    for label, items in rules.items():
        table = label.split(" [", 1)[0]
        by_table.setdefault(table, []).append((label, items))

    for table in state["tables"]:
        entries = by_table.get(table)
        if not entries:
            problems.append(f"Table « {table} » : aucune règle d'accès.")
            continue
        fallback = [it for label, items in entries if label.endswith("[*]") for it in items if it["formule"] == ""]
        if not fallback:
            problems.append(f"Table « {table} » : pas de règle de repli (colonnes *, sans condition).")

    for label, items in rules.items():
        table = label.split(" [", 1)[0]
        for it in items:
            granted, _ = parse_permissions(it["droits"])
            writes = granted & set("CUD")
            if it["formule"] == "" and writes:
                if table in allowed_public_create and writes == {"C"}:
                    continue
                problems.append(
                    f"{label} : règle sans condition qui accorde {''.join(sorted(writes))} à tous ({it['droits']})."
                )

    schema_ok = any(
        "OWNER" in it["formule"] and it["droits"].startswith("-") and "S" in it["droits"]
        for label, items in rules.items()
        if label.startswith(("(défaut)", "*SPECIAL", "* ["))
        for it in items
    )
    if not schema_ok:
        problems.append("La modification du schéma n'est pas refusée aux non-propriétaires (règle « -S » absente).")

    for info in state["roles"].values():
        if info.get("public") and effective_role(info):
            problems.append(f"Accès « tout le monde » au niveau du document (rôle « {effective_role(info)} »).")
    return problems


def diff_states(ref: dict[str, Any], cur: dict[str, Any]) -> list[str]:
    out: list[str] = []
    for t in sorted(set(cur["tables"]) - set(ref["tables"])):
        out.append(f"NOUVELLE TABLE : « {t} » (absente de la référence).")
    for t in sorted(set(ref["tables"]) - set(cur["tables"])):
        out.append(f"TABLE SUPPRIMÉE : « {t} ».")

    for label in sorted(set(cur["regles"]) | set(ref["regles"])):
        a, b = ref["regles"].get(label), cur["regles"].get(label)
        if a is None:
            out.append(f"RÈGLES AJOUTÉES sur {label} : " + " ; ".join(f"[{x['formule'] or 'tous'}] {x['droits']}" for x in b))
        elif b is None:
            out.append(f"RÈGLES RETIRÉES sur {label} : " + " ; ".join(f"[{x['formule'] or 'tous'}] {x['droits']}" for x in a))
        elif a != b:
            if sorted(map(json.dumps, a)) == sorted(map(json.dumps, b)):
                out.append(f"ORDRE MODIFIÉ sur {label} (même règles, autre ordre : l'effet peut changer).")
            else:
                out.append(f"RÈGLES MODIFIÉES sur {label} :")
                for x in a:
                    if x not in b:
                        out.append(f"    - [{x['formule'] or 'tous'}] {x['droits']}")
                for x in b:
                    if x not in a:
                        out.append(f"    + [{x['formule'] or 'tous'}] {x['droits']}")

    for h, info in cur["roles"].items():
        before = ref["roles"].get(h)
        role = effective_role(info)
        if before is None:
            out.append(f"NOUVEAU COMPTE : {info['masque']} avec le rôle « {role or 'aucun'} ».")
        elif ROLE_RANK.get(role, 0) > ROLE_RANK.get(effective_role(before), 0):
            out.append(f"RÔLE ÉLEVÉ : {info['masque']} passe de « {effective_role(before) or 'aucun'} » à « {role} ».")
    for h, info in ref["roles"].items():
        if h not in cur["roles"]:
            out.append(f"COMPTE RETIRÉ : {info['masque']} (rôle « {effective_role(info) or 'aucun'} »).")
    if ref.get("role_maximal_herite") != cur.get("role_maximal_herite"):
        out.append(
            f"RÔLE MAXIMAL HÉRITÉ modifié : « {ref.get('role_maximal_herite')} » -> « {cur.get('role_maximal_herite')} »."
        )
    return out


def summary_roles(state: dict[str, Any]) -> str:
    count: dict[str, int] = {}
    for info in state["roles"].values():
        r = effective_role(info) or "aucun"
        count[r] = count.get(r, 0) + 1
    return ", ".join(f"{n} {r}" for r, n in sorted(count.items(), key=lambda kv: -ROLE_RANK.get(kv[0], 0)))


# ----------------------------------------------------------------------------- principal
def main() -> int:
    ap = argparse.ArgumentParser(description="Contrôle des règles d'accès du document Grist du PNF (lecture seule).")
    ap.add_argument("--init", action="store_true", help="enregistre l'état actuel comme référence validée")
    ap.add_argument("--reference", default=str(DEFAULT_REFERENCE), help="fichier de référence (hors dépôt)")
    ap.add_argument("--autoriser-creation-publique", action="append", default=[], metavar="TABLE",
                    help="table dont la création sans condition est voulue (formulaire public publié) ; répétable")
    ap.add_argument("--json", action="store_true", help="sortie JSON (pour automatisation)")
    args = ap.parse_args()

    load_env(PROJECT_ROOT / ".env")
    server = (os.environ.get("GRIST_SERVER") or "").rstrip("/")
    key, doc = os.environ.get("GRIST_API_KEY", ""), os.environ.get("GRIST_DOC_ID", "")
    if not (server and key and doc):
        print("Erreur : GRIST_SERVER, GRIST_API_KEY et GRIST_DOC_ID doivent être définis (.env).", file=sys.stderr)
        return 2
    if not server.startswith("http"):
        server = "https://" + server

    try:
        state = read_state(server, key, doc)
    except Exception as e:  # noqa: BLE001
        print(f"Erreur : {e}", file=sys.stderr)
        return 2

    ref_path = Path(args.reference).expanduser()
    allowed = set(args.autoriser_creation_publique)
    problems = guard_rails(state, allowed)

    if args.init:
        ref_path.parent.mkdir(parents=True, exist_ok=True)
        ref_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        os.chmod(ref_path, 0o600)
        print(f"Référence enregistrée : {ref_path}")
        print(f"  {len(state['tables'])} tables, {sum(len(v) for v in state['regles'].values())} règles ; comptes : {summary_roles(state)}.")
        if problems:
            print("\nATTENTION : l'état enregistré comme référence viole des garde-fous :")
            for p in problems:
                print(f"  - {p}")
            return 1
        return 0

    if not ref_path.exists():
        print(f"Erreur : pas de référence ({ref_path}). Lancez d'abord avec --init.", file=sys.stderr)
        return 2
    ref = json.loads(ref_path.read_text(encoding="utf-8"))
    differences = diff_states(ref, state)

    if args.json:
        print(json.dumps({"ecarts": differences, "garde_fous": problems, "reference": ref.get("date"), "controle": state["date"]},
                         ensure_ascii=False, indent=2))
    else:
        print(f"Contrôle du {state['date']} — référence du {ref.get('date')}")
        print(f"  {len(state['tables'])} tables, {sum(len(v) for v in state['regles'].values())} règles ; comptes : {summary_roles(state)}.")
        print()
        print("Écarts par rapport à la référence :" if differences else "Aucun écart par rapport à la référence.")
        for d in differences:
            print(f"  - {d}" if not d.startswith("    ") else d)
        print()
        print("Garde-fous violés :" if problems else "Garde-fous : tous respectés.")
        for p in problems:
            print(f"  - {p}")
    return 1 if (differences or problems) else 0


if __name__ == "__main__":
    raise SystemExit(main())
