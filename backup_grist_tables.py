#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sauvegarde complète des tables d'un document Grist.

Sortie par défaut :
/Users/francoisvandenbrouck/grist-backups/PNF/table_backup/YYMMDD/

Ce dossier est volontairement hors de ~/Documents : macOS protège ce
répertoire ("Fichiers et dossiers") et un job lancé par launchd via un
binaire brut (/bin/zsh) n'obtient pas toujours cette autorisation de
façon fiable, même accordée dans les Réglages Système.

Pour chaque table :
- <table>.csv  : export lisible dans Excel/Numbers
- <table>.json : export brut conservant mieux listes, références et booléens
- manifest.json : résumé de la sauvegarde
- _structure/ : structure du document (sans elle, les données seules ne
  permettent pas de reconstruire le document)
    - columns/<table>.json : colonnes de chaque table (type, formule,
      options de widget, référence, etc.)
    - <table de métadonnées>.json : tables internes _grist_* (règles d'accès,
      pages, vues, widgets, champs, définition des tables et des colonnes)

Configuration minimale :
export GRIST_API_KEY="..."
export GRIST_DOC_ID="..."

Optionnel :
export GRIST_SERVER="https://docs.getgrist.com"
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
from dotenv import load_dotenv
load_dotenv()
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any


DEFAULT_BACKUP_ROOT = Path(
    "/Users/francoisvandenbrouck/grist-backups/PNF/table_backup"
)


def today_yymmdd() -> str:
    return dt.datetime.now().strftime("%y%m%d")


def now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def safe_filename(name: str) -> str:
    """Nom de fichier stable et lisible."""
    s = str(name).strip()
    s = re.sub(r"[^\w.\-À-ÖØ-öø-ÿ]+", "_", s, flags=re.UNICODE)
    s = s.strip("._")
    return s or "table"


def normalize_server(server: str) -> str:
    server = (server or "").strip().rstrip("/")
    if not server:
        raise ValueError("Serveur Grist vide.")
    if not server.startswith(("http://", "https://")):
        server = "https://" + server
    return server


def api_get_json(server: str, api_key: str, path: str) -> dict[str, Any]:
    url = f"{server}{path}"
    req = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Accept": "application/json",
        },
        method="GET",
    )

    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Erreur HTTP {e.code} sur {url}\n{body}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Erreur réseau sur {url}\n{e}") from e


def list_tables(server: str, api_key: str, doc_id: str) -> list[dict[str, Any]]:
    doc = urllib.parse.quote(doc_id, safe="")
    data = api_get_json(server, api_key, f"/api/docs/{doc}/tables")

    tables = data.get("tables")
    if not isinstance(tables, list):
        raise RuntimeError(
            "Réponse inattendue de l'API Grist pour la liste des tables. "
            "Clé 'tables' absente ou invalide."
        )

    return tables


def get_records(server: str, api_key: str, doc_id: str, table_id: str) -> list[dict[str, Any]]:
    doc = urllib.parse.quote(doc_id, safe="")
    table = urllib.parse.quote(table_id, safe="")
    data = api_get_json(server, api_key, f"/api/docs/{doc}/tables/{table}/records")

    records = data.get("records")
    if not isinstance(records, list):
        raise RuntimeError(
            f"Réponse inattendue de l'API Grist pour la table {table_id}. "
            "Clé 'records' absente ou invalide."
        )

    return records


def collect_columns(records: list[dict[str, Any]]) -> list[str]:
    columns: list[str] = []
    seen: set[str] = set()

    for record in records:
        fields = record.get("fields") or {}
        if not isinstance(fields, dict):
            continue
        for key in fields.keys():
            if key not in seen:
                seen.add(key)
                columns.append(key)

    return columns


def csv_value(v: Any) -> Any:
    """Conserve les structures complexes sous forme JSON dans le CSV."""
    if v is None:
        return ""
    if isinstance(v, (list, dict)):
        return json.dumps(v, ensure_ascii=False)
    return v


def write_table_csv(path: Path, records: list[dict[str, Any]]) -> None:
    columns = collect_columns(records)
    fieldnames = ["id", *columns]

    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()

        for record in records:
            fields = record.get("fields") or {}
            row = {"id": record.get("id")}
            for col in columns:
                row[col] = csv_value(fields.get(col))
            writer.writerow(row)


def write_table_json(path: Path, table_id: str, records: list[dict[str, Any]]) -> None:
    payload = {
        "table_id": table_id,
        "exported_at": now_iso(),
        "record_count": len(records),
        "records": records,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


# Tables internes de Grist qui portent la structure du document.
# Lecture réservée aux Owners : un échec est signalé, jamais ignoré.
METADATA_TABLES = [
    "_grist_Tables",
    "_grist_Tables_column",
    "_grist_ACLResources",
    "_grist_ACLRules",
    "_grist_Views",
    "_grist_Views_section",
    "_grist_Views_section_field",
    "_grist_Pages",
]


def get_columns(server: str, api_key: str, doc_id: str, table_id: str) -> list[dict[str, Any]]:
    doc = urllib.parse.quote(doc_id, safe="")
    table = urllib.parse.quote(table_id, safe="")
    data = api_get_json(server, api_key, f"/api/docs/{doc}/tables/{table}/columns")

    columns = data.get("columns")
    if not isinstance(columns, list):
        raise RuntimeError(
            f"Réponse inattendue de l'API Grist pour les colonnes de {table_id}."
        )

    return columns


def backup_structure(
    server: str,
    api_key: str,
    doc_id: str,
    out_dir: Path,
    table_ids: list[str],
) -> tuple[dict[str, Any], list[str]]:
    """Sauvegarde colonnes, règles d'accès, pages et widgets.

    Retourne (résumé pour le manifest, liste des échecs).
    """
    struct_dir = out_dir / "_structure"
    columns_dir = struct_dir / "columns"
    columns_dir.mkdir(parents=True, exist_ok=True)

    summary: dict[str, Any] = {"columns": {}, "metadata": {}}
    failures: list[str] = []

    for table_id in table_ids:
        try:
            columns = get_columns(server, api_key, doc_id, table_id)
            path = columns_dir / f"{safe_filename(table_id)}.json"
            path.write_text(
                json.dumps(
                    {
                        "table_id": table_id,
                        "exported_at": now_iso(),
                        "column_count": len(columns),
                        "columns": columns,
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
            summary["columns"][table_id] = len(columns)
        except Exception as e:
            failures.append(f"colonnes de {table_id} : {e}")

    for meta_id in METADATA_TABLES:
        print(f"Export structure : {meta_id}")
        try:
            records = get_records(server, api_key, doc_id, meta_id)
            write_table_json(struct_dir / f"{meta_id}.json", meta_id, records)
            summary["metadata"][meta_id] = len(records)
        except Exception as e:
            failures.append(f"{meta_id} : {e}")

    return summary, failures


def make_output_dir(root: Path, folder_name: str, force: bool) -> Path:
    out_dir = root / folder_name

    if out_dir.exists() and any(out_dir.iterdir()) and not force:
        raise FileExistsError(
            f"Le dossier existe déjà et n'est pas vide : {out_dir}\n"
            "Relance avec --force pour écraser les fichiers de sauvegarde du jour."
        )

    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def backup_all_tables(
    server: str,
    api_key: str,
    doc_id: str,
    backup_root: Path,
    folder_name: str,
    force: bool,
    with_structure: bool = True,
) -> tuple[Path, list[str]]:
    server = normalize_server(server)
    out_dir = make_output_dir(backup_root, folder_name, force)

    print(f"Serveur Grist : {server}")
    print(f"Document      : {doc_id}")
    print(f"Dossier       : {out_dir}")

    tables = list_tables(server, api_key, doc_id)
    if not tables:
        raise RuntimeError("Aucune table trouvée dans ce document.")

    manifest: dict[str, Any] = {
        "created_at": now_iso(),
        "server": server,
        "doc_id": doc_id,
        "backup_dir": str(out_dir),
        "table_count": len(tables),
        "tables": [],
    }

    for table in tables:
        table_id = table.get("id")
        if not table_id:
            print("Table ignorée : identifiant absent.")
            continue

        table_id = str(table_id)
        filename = safe_filename(table_id)

        print(f"Export table : {table_id}")

        records = get_records(server, api_key, doc_id, table_id)

        csv_path = out_dir / f"{filename}.csv"
        json_path = out_dir / f"{filename}.json"

        write_table_csv(csv_path, records)
        write_table_json(json_path, table_id, records)

        manifest["tables"].append(
            {
                "table_id": table_id,
                "record_count": len(records),
                "csv": csv_path.name,
                "json": json_path.name,
            }
        )

    failures: list[str] = []
    if with_structure:
        table_ids = [t["table_id"] for t in manifest["tables"]]
        manifest["structure"], failures = backup_structure(
            server, api_key, doc_id, out_dir, table_ids
        )
        manifest["structure_failures"] = failures

    manifest_path = out_dir / "manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print()
    print(f"Sauvegarde terminée : {out_dir}")
    print(f"Manifest : {manifest_path}")
    if failures:
        print("ATTENTION : structure incomplète :", file=sys.stderr)
        for failure in failures:
            print(f"  - {failure}", file=sys.stderr)

    return out_dir, failures


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sauvegarde toutes les tables d'un document Grist en CSV et JSON."
    )

    parser.add_argument(
        "--server",
        default=os.getenv("GRIST_SERVER", "https://docs.getgrist.com"),
        help="Serveur Grist. Ex. https://docs.getgrist.com ou https://mon-equipe.getgrist.com",
    )
    parser.add_argument(
        "--doc-id",
        default=os.getenv("GRIST_DOC_ID", ""),
        help="Identifiant du document Grist. Peut aussi être fourni via GRIST_DOC_ID.",
    )
    parser.add_argument(
        "--api-key",
        default=os.getenv("GRIST_API_KEY", ""),
        help="Clé API Grist. Peut aussi être fournie via GRIST_API_KEY.",
    )
    parser.add_argument(
        "--backup-root",
        default=str(DEFAULT_BACKUP_ROOT),
        help="Dossier racine des sauvegardes.",
    )
    parser.add_argument(
        "--folder",
        default=today_yymmdd(),
        help="Nom du sous-dossier de sauvegarde. Par défaut : YYMMDD.",
    )
    parser.add_argument(
        "--no-structure",
        action="store_true",
        help="Ne sauvegarde que les données (sans colonnes, règles d'accès, pages ni widgets).",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Autorise l'écriture dans un dossier de sauvegarde déjà existant.",
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if not args.api_key:
        print("Erreur : GRIST_API_KEY manquant.", file=sys.stderr)
        print('Exemple : export GRIST_API_KEY="..."', file=sys.stderr)
        return 2

    if not args.doc_id:
        print("Erreur : GRIST_DOC_ID manquant.", file=sys.stderr)
        print('Exemple : export GRIST_DOC_ID="..."', file=sys.stderr)
        return 2

    try:
        _, failures = backup_all_tables(
            server=args.server,
            api_key=args.api_key,
            doc_id=args.doc_id,
            backup_root=Path(args.backup_root).expanduser(),
            folder_name=args.folder,
            force=args.force,
            with_structure=not args.no_structure,
        )
        return 1 if failures else 0
    except Exception as e:
        print(f"Erreur : {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
