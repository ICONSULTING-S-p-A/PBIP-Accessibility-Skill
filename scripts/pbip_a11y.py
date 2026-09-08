#!/usr/bin/env python3
"""
Motore della skill "pbip-accessibility": analizza e corregge le proprieta' di
accessibilita' (alt text, tooltip, tab order) di un progetto Power BI in formato PBIP/PBIR.

Comandi
    scan   <progetto> [--output-dir DIR] [--print summary|full|json]
    apply  <progetto> [--confirm-all] [--issue-ids id1,id2,...] [--scan-id ID]
    issues <progetto> [--scan-id ID] [--category CAT] [--page NOME] [--only-fixable] [--limit N]

Lo scan e' sempre in sola lettura: produce un Excel di riepilogo e una cache di scan
fuori dal progetto. apply modifica i visual.json in-place con edit testuali minimi
e salva sempre in UTF-8 senza BOM.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import uuid
from datetime import datetime, timezone
from typing import Dict, List, Optional

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from jsonedit import apply_edits, path_exists, write_utf8_no_bom  # noqa: E402
from pbip_rules import (  # noqa: E402
    alt_text_issues,
    read_project,
    tab_order_issues,
    tooltip_issues,
)

CATEGORY_LABELS = {
    "altText_mancante": "altText mancante",
    "tooltip_mancante": "tooltip mancante",
    "tabOrder_da_assegnare": "tabOrder da assegnare",
    "proprieta_da_spostare": "proprieta' da spostare",
    "anomalia": "anomalie",
}


# ================================================================ cache di scan
#
# Fuori dalla cartella del progetto: la cache non deve sporcare il repo git del report.


def cache_base_dir() -> str:
    override = os.environ.get("PBIP_A11Y_CACHE_DIR")
    if override:
        return override
    local = os.environ.get("LOCALAPPDATA")
    base = os.path.join(local, "pbip-accessibility-mcp") if local else os.path.join(
        os.path.expanduser("~"), ".pbip-accessibility-mcp"
    )
    return os.path.join(base, "scans")


def project_key(project_path: str) -> str:
    normalized = os.path.abspath(project_path).lower().replace(os.sep, "/")
    return hashlib.sha1(normalized.encode("utf-8")).hexdigest()


def project_dir(project_path: str) -> str:
    return os.path.join(cache_base_dir(), project_key(project_path))


def fingerprint_file(abs_path: str) -> dict:
    st = os.stat(abs_path)
    with open(abs_path, "rb") as f:
        sha = hashlib.sha256(f.read()).hexdigest()
    return {"mtimeMs": st.st_mtime * 1000.0, "sha256": sha}


def save_scan(record: dict) -> str:
    d = project_dir(record["project_path"])
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, record["scan_id"] + ".json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(record, f, indent=2, ensure_ascii=False)
    return path


def list_scans(project_path: str) -> List[dict]:
    d = project_dir(project_path)
    if not os.path.isdir(d):
        return []
    out = []
    for fn in os.listdir(d):
        if not fn.endswith(".json"):
            continue
        try:
            with open(os.path.join(d, fn), encoding="utf-8") as f:
                out.append(json.load(f))
        except Exception:
            pass
    out.sort(key=lambda r: r.get("timestamp", ""), reverse=True)
    return out


def latest_scan(project_path: str) -> Optional[dict]:
    scans = list_scans(project_path)
    return scans[0] if scans else None


def get_scan_by_id(scan_id: str, project_path: Optional[str]) -> Optional[dict]:
    if project_path:
        p = os.path.join(project_dir(project_path), scan_id + ".json")
        if os.path.exists(p):
            with open(p, encoding="utf-8") as f:
                return json.load(f)
    base = cache_base_dir()
    if not os.path.isdir(base):
        return None
    for d in os.listdir(base):
        p = os.path.join(base, d, scan_id + ".json")
        if os.path.exists(p):
            with open(p, encoding="utf-8") as f:
                return json.load(f)
    return None


def stale_files(record: dict) -> List[str]:
    """File del progetto cambiati dopo lo scan: applicare fix su uno scan stantio e' pericoloso."""
    changed = []
    for rel, fp in record.get("fingerprints", {}).items():
        abs_path = os.path.join(record["project_path"], *rel.split("/"))
        if not os.path.exists(abs_path):
            changed.append(rel)
            continue
        st = os.stat(abs_path)
        if abs(st.st_mtime * 1000.0 - fp["mtimeMs"]) < 1e-6:
            continue
        with open(abs_path, "rb") as f:
            sha = hashlib.sha256(f.read()).hexdigest()
        if sha != fp["sha256"]:
            changed.append(rel)
    return changed


# ================================================================ Excel

DETAIL_SHEET = "Dettaglio issue"
DETAIL_HEADERS = [
    "Pagina",
    "Visual",
    "Tipo visual",
    "Categoria",
    "Valore attuale",
    "Valore proposto",
    "ID issue",
    "Auto-fixable",
    "Stato",
    "Motivo",
]


def _style_header(ws, row_idx: int, ncols: int) -> None:
    from openpyxl.styles import Border, Font, PatternFill, Side

    fill = PatternFill("solid", fgColor="FFD9E1F2")
    border = Border(bottom=Side(style="thin"))
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.border = border


def write_scan_excel(path: str, scan: dict, issues: List[dict]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    # ---- Foglio 1: Riepilogo
    ws = wb.active
    ws.title = "Riepilogo"
    for i, w in enumerate([34, 60, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.append(["Report accessibilita' PBIP"])
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.append([])
    ws.append(["Data scan", scan["timestamp"]])
    ws.append(["Progetto", scan["project_path"]])
    ws.append(["Scan ID", scan["scan_id"]])
    ws.append(["Totale issue", len(issues)])
    ws.append([])

    by_cat: Dict[str, int] = {}
    for i in issues:
        by_cat[i["category"]] = by_cat.get(i["category"], 0) + 1
    ws.append(["Categoria", "Conteggio"])
    _style_header(ws, ws.max_row, 2)
    for cat, n in by_cat.items():
        ws.append([cat, n])
    ws.append([])

    by_page: Dict[str, int] = {}
    for i in issues:
        by_page[i["page"]] = by_page.get(i["page"], 0) + 1
    ws.append(["Pagina", "Conteggio"])
    _style_header(ws, ws.max_row, 2)
    for page, n in by_page.items():
        ws.append([page, n])

    # ---- Foglio 2: Dettaglio issue
    ws2 = wb.create_sheet(DETAIL_SHEET)
    for i, w in enumerate([22, 26, 20, 22, 40, 60, 40, 12, 12, 50], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.append(DETAIL_HEADERS)
    _style_header(ws2, 1, len(DETAIL_HEADERS))
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = "A1:J1"
    for i in issues:
        ws2.append(
            [
                i["page"],
                i["visual_name"],
                i["visual_type"],
                i["category"],
                i.get("current_value") or "",
                i.get("proposed_value") or i.get("reason") or "",
                i["id"],
                "Sì" if i["auto_fixable"] else "No",
                "",
                "",
            ]
        )

    # ---- Foglio 3: Anomalie
    ws3 = wb.create_sheet("Anomalie")
    for i, w in enumerate([22, 26, 20, 90, 40], start=1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.append(["Pagina", "Visual", "Tipo visual", "Descrizione anomalia", "ID issue"])
    _style_header(ws3, 1, 5)
    ws3.freeze_panes = "A2"
    for i in issues:
        if i["auto_fixable"]:
            continue
        ws3.append([i["page"], i["visual_name"], i["visual_type"], i.get("reason") or i.get("proposed_value") or "", i["id"]])

    wb.save(path)


def update_excel_status(path: str, outcomes: Dict[str, dict]) -> None:
    """Aggiorna Stato/Motivo nello STESSO Excel dello scan: resta l'unica fonte di verita' nel tempo."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    if DETAIL_SHEET not in wb.sheetnames:
        raise RuntimeError(f'Foglio "{DETAIL_SHEET}" non trovato in {path}')
    ws = wb[DETAIL_SHEET]
    id_col = stato_col = motivo_col = None
    for c in range(1, ws.max_column + 1):
        v = str(ws.cell(row=1, column=c).value or "").strip()
        if v == "ID issue":
            id_col = c
        elif v == "Stato":
            stato_col = c
        elif v == "Motivo":
            motivo_col = c
    if not (id_col and stato_col and motivo_col):
        raise RuntimeError(f'Colonne "ID issue"/"Stato"/"Motivo" non trovate nel foglio "{DETAIL_SHEET}"')
    for r in range(2, ws.max_row + 1):
        issue_id = str(ws.cell(row=r, column=id_col).value or "")
        outcome = outcomes.get(issue_id)
        if not outcome:
            continue
        ws.cell(row=r, column=stato_col).value = outcome["stato"]
        ws.cell(row=r, column=motivo_col).value = outcome.get("motivo") or ""
    wb.save(path)


# ================================================================ scan


def issue_view(i: dict) -> dict:
    view = {
        k: i.get(k)
        for k in (
            "id",
            "page",
            "visual_name",
            "visual_type",
            "category",
            "current_value",
            "proposed_value",
            "target_path",
            "auto_fixable",
        )
    }
    if i.get("reason"):
        view["reason"] = i["reason"]
    return view


def cmd_scan(args) -> int:
    project = read_project(args.project_path)
    timestamp = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    scan_id = str(uuid.uuid4())

    issues: List[dict] = []
    for idx, msg in enumerate(project.schema_anomalies):
        issues.append(
            {
                "id": f"anomalia-progetto-{idx + 1}",
                "page": "—",
                "page_name": "—",
                "visual_name": "—",
                "visual_type": "—",
                "category": "anomalia",
                "current_value": None,
                "proposed_value": None,
                "target_path": "",
                "file": "",
                "auto_fixable": False,
                "reason": msg,
                "edits": [],
            }
        )

    for page in project.pages:
        issues.extend(tab_order_issues(page))
        issues.extend(alt_text_issues(page, project))
        issues.extend(tooltip_issues(page, project))

    excel_dir = os.path.abspath(args.output_dir) if args.output_dir else project.project_path
    os.makedirs(excel_dir, exist_ok=True)
    stamp = timestamp.replace("-", "").replace(":", "").split(".")[0].replace("Z", "").replace("T", "-")
    excel_path = os.path.join(excel_dir, f"accessibilita_{project.report_name}_{stamp}.xlsx")
    write_scan_excel(excel_path, {"scan_id": scan_id, "timestamp": timestamp, "project_path": project.project_path}, issues)

    fingerprints = {}
    for rel in project.scanned_files:
        abs_path = os.path.join(project.project_path, *rel.split("/"))
        if os.path.exists(abs_path):
            fingerprints[rel] = fingerprint_file(abs_path)

    record = {
        "scan_id": scan_id,
        "timestamp": timestamp,
        "project_path": project.project_path,
        "excel_path": excel_path,
        "fingerprints": fingerprints,
        "issues": issues,
    }
    cache_path = save_scan(record)

    by_type: Dict[str, int] = {}
    for i in issues:
        key = "anomalie" if i["category"] == "anomalia" else i["category"]
        by_type[key] = by_type.get(key, 0) + 1

    result = {
        "scan_id": scan_id,
        "timestamp": timestamp,
        "project_path": project.project_path,
        "report_name": project.report_name,
        "pbir_version": project.pbir_version,
        "pages": [{"name": p.name, "display_name": p.display_name, "visuals": len(p.visuals)} for p in project.pages],
        "total_issues": len(issues),
        "issues_by_type": by_type,
        "excel_path": excel_path,
        "scan_cache": cache_path,
    }

    if args.print == "json":
        result["issues"] = [issue_view(i) for i in issues]
        print(json.dumps(result, indent=2, ensure_ascii=False))
        return 0

    _print_scan_summary(result, issues, full=(args.print == "full"))
    return 0


def _print_scan_summary(result: dict, issues: List[dict], full: bool) -> None:
    print(f"SCAN OK  scan_id={result['scan_id']}")
    print(f"Progetto: {result['project_path']}")
    print(f"Report: {result['report_name']}  (definition.pbir version: {result['pbir_version']})")
    print(f"Excel:  {result['excel_path']}")
    print(f"Pagine: {len(result['pages'])}   Issue totali: {result['total_issues']}")
    print()
    print("Issue per categoria:")
    for k, n in result["issues_by_type"].items():
        print(f"  - {CATEGORY_LABELS.get(k, k)}: {n}")
    print()
    print("Issue per pagina (auto-fixable / anomalie):")
    for p in result["pages"]:
        page_issues = [i for i in issues if i["page_name"] == p["name"]]
        fix = sum(1 for i in page_issues if i["auto_fixable"])
        anom = len(page_issues) - fix
        print(f"  - {p['display_name']}: {fix} / {anom}   ({p['visuals']} visuals)")

    anomalies = [i for i in issues if not i["auto_fixable"]]
    if anomalies:
        print()
        print(f"Anomalie da valutare a mano ({len(anomalies)}):")
        for a in anomalies:
            print(f"  [{a['id']}] {a['page']} / {a['visual_name']} ({a['visual_type']}): {a.get('reason')}")

    if full:
        print()
        print("Dettaglio issue auto-fixable:")
        for i in issues:
            if not i["auto_fixable"]:
                continue
            print(f"  [{i['category']}] {i['page']} / {i['visual_name']} ({i['visual_type']})")
            print(f"      {i['target_path']}: {i.get('current_value') or '—'}  ->  {i.get('proposed_value')}")


# ================================================================ issues (query sulla cache)


def cmd_issues(args) -> int:
    scan = get_scan_by_id(args.scan_id, args.project_path) if args.scan_id else latest_scan(args.project_path)
    if not scan:
        print("Nessuno scan trovato: esegui prima il comando scan.", file=sys.stderr)
        return 2
    issues = scan["issues"]
    if args.category:
        cats = set(args.category.split(","))
        issues = [i for i in issues if i["category"] in cats]
    if args.page:
        needle = args.page.lower()
        issues = [i for i in issues if needle in (i["page"] or "").lower()]
    if args.only_fixable:
        issues = [i for i in issues if i["auto_fixable"]]
    if args.limit:
        issues = issues[: args.limit]
    print(json.dumps({"scan_id": scan["scan_id"], "count": len(issues), "issues": [issue_view(i) for i in issues]}, indent=2, ensure_ascii=False))
    return 0


# ================================================================ apply


def _apply_issue(scan: dict, issue: dict, outcome: dict) -> None:
    if not issue["auto_fixable"]:
        outcome["skipped"].append(
            {"id": issue["id"], "reason": issue.get("reason") or "issue non auto-fixable, richiede intervento manuale"}
        )
        return
    if not issue.get("file"):
        outcome["failed"].append({"id": issue["id"], "reason": "issue senza file associato"})
        return
    abs_path = os.path.join(scan["project_path"], *issue["file"].split("/"))
    if not os.path.exists(abs_path):
        outcome["failed"].append({"id": issue["id"], "reason": f"file non trovato: {issue['file']}"})
        return
    try:
        with open(abs_path, encoding="utf-8") as f:
            text = f.read()
        # Le precondizioni si rivalutano ADESSO, non allo scan: nel frattempo l'utente
        # potrebbe aver scritto a mano un altText o un tooltip che non va sovrascritto.
        for pre in issue.get("preconditions") or []:
            if pre["kind"] == "absent" and path_exists(text, pre["path"]):
                outcome["skipped"].append({"id": issue["id"], "reason": pre["skip_reason"]})
                return
        updated = apply_edits(text, issue["edits"])
        write_utf8_no_bom(abs_path, updated)
        outcome["applied"].append(issue["id"])
        outcome["modified_files"].add(issue["file"])
    except Exception as e:
        outcome["failed"].append({"id": issue["id"], "reason": str(e)})


def _finalize(scan: dict, requested, total: int, outcome: dict) -> dict:
    # I file modificati da QUESTO apply aggiornano il fingerprint: lo stale check deve
    # rilevare le modifiche esterne, non quelle fatte dalla skill stessa.
    if outcome["modified_files"]:
        for rel in outcome["modified_files"]:
            abs_path = os.path.join(scan["project_path"], *rel.split("/"))
            if os.path.exists(abs_path):
                scan["fingerprints"][rel] = fingerprint_file(abs_path)
        save_scan(scan)

    statuses: Dict[str, dict] = {}
    for i in outcome["applied"]:
        statuses[i] = {"stato": "Applicato"}
    for s in outcome["skipped"]:
        statuses[s["id"]] = {"stato": "Saltato", "motivo": s["reason"]}
    for f in outcome["failed"]:
        statuses[f["id"]] = {"stato": "Fallito", "motivo": f["reason"]}

    excel_update = None
    try:
        if os.path.exists(scan["excel_path"]):
            update_excel_status(scan["excel_path"], statuses)
        else:
            excel_update = f"file Excel non trovato ({scan['excel_path']}): stato non aggiornato"
    except Exception as e:
        excel_update = f"aggiornamento Excel fallito: {e}"

    res = {
        "status": "applied",
        "scan_id": scan["scan_id"],
        "requested": requested,
        "total": total,
        "applied": len(outcome["applied"]),
        "applied_ids": outcome["applied"],
        "skipped": outcome["skipped"],
        "failed": outcome["failed"],
        "excel_path": scan["excel_path"],
    }
    if excel_update:
        res["excel_update"] = excel_update
    return res


def cmd_apply(args) -> int:
    issue_ids = None
    if args.issue_ids is not None:
        issue_ids = [s for s in args.issue_ids.split(",") if s.strip()]
        if not issue_ids:
            print(
                json.dumps(
                    {"error": "no_issues_selected", "message": "Nessuna issue selezionata: nessuna azione eseguita."},
                    indent=2,
                    ensure_ascii=False,
                )
            )
            return 1

    scan = get_scan_by_id(args.scan_id, args.project_path) if args.scan_id else latest_scan(args.project_path)
    if not scan:
        print(
            json.dumps(
                {
                    "status": "no_scan_found",
                    "message": "Nessuno scan trovato per questo progetto: esegui prima il comando scan.",
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        return 2

    # Selezione esplicita di id = conferma implicita: si applicano solo quelle.
    if issue_ids is not None:
        by_id = {i["id"]: i for i in scan["issues"]}
        outcome = {"applied": [], "skipped": [], "failed": [], "modified_files": set()}
        for iid in issue_ids:
            issue = by_id.get(iid)
            if not issue:
                outcome["failed"].append({"id": iid, "reason": "issue non trovata, riesegui lo scan"})
                continue
            _apply_issue(scan, issue, outcome)
        print(json.dumps(_finalize(scan, issue_ids, len(issue_ids), outcome), indent=2, ensure_ascii=False))
        return 0

    changed = stale_files(scan)
    if changed:
        print(
            json.dumps(
                {
                    "status": "stale_scan",
                    "scan_id": scan["scan_id"],
                    "scan_timestamp": scan["timestamp"],
                    "changed_files": changed[:20],
                    "message": "I file del progetto sono stati modificati dopo lo scan: rilancia lo scan prima di applicare i fix.",
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        return 3

    if not args.confirm_all:
        print(
            json.dumps(
                {
                    "status": "needs_confirmation",
                    "scan_id": scan["scan_id"],
                    "scan_timestamp": scan["timestamp"],
                    "total_issues": len(scan["issues"]),
                    "auto_fixable": sum(1 for i in scan["issues"] if i["auto_fixable"]),
                    "message": "Nessuna modifica eseguita. Richiama con --confirm-all dopo la conferma dell'utente, oppure passa --issue-ids.",
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        return 4

    outcome = {"applied": [], "skipped": [], "failed": [], "modified_files": set()}
    for issue in scan["issues"]:
        _apply_issue(scan, issue, outcome)
    print(json.dumps(_finalize(scan, "all", len(scan["issues"]), outcome), indent=2, ensure_ascii=False))
    return 0


# ================================================================ CLI


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Accessibilita' PBIP/PBIR: alt text, tooltip, tab order")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("scan", help="analisi in sola lettura + Excel + cache")
    s.add_argument("project_path")
    s.add_argument("--output-dir", default=None, help="cartella dove salvare l'Excel (default: radice del progetto)")
    s.add_argument("--print", choices=["summary", "full", "json"], default="summary")
    s.set_defaults(func=cmd_scan)

    a = sub.add_parser("apply", help="applica i fix dell'ultimo scan")
    a.add_argument("project_path")
    a.add_argument("--confirm-all", action="store_true")
    a.add_argument("--issue-ids", default=None, help="lista di ID separati da virgola")
    a.add_argument("--scan-id", default=None)
    a.set_defaults(func=cmd_apply)

    i = sub.add_parser("issues", help="interroga le issue dell'ultimo scan")
    i.add_argument("project_path")
    i.add_argument("--scan-id", default=None)
    i.add_argument("--category", default=None, help="una o piu' categorie separate da virgola")
    i.add_argument("--page", default=None)
    i.add_argument("--only-fixable", action="store_true")
    i.add_argument("--limit", type=int, default=None)
    i.set_defaults(func=cmd_issues)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
