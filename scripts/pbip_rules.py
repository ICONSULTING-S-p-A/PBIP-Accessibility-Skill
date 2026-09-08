"""
Lettura di un progetto PBIP/PBIR e regole di accessibilita' (alt text, tooltip, tab order).

Porting fedele delle regole del prompt "accessibilita' PBIR": ogni funzione e' pura,
lavora sui dati letti dal disco e ritorna una lista di "issue" senza scrivere nulla.
La separazione lettura/decisione/scrittura e' quello che rende possibile mostrare
all'utente cosa cambierebbe prima di toccare il progetto.
"""
from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional

from jsonedit import strip_bom

# ---------------------------------------------------------------- tipi di visual

ALT_TEXT_SUPPORTED = {
    "actionButton",
    "card",
    "cardVisual",
    "barChart",
    "clusteredBarChart",
    "clusteredColumnChart",
    "columnChart",
    "donutChart",
    "slicer",
    "textbox",
    "lineChart",
    "lineStackedColumnComboChart",
    "lineClusteredColumnComboChart",
    "hundredPercentStackedColumnChart",
    "pageNavigator",
    "shapeMap",
    # Tipi tabellari, aggiunti il 2026-08-31. Usano lo stesso percorso e lo stesso formato
    # Literal degli altri visual core, e Power BI Desktop espone il campo alt text nel
    # pannello di formattazione anche per tabelle e matrici. Prima di questa aggiunta
    # finivano fra le anomalie ("tipo non presente negli elenchi dei tipi supportati"),
    # il che era un falso negativo su due dei tipi piu' diffusi nei report reali.
    "tableEx",
    "pivotTable",
    "matrix",
}

# Tipi decorativi: non ricevono mai un altText descrittivo, solo la stringa vuota che dice
# allo screen reader di ignorarli, piu' l'esclusione dall'ordine di tabulazione.
DECORATIVE_TYPES = {"image", "shape", "basicShape"}
ACTION_BUTTON = "actionButton"
GROUP_TYPE = "visualGroup"

ITALIAN_TYPE_NAMES = {
    "card": "KPI",
    "cardVisual": "KPI",
    "barChart": "Grafico a barre",
    "clusteredBarChart": "Grafico a barre",
    "clusteredColumnChart": "Grafico a colonne",
    "columnChart": "Grafico a colonne",
    "donutChart": "Grafico ad anello",
    "slicer": "Filtro",
    "textbox": "Casella di testo",
    "visualGroup": "Gruppo di visuals",
    "actionButton": "Pulsante",
    "image": "Immagine",
    "shape": "Forma",
    "basicShape": "Forma",
    "lineChart": "Grafico a linee",
    "lineStackedColumnComboChart": "Grafico combinato a linee e colonne impilate",
    "lineClusteredColumnComboChart": "Grafico combinato a linee e colonne raggruppate",
    "hundredPercentStackedColumnChart": "Grafico a colonne impilate al 100%",
    "pageNavigator": "Barra di navigazione tra pagine",
    "shapeMap": "Mappa a forme",
    "tableEx": "Tabella",
    "pivotTable": "Tabella pivot",
    "matrix": "Matrice",
}


def is_known_type(visual_type: str) -> bool:
    return (
        visual_type in ALT_TEXT_SUPPORTED
        or visual_type in DECORATIVE_TYPES
        or visual_type == ACTION_BUTTON
        or visual_type == GROUP_TYPE
    )


# ---------------------------------------------------------------- percorsi JSON

ALT_TEXT_PATH = ["visual", "visualContainerObjects", "general", 0, "properties", "altText"]
GROUP_ALT_TEXT_PATH = ["visualGroup", "objects", "general", 0, "properties", "altText"]
TOOLTIP_PATH = ["visual", "visualContainerObjects", "visualLink", 0, "properties", "tooltip"]
GENERAL_TOOLTIP_TEXT_PATH = ["visual", "visualContainerObjects", "general", 0, "properties", "tooltipText"]

ALT_TEXT_PATH_LABEL = "visualContainerObjects.general[0].properties.altText"
GROUP_ALT_TEXT_PATH_LABEL = "visualGroup.objects.general[0].properties.altText"
TOOLTIP_PATH_LABEL = "visualContainerObjects.visualLink[0].properties.tooltip"

MAX_TEXT_LENGTH = 150
STEP = 1000

# Valore sentinella con cui Power BI Desktop esclude davvero un visual dall'ordine di
# tabulazione (pannello Selezione -> icona occhio barrato). tabOrder 0 lascerebbe
# l'elemento raggiungibile da tastiera, semplicemente per ultimo.
HIDDEN_TAB_ORDER = -9999000


# ---------------------------------------------------------------- utilita' testo


def truncate(text: str) -> str:
    if len(text) <= MAX_TEXT_LENGTH:
        return text
    return text[: MAX_TEXT_LENGTH - 3] + "..."


def sanitize_for_literal(text: str) -> str:
    """Apostrofo tipografico al posto dell'apice: evita ambiguita' di escaping nei Literal PBIR."""
    return text.replace("'", "’")


def quote_literal(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"


def literal_expr(text: str) -> dict:
    return {"expr": {"Literal": {"Value": quote_literal(text)}}}


def unquote_literal(v: str) -> str:
    if len(v) >= 2 and v.startswith("'") and v.endswith("'"):
        return v[1:-1].replace("''", "'")
    return v


def literal_value(node: Any) -> Optional[str]:
    if not isinstance(node, dict):
        return None
    v = node.get("expr", {}).get("Literal", {}).get("Value") if isinstance(node.get("expr"), dict) else None
    if not isinstance(v, str):
        return None
    return unquote_literal(v)


def is_descriptive_title(title: Optional[str], visual_name: Optional[str] = None) -> bool:
    """
    Un titolo e' "descrittivo" se non e' un UUID, un nome autogenerato (visualContainer1,
    Group 2) o un nome interno con underscore (Testo_Donazioni): in quei casi non dice
    nulla di utile a chi usa uno screen reader.
    """
    if not title:
        return False
    t = title.strip()
    if len(t) < 3:
        return False
    if visual_name and t == visual_name:
        return False
    if re.fullmatch(r"visualContainer\d*", t, re.I):
        return False
    if re.fullmatch(r"(group|gruppo)\s*\d*", t, re.I):
        return False
    if re.fullmatch(r"[0-9a-f]{8,}", t.replace("-", ""), re.I):
        return False
    if re.fullmatch(r"[A-Za-zÀ-ÿ0-9]+(_[A-Za-zÀ-ÿ0-9]+)+", t):
        return False
    return True


MEASURE_ROLES = ["Values", "Y", "Y2", "Value", "Data", "Size"]
DIMENSION_ROLES = ["Category", "Axis", "Rows", "Series", "Columns", "Details", "Legend"]


def _first_of_roles(v: "VisualInfo", roles: List[str]) -> Optional[str]:
    for r in roles:
        names = v.query_roles.get(r)
        if names:
            return names[0]
    return None


def semantic_subject(v: "VisualInfo", ignore_title: bool = False) -> Optional[str]:
    """
    Soggetto semantico: 1) titolo descrittivo, 2) misura/dimensione dal blocco query.
    None = contesto insufficiente (in quel caso non si inventa un alt text, si segnala).

    ignore_title serve per gli slicer: il loro titolo e' spesso un residuo di
    copia-incolla non aggiornato dopo il rebind del campo, quindi mente.
    """
    if not ignore_title and is_descriptive_title(v.title, v.name):
        return (v.title or "").strip()
    m = _first_of_roles(v, MEASURE_ROLES)
    d = _first_of_roles(v, DIMENSION_ROLES)
    if m and d:
        return f"{m} per {d}"
    if m:
        return m
    if d:
        return d
    for names in v.query_roles.values():
        if names:
            return names[0]
    return None


# ---------------------------------------------------------------- modello dati


class VisualInfo:
    def __init__(self, **kw):
        self.__dict__.update(kw)


class PageInfo:
    def __init__(self, name, display_name, folder_path, visuals):
        self.name = name
        self.display_name = display_name
        self.folder_path = folder_path
        self.visuals: List[VisualInfo] = visuals


class ProjectInfo:
    def __init__(self, project_path, report_folder, report_name, pbir_version, schema_anomalies, pages, scanned_files):
        self.project_path = project_path
        self.report_folder = report_folder
        self.report_name = report_name
        self.pbir_version = pbir_version
        self.schema_anomalies: List[str] = schema_anomalies
        self.pages: List[PageInfo] = pages
        self.scanned_files: List[str] = scanned_files


def read_json_file(path: str) -> Any:
    import json

    with open(path, encoding="utf-8") as f:
        return json.loads(strip_bom(f.read()))


def _projection_name(p: Any) -> Optional[str]:
    if not isinstance(p, dict):
        return None
    nq = p.get("nativeQueryRef")
    if isinstance(nq, str) and nq.strip():
        return nq.strip()
    f = p.get("field") or {}
    prop = None
    for getter in (
        lambda: f.get("Measure", {}).get("Property"),
        lambda: f.get("Column", {}).get("Property"),
        lambda: f.get("Aggregation", {}).get("Expression", {}).get("Column", {}).get("Property"),
        lambda: f.get("HierarchyLevel", {}).get("Level"),
    ):
        try:
            prop = getter()
        except AttributeError:
            prop = None
        if isinstance(prop, str) and prop.strip():
            return prop.strip()
    qr = p.get("queryRef")
    if isinstance(qr, str) and qr.strip():
        return qr.split(".")[-1].strip()
    return None


def _group_general_props(js: Any) -> dict:
    g = (js.get("visualGroup") or {}).get("objects", {}).get("general")
    if isinstance(g, list):
        return (g[0] or {}).get("properties", {}) if g else {}
    if isinstance(g, dict):
        return g.get("properties", {}) or {}
    return {}


def parse_visual(page_name, page_display_name, file_path, rel_path, js) -> VisualInfo:
    is_group = js.get("visualGroup") is not None
    visual = js.get("visual") or {}
    visual_type = GROUP_TYPE if is_group else (visual.get("visualType") or "sconosciuto")
    pos = js.get("position") or {}
    position = {
        "x": float(pos.get("x", 0) or 0),
        "y": float(pos.get("y", 0) or 0),
        "width": float(pos.get("width", 0) or 0),
        "height": float(pos.get("height", 0) or 0),
        "z": float(pos["z"]) if pos.get("z") is not None else None,
        "tabOrder": float(pos["tabOrder"]) if pos.get("tabOrder") is not None else None,
    }

    vco = None if is_group else (visual.get("visualContainerObjects") or {})
    title = None
    if vco:
        tl = vco.get("title")
        if isinstance(tl, list) and tl:
            title = literal_value((tl[0] or {}).get("properties", {}).get("text"))

    if is_group:
        general_props = _group_general_props(js)
    else:
        gl = (vco or {}).get("general")
        general_props = (gl[0] or {}).get("properties", {}) if isinstance(gl, list) and gl else {}
    general_props = general_props or {}

    alt_raw = general_props.get("altText", "__MISSING__")
    has_alt_text = alt_raw != "__MISSING__"
    alt_text = literal_value(alt_raw) if has_alt_text else None

    tt_raw = general_props.get("tooltipText", "__MISSING__")
    general_tooltip_text_present = tt_raw != "__MISSING__"
    general_tooltip_text_value = literal_value(tt_raw) if general_tooltip_text_present else None

    vl_props = None
    if vco:
        vl = vco.get("visualLink")
        if isinstance(vl, list) and vl:
            vl_props = (vl[0] or {}).get("properties") or {}
    if vl_props is not None:
        visual_link = {
            "present": True,
            "type": literal_value(vl_props.get("type")),
            "navigationSection": literal_value(vl_props.get("navigationSection")),
            "hasTooltip": "tooltip" in vl_props,
            "tooltip": literal_value(vl_props.get("tooltip")),
        }
    else:
        visual_link = {"present": False, "type": None, "navigationSection": None, "hasTooltip": False, "tooltip": None}

    textbox_text = None
    if visual_type == "textbox":
        gl = (visual.get("objects") or {}).get("general")
        paragraphs = (gl[0] or {}).get("properties", {}).get("paragraphs") if isinstance(gl, list) and gl else None
        if isinstance(paragraphs, list):
            parts = []
            for p in paragraphs:
                runs = (p or {}).get("textRuns")
                if not isinstance(runs, list):
                    continue
                for r in runs:
                    val = (r or {}).get("value")
                    if isinstance(val, str) and val.strip():
                        parts.append(val.strip())
            joined = re.sub(r"\s+", " ", " ".join(parts)).strip()
            if joined:
                textbox_text = joined

    button_text = None
    text_objs = (visual.get("objects") or {}).get("text")
    if isinstance(text_objs, list):
        for t in text_objs:
            val = literal_value((t or {}).get("properties", {}).get("text"))
            if val and val.strip():
                button_text = val.strip()
                break

    has_non_transparent_fill = False
    fills = (visual.get("objects") or {}).get("fill")
    if isinstance(fills, list):
        for f in fills:
            props = (f or {}).get("properties") or {}
            if literal_value(props.get("show")) == "false":
                continue
            tr = literal_value(props.get("transparency"))
            try:
                num = float(re.sub(r"[^0-9.\-]", "", tr)) if tr else 0.0
            except ValueError:
                num = 0.0
            if not num >= 100:
                has_non_transparent_fill = True
                break

    query_roles: Dict[str, List[str]] = {}
    qs = (visual.get("query") or {}).get("queryState")
    if isinstance(qs, dict):
        for role, block in qs.items():
            projs = (block or {}).get("projections")
            if not isinstance(projs, list):
                continue
            names = [n for n in (_projection_name(p) for p in projs) if n]
            if names:
                query_roles[role] = names

    is_interactive_shape = visual_type in DECORATIVE_TYPES and visual_link["present"]

    return VisualInfo(
        page_name=page_name,
        page_display_name=page_display_name,
        name=js.get("name") if isinstance(js.get("name"), str) else os.path.basename(os.path.dirname(file_path)),
        file_path=file_path,
        rel_path=rel_path,
        visual_type=visual_type,
        is_group=is_group,
        group_display_name=(js.get("visualGroup") or {}).get("displayName") if is_group else None,
        parent_group_name=js.get("parentGroupName") if isinstance(js.get("parentGroupName"), str) else None,
        position=position,
        title=title,
        query_roles=query_roles,
        has_alt_text=has_alt_text,
        alt_text=alt_text,
        general_tooltip_text_present=general_tooltip_text_present,
        general_tooltip_text_value=general_tooltip_text_value,
        visual_link=visual_link,
        textbox_text=textbox_text,
        button_text=button_text,
        has_non_transparent_fill=has_non_transparent_fill,
        is_interactive_shape=is_interactive_shape,
    )


def read_project(project_path: str) -> ProjectInfo:
    """
    project_path puo' essere la cartella che contiene il .pbip e le cartelle
    *.Report / *.SemanticModel, oppure direttamente la cartella *.Report.
    """
    abs_path = os.path.abspath(project_path)
    if not os.path.isdir(abs_path):
        raise SystemExit(f"Percorso non trovato o non e' una cartella: {abs_path}")

    root = abs_path
    report_folder = None
    if re.search(r"\.Report$", abs_path, re.I):
        report_folder = abs_path
        root = os.path.dirname(abs_path)
    else:
        candidates = [
            os.path.join(abs_path, d)
            for d in sorted(os.listdir(abs_path))
            if re.search(r"\.Report$", d, re.I) and os.path.isdir(os.path.join(abs_path, d))
        ]
        for c in candidates:
            if os.path.exists(os.path.join(c, "definition.pbir")) or os.path.exists(os.path.join(c, "definition")):
                report_folder = c
                break
    if not report_folder:
        raise SystemExit(
            f"Il percorso non e' un progetto PBIP valido: nessuna cartella *.Report con definition.pbir trovata in {abs_path}"
        )

    report_name = re.sub(r"\.Report$", "", os.path.basename(report_folder), flags=re.I)
    scanned_files: List[str] = []
    schema_anomalies: List[str] = []

    def rel(p: str) -> str:
        return os.path.relpath(p, root).replace(os.sep, "/")

    pbir_version = None
    pbir_file = os.path.join(report_folder, "definition.pbir")
    if os.path.exists(pbir_file):
        scanned_files.append(rel(pbir_file))
        try:
            pbir = read_json_file(pbir_file)
            pbir_version = pbir.get("version") if isinstance(pbir.get("version"), str) else None
            if not pbir_version:
                schema_anomalies.append('definition.pbir non contiene un campo "version" leggibile.')
            else:
                try:
                    major = float(re.match(r"[0-9.]+", pbir_version).group(0))
                except Exception:
                    major = None
                if major is not None and major < 4.0:
                    schema_anomalies.append(
                        f"definition.pbir dichiara version {pbir_version}: versione precedente allo schema atteso (>= 4.0). "
                        "Nota: in Power BI Desktop 2.147.x la proprieta' tabOrder era ignorata silenziosamente. "
                        "Verificare la versione di Desktop in uso."
                    )
        except Exception as e:
            schema_anomalies.append(f"definition.pbir non e' un JSON valido: {e}")
    else:
        schema_anomalies.append("File definition.pbir non trovato nella cartella *.Report.")

    pages_dir = os.path.join(report_folder, "definition", "pages")
    pages: List[PageInfo] = []

    if not os.path.isdir(pages_dir):
        if os.path.exists(os.path.join(report_folder, "report.json")):
            schema_anomalies.append(
                "Il report e' in formato PBIR-legacy (report.json monolitico, senza definition/pages): "
                "formato precedente non supportato dallo scan per-visual. "
                "Salvare il progetto con il formato PBIR avanzato da Power BI Desktop."
            )
        else:
            schema_anomalies.append("Cartella definition/pages non trovata: struttura del report non riconosciuta.")
        return ProjectInfo(root, report_folder, report_name, pbir_version, schema_anomalies, pages, scanned_files)

    page_order = None
    pages_json_path = os.path.join(pages_dir, "pages.json")
    if os.path.exists(pages_json_path):
        scanned_files.append(rel(pages_json_path))
        try:
            pj = read_json_file(pages_json_path)
            if isinstance(pj.get("pageOrder"), list):
                page_order = pj["pageOrder"]
        except Exception as e:
            schema_anomalies.append(f"pages/pages.json non e' un JSON valido: {e}")

    page_folders = [
        os.path.join(pages_dir, d)
        for d in sorted(os.listdir(pages_dir))
        if os.path.isdir(os.path.join(pages_dir, d)) and os.path.exists(os.path.join(pages_dir, d, "page.json"))
    ]
    if page_order:
        order_index = {n: i for i, n in enumerate(page_order)}
        page_folders.sort(key=lambda p: order_index.get(os.path.basename(p), 10**9))

    for page_folder in page_folders:
        page_json_path = os.path.join(page_folder, "page.json")
        scanned_files.append(rel(page_json_path))
        try:
            page_json = read_json_file(page_json_path)
        except Exception as e:
            schema_anomalies.append(f"{rel(page_json_path)} non e' un JSON valido: {e}")
            continue
        page_name = page_json.get("name") if isinstance(page_json.get("name"), str) else os.path.basename(page_folder)
        display_name = page_json.get("displayName") if isinstance(page_json.get("displayName"), str) else page_name

        visuals: List[VisualInfo] = []
        visuals_dir = os.path.join(page_folder, "visuals")
        if os.path.isdir(visuals_dir):
            for v_dir in sorted(os.listdir(visuals_dir)):
                v_json_path = os.path.join(visuals_dir, v_dir, "visual.json")
                if not os.path.exists(v_json_path):
                    continue
                scanned_files.append(rel(v_json_path))
                try:
                    v_json = read_json_file(v_json_path)
                    visuals.append(parse_visual(page_name, display_name, v_json_path, rel(v_json_path), v_json))
                except Exception as e:
                    schema_anomalies.append(f"{rel(v_json_path)} non e' un JSON valido: {e}")

        pages.append(PageInfo(page_name, display_name, page_folder, visuals))

    return ProjectInfo(root, report_folder, report_name, pbir_version, schema_anomalies, pages, scanned_files)


# ================================================================ REGOLE
#
# Passo 2 - Tab order
#
# tabOrder funziona in ordine DECRESCENTE: il valore piu' alto riceve il focus per
# primo. L'ordine di lettura e' sinistra->destra, alto->basso; il titolo di pagina
# (textbox in alto) va per primo.
#
# position.z non va MAI toccato: governa la sovrapposizione visiva, non la
# tabulazione. Riassegnarlo "in sincrono" con tabOrder inverte lo stacking di
# elementi sovrapposti e fa sparire etichette sotto i pulsanti.


def _same_row(a: VisualInfo, b: VisualInfo) -> bool:
    top = max(a.position["y"], b.position["y"])
    bottom = min(a.position["y"] + a.position["height"], b.position["y"] + b.position["height"])
    overlap = bottom - top
    min_h = min(a.position["height"], b.position["height"])
    return min_h > 0 and overlap >= min_h * 0.5


def reading_order(items: List[VisualInfo]) -> List[VisualInfo]:
    ordered = sorted(items, key=lambda v: (v.position["y"], v.position["x"]))
    rows: List[List[VisualInfo]] = []
    for it in ordered:
        row = next((r for r in rows if _same_row(r[0], it)), None)
        if row is not None:
            row.append(it)
        else:
            rows.append([it])
    rows.sort(key=lambda r: r[0].position["y"])
    for r in rows:
        r.sort(key=lambda v: (v.position["x"], v.position["y"]))
    flat = [v for r in rows for v in r]

    if rows:
        title_box = next((v for v in rows[0] if v.visual_type == "textbox"), None)
        if title_box is not None and flat and flat[0] is not title_box:
            flat.remove(title_box)
            flat.insert(0, title_box)
    return flat


def _is_decorative(v: VisualInfo) -> bool:
    # una shape con visualLink e' un pulsante mascherato: non va esclusa dal focus
    return v.visual_type in DECORATIVE_TYPES and not v.is_interactive_shape


def _fmt(n) -> str:
    if n is None:
        return "—"
    return str(int(n)) if float(n).is_integer() else str(n)


def tab_order_issues(page: PageInfo) -> List[dict]:
    issues: List[dict] = []

    # I gruppi contano come unita' singola: i figli non ricevono un tabOrder proprio.
    top_level = [v for v in page.visuals if not v.parent_group_name]
    # I decorativi vanno pero' nascosti dalla tastiera ovunque siano, anche dentro un gruppo.
    decorative = [v for v in page.visuals if _is_decorative(v)]
    focusable = [v for v in top_level if not _is_decorative(v)]

    ordered = reading_order(focusable)
    n = len(ordered)

    for i, v in enumerate(ordered):
        value = (n - i) * STEP
        if v.position["tabOrder"] == value:
            continue
        issues.append(
            {
                "id": f"tabOrder-{v.page_name}-{v.name}",
                "page": v.page_display_name,
                "page_name": v.page_name,
                "visual_name": v.name,
                "visual_type": v.visual_type,
                "category": "tabOrder_da_assegnare",
                "current_value": f"tabOrder={_fmt(v.position['tabOrder'])}",
                "proposed_value": f"tabOrder={value}",
                "target_path": "position.tabOrder",
                "file": v.rel_path,
                "auto_fixable": True,
                "edits": [{"path": ["position", "tabOrder"], "value": value}],
            }
        )

    for v in decorative:
        edits = []
        if v.position["tabOrder"] != HIDDEN_TAB_ORDER:
            edits.append({"path": ["position", "tabOrder"], "value": HIDDEN_TAB_ORDER})
        proposed = f"tabOrder={HIDDEN_TAB_ORDER}"
        # altText sempre portato a stringa vuota, anche quando la proprieta' non c'e'.
        # Escludere un elemento dall'ordine di tabulazione lo rende irraggiungibile col
        # Tab, ma non lo nasconde alla modalita' di lettura dello screen reader, che
        # percorre la pagina elemento per elemento: senza altText="" un logo verrebbe
        # comunque annunciato, tipicamente leggendo il nome del file. La stringa vuota e'
        # il modo dichiarativo di dire "elemento decorativo, ignoralo".
        #
        # ATTENZIONE: questa riga scrive altText su image/shape/basicShape, tipi che una
        # versione precedente delle regole dava per incompatibili con altText in general.
        # La verifica in Power BI Desktop non e' ancora stata fatta - vedi il riquadro
        # "Verifica ancora aperta" in references/regole-accessibilita-pbir.md. Se Desktop
        # dovesse rifiutare il file, ripristinare: if v.has_alt_text and v.alt_text != "":
        if v.alt_text != "":
            edits.append({"path": ALT_TEXT_PATH, "value": literal_expr("")})
            proposed += ", altText=''"
        if not edits:
            continue
        current = f"tabOrder={_fmt(v.position['tabOrder'])}"
        if v.has_alt_text:
            current += f', altText="{v.alt_text or ""}"'
        issues.append(
            {
                "id": f"tabOrder-{v.page_name}-{v.name}",
                "page": v.page_display_name,
                "page_name": v.page_name,
                "visual_name": v.name,
                "visual_type": v.visual_type,
                "category": "tabOrder_da_assegnare",
                "current_value": current,
                "proposed_value": proposed,
                "target_path": "position.tabOrder",
                "file": v.rel_path,
                "auto_fixable": True,
                "edits": edits,
            }
        )

    return issues


# ---------------------------------------------------------------- Passo 3 - tooltip pulsanti


def button_tooltip_text(v: VisualInfo, page: PageInfo, project: ProjectInfo) -> Optional[str]:
    """Descrizione comando del pulsante. None = destinazione/azione non determinabile."""
    link = v.visual_link
    if not link["present"]:
        return None
    t = link["type"]
    if t == "PageNavigation":
        dest = link["navigationSection"]
        dest_label = next((p.display_name for p in project.pages if p.name == dest), None) if dest else None
        # Il nome della pagina di destinazione viene prima del testo del pulsante: la frase
        # e' "Vai alla pagina X", quindi X deve essere una pagina. Usare il testo visibile
        # produceva formulazioni senza senso come "Vai alla pagina Vai al dettaglio" per un
        # pulsante etichettato "Vai al dettaglio". Il testo del pulsante resta il fallback
        # per quando la destinazione non e' risolvibile in un displayName.
        label = dest_label or v.button_text
        if not label:
            return None
        if dest and dest == page.name:
            return f"{label} (pagina corrente)"
        return f"Vai alla pagina {label}"
    if t == "Back":
        return "Torna alla pagina precedente"
    if t == "Bookmark":
        return f"Attiva la vista {v.button_text}" if v.button_text else None
    if t in ("WebUrl", "WebHyperlink"):
        return f"Apri il collegamento {v.button_text}" if v.button_text else None
    return None


def _issue_base(v: VisualInfo) -> dict:
    return {
        "page": v.page_display_name,
        "page_name": v.page_name,
        "visual_name": v.name,
        "visual_type": v.visual_type,
        "file": v.rel_path,
    }


def tooltip_issues(page: PageInfo, project: ProjectInfo) -> List[dict]:
    issues: List[dict] = []

    for v in page.visuals:
        if v.visual_type != ACTION_BUTTON:
            continue

        computed = button_tooltip_text(v, page, project)

        # Caso 1: tooltipText in general e' una proprieta' inesistente in Power BI
        # (probabile refuso per tooltip): va rimossa e la descrizione spostata nel visualLink.
        # altText in general su un actionButton e' invece valido: non si tocca qui.
        if v.general_tooltip_text_present:
            if not v.visual_link["present"]:
                issues.append(
                    dict(
                        _issue_base(v),
                        id=f"anomalia-tooltip-{v.page_name}-{v.name}",
                        category="anomalia",
                        current_value="general: tooltipText presenti",
                        proposed_value=None,
                        target_path=TOOLTIP_PATH_LABEL,
                        auto_fixable=False,
                        reason=(
                            "actionButton con tooltipText in visualContainerObjects.general ma senza blocco "
                            "visualLink: impossibile spostare la descrizione nel tooltip. Richiede intervento manuale."
                        ),
                        edits=[],
                    )
                )
                continue

            edits = [{"path": GENERAL_TOOLTIP_TEXT_PATH, "_remove": True}]
            source_text = v.general_tooltip_text_value or computed
            proposed = "rimozione di tooltipText da general"
            preconditions = []
            if not v.visual_link["hasTooltip"] and source_text:
                text = truncate(sanitize_for_literal(source_text))
                edits.append({"path": TOOLTIP_PATH, "value": literal_expr(text)})
                proposed += f'; tooltip: "{text}"'
                preconditions = [
                    {"kind": "absent", "path": TOOLTIP_PATH, "skip_reason": "proprieta' gia' presente, non sovrascritta"}
                ]
            elif v.visual_link["hasTooltip"]:
                proposed += " (tooltip gia' presente nel visualLink, non sovrascritto)"

            issues.append(
                dict(
                    _issue_base(v),
                    id=f"spostamento-{v.page_name}-{v.name}",
                    category="proprieta_da_spostare",
                    current_value="general: tooltipText presente (proprieta' inesistente su actionButton)",
                    proposed_value=proposed,
                    target_path=TOOLTIP_PATH_LABEL,
                    auto_fixable=True,
                    edits=edits,
                    preconditions=preconditions,
                )
            )
            continue

        # Caso 2: tooltip gia' presente -> non toccare
        if v.visual_link["hasTooltip"]:
            continue

        # Caso 3: tooltip mancante
        if not computed:
            issues.append(
                dict(
                    _issue_base(v),
                    id=f"anomalia-tooltip-{v.page_name}-{v.name}",
                    category="anomalia",
                    current_value=None,
                    proposed_value=None,
                    target_path=TOOLTIP_PATH_LABEL,
                    auto_fixable=False,
                    reason=(
                        "Destinazione o azione del pulsante non determinabile con certezza: tooltip non aggiunto."
                        if v.visual_link["present"]
                        else "actionButton senza blocco visualLink: azione non determinabile, tooltip non aggiunto."
                    ),
                    edits=[],
                )
            )
            continue

        text = truncate(sanitize_for_literal(computed))
        issues.append(
            dict(
                _issue_base(v),
                id=f"tooltip-{v.page_name}-{v.name}",
                category="tooltip_mancante",
                current_value=None,
                proposed_value=text,
                target_path=TOOLTIP_PATH_LABEL,
                auto_fixable=True,
                edits=[{"path": TOOLTIP_PATH, "value": literal_expr(text)}],
                preconditions=[
                    {"kind": "absent", "path": TOOLTIP_PATH, "skip_reason": "proprieta' gia' presente, non sovrascritta"}
                ],
            )
        )

    return issues


# ---------------------------------------------------------------- Passo 3 - alt text


def _anomaly(v: VisualInfo, issue_id: str, reason: str) -> dict:
    return dict(
        _issue_base(v),
        id=issue_id,
        category="anomalia",
        current_value=None,
        proposed_value=None,
        target_path="",
        auto_fixable=False,
        reason=reason,
        edits=[],
    )


def _alt_text_issue(v: VisualInfo, text: str, path, path_label: str) -> dict:
    final_text = truncate(sanitize_for_literal(text))
    return dict(
        _issue_base(v),
        id=f"altText-{v.page_name}-{v.name}",
        category="altText_mancante",
        current_value=None,
        proposed_value=final_text,
        target_path=path_label,
        auto_fixable=True,
        edits=[{"path": path, "value": literal_expr(final_text)}],
        preconditions=[{"kind": "absent", "path": path, "skip_reason": "proprieta' gia' presente, non sovrascritta"}],
    )


def _chart_sentence(visual_type: str, subject: str) -> str:
    if visual_type == "card":
        return f"KPI che mostra {subject}."
    if visual_type in ("barChart", "clusteredBarChart"):
        return f"Grafico a barre che confronta {subject}."
    if visual_type == "clusteredColumnChart":
        return f"Grafico a colonne che confronta {subject}."
    if visual_type == "donutChart":
        return f"Grafico ad anello che mostra la distribuzione di {subject}."
    # I tipi tabellari non "mostrano un andamento": riepilogano valori incrociati. Il verbo
    # riepiloga descrive meglio a chi non vede cosa trovera' leggendo il visual con lo
    # screen reader, che lo percorre cella per cella.
    if visual_type == "matrix":
        return f"Matrice che riepiloga {subject}."
    if visual_type in ("tableEx", "pivotTable"):
        return f"Tabella che riepiloga {subject}."
    return f"{ITALIAN_TYPE_NAMES.get(visual_type, 'Visual')} che mostra {subject}."


def _child_summary(c: VisualInfo, page: PageInfo, project: ProjectInfo) -> Optional[str]:
    if c.visual_type in DECORATIVE_TYPES and not c.is_interactive_shape:
        return None
    tipo = ITALIAN_TYPE_NAMES.get(c.visual_type, c.visual_type).lower()
    if c.visual_type == "textbox" and c.textbox_text:
        return f'testo "{c.textbox_text}"'
    if c.visual_type == ACTION_BUTTON:
        if c.button_text:
            return f"pulsante {c.button_text}"
        t = button_tooltip_text(c, page, project)
        return f'pulsante "{t}"' if t else None
    if c.visual_type == "slicer":
        subject = semantic_subject(c, ignore_title=True)
        return f"filtro per {subject}" if subject else None
    subject = semantic_subject(c)
    return f"{tipo} di {subject}" if subject else None


def alt_text_issues(page: PageInfo, project: ProjectInfo) -> List[dict]:
    issues: List[dict] = []

    for v in page.visuals:
        # decorativi: gestiti dalle regole di tab order (altText '' solo se presente)
        if v.visual_type in DECORATIVE_TYPES:
            continue

        # Gruppi: altText in visualGroup.objects.general, descritto dai figli
        if v.visual_type == GROUP_TYPE:
            if v.has_alt_text:
                continue
            children = [c for c in page.visuals if c.parent_group_name == v.name]
            parts = [s for s in (_child_summary(c, page, project) for c in children) if s]
            if not parts:
                issues.append(
                    _anomaly(
                        v,
                        f"anomalia-altText-{v.page_name}-{v.name}",
                        "Gruppo senza contesto semantico sufficiente: i visuals figli non forniscono "
                        "informazioni per generare un altText significativo.",
                    )
                )
                continue
            if v.group_display_name and is_descriptive_title(v.group_display_name, v.name):
                prefix = f"Sezione {v.group_display_name}: "
            else:
                prefix = "Sezione che contiene: "
            text = (prefix + "; ".join(parts) + ".").replace("..", ".")
            issues.append(_alt_text_issue(v, text, GROUP_ALT_TEXT_PATH, GROUP_ALT_TEXT_PATH_LABEL))
            continue

        # Un altText gia' presente chiude il discorso, qualunque sia il tipo del visual.
        # Va controllato PRIMA del tipo: un visual custom con un altText scritto a mano e'
        # gia' a norma, e segnalarlo come anomalia "nessun altText aggiunto" mandava chi
        # legge il report a cercare un problema che non esiste.
        if v.has_alt_text:
            continue

        # Tipo non riconosciuto (es. visual custom da marketplace): le capabilities non
        # sono garantite, aggiungere altText potrebbe rompere il report -> si segnala.
        if not is_known_type(v.visual_type):
            issues.append(
                _anomaly(
                    v,
                    f"anomalia-tipo-{v.page_name}-{v.name}",
                    f'Tipo di visual "{v.visual_type}" non presente negli elenchi dei tipi supportati: '
                    "nessun altText aggiunto.",
                )
            )
            continue

        if v.visual_type not in ALT_TEXT_SUPPORTED:
            continue

        if v.visual_type == "textbox":
            if v.textbox_text:
                issues.append(_alt_text_issue(v, v.textbox_text, ALT_TEXT_PATH, ALT_TEXT_PATH_LABEL))
            else:
                issues.append(
                    _anomaly(
                        v,
                        f"anomalia-altText-{v.page_name}-{v.name}",
                        "Textbox senza testo leggibile nei textRuns: impossibile generare l’altText obbligatorio.",
                    )
                )
            continue

        if v.visual_type == "pageNavigator":
            issues.append(
                _alt_text_issue(v, "Barra di navigazione tra le pagine del report.", ALT_TEXT_PATH, ALT_TEXT_PATH_LABEL)
            )
            continue

        if v.visual_type == ACTION_BUTTON:
            if v.button_text:
                issues.append(_alt_text_issue(v, f"Pulsante {v.button_text}.", ALT_TEXT_PATH, ALT_TEXT_PATH_LABEL))
            else:
                destination = button_tooltip_text(v, page, project)
                if destination:
                    issues.append(
                        _alt_text_issue(v, f"Pulsante: {destination}.", ALT_TEXT_PATH, ALT_TEXT_PATH_LABEL)
                    )
                else:
                    issues.append(
                        _anomaly(
                            v,
                            f"anomalia-altText-{v.page_name}-{v.name}",
                            "actionButton senza testo visibile ne' destinazione determinabile: "
                            "altText non generabile automaticamente.",
                        )
                    )
            continue

        if v.visual_type == "slicer":
            subject = semantic_subject(v, ignore_title=True)
            if not subject:
                issues.append(
                    _anomaly(
                        v,
                        f"anomalia-altText-{v.page_name}-{v.name}",
                        "Filtro senza campo dati determinabile: altText non generabile automaticamente.",
                    )
                )
                continue
            issues.append(
                _alt_text_issue(v, f"Filtro per selezionare {subject}.", ALT_TEXT_PATH, ALT_TEXT_PATH_LABEL)
            )
            continue

        subject = semantic_subject(v)
        if not subject:
            issues.append(
                _anomaly(
                    v,
                    f"anomalia-altText-{v.page_name}-{v.name}",
                    "Contesto semantico insufficiente (nessun titolo descrittivo ne' campi dati): "
                    "altText non determinabile automaticamente.",
                )
            )
            continue
        issues.append(
            _alt_text_issue(v, _chart_sentence(v.visual_type, subject), ALT_TEXT_PATH, ALT_TEXT_PATH_LABEL)
        )

    return issues
