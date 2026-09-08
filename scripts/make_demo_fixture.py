#!/usr/bin/env python3
"""Genera un progetto PBIP sintetico per provare la skill senza dati reali.

Il fixture serve a due cose: dare a chi installa la skill un progetto su cui vedere subito
cosa fa, e dare agli eval in `evals/evals.json` un input riproducibile. Contiene di
proposito sia violazioni da correggere sia casi che *non* devono essere toccati, perche' la
parte difficile di questa skill non e' scrivere gli alt text: e' sapere quando non scriverli.

Uso:
    python3 scripts/make_demo_fixture.py /percorso/destinazione
    python3 scripts/pbip_a11y.py scan /percorso/destinazione/demo-pbip

Nessun modello semantico, nessun dato: i visual dichiarano i campi nel blocco query ma non
esiste un .SemanticModel, quindi il progetto non si apre in Power BI Desktop. Va bene: la
skill legge struttura e metadati, non risultati di query.
"""
from __future__ import annotations

import json
import os
import sys

# Ogni voce: (nome, visualType, x, y, z, extra) dove extra popola il visual.
# I nomi sono quelli citati in SUBMISSION.md, cosi' la documentazione resta verificabile.
PAGES = {
    "pagPanoramica": {
        "displayName": "Panoramica controlli",
        "visuals": [
            # titolo di pagina: l'alt text si ricava dal testo visibile, non dal title interno
            ("txtTitolo", "textbox", 20, 16, 1, {
                "objects": {"general": [{"properties": {"paragraphs": [
                    {"textRuns": [{"value": "Controlli impianti termici"}]}]}}]},
                "title": "'Testo_TitoloPagina'",
            }),
            # logo decorativo: nessun altText nel file, va creato vuoto
            ("imgLogoArpa", "image", 1100, 12, 7000, {}),
            # divisore decorativo
            ("shpDivisore", "shape", 20, 70, 5, {}),
            # card con titolo descrittivo: alt text derivabile
            ("cardControlliTotali", "card", 20, 90, 3, {
                "title": "'Controlli totali'",
                "measures": ["Numero controlli"],
            }),
            # barChart con titolo e campi: alt text derivabile
            ("barControlliPerProvincia", "barChart", 300, 90, 3, {
                "title": "'Controlli per provincia'",
                "measures": ["Numero controlli"], "dims": ["Provincia"],
            }),
            # anomalia: nessun titolo descrittivo, nessun campo -> alt text non inventabile
            ("cardSenzaContesto", "card", 700, 90, 3, {}),
            # slicer con altText scritto a mano: non va sovrascritto
            ("slicerAnno", "slicer", 1000, 90, 3, {
                "title": "'Anno'", "dims": ["Anno"],
                "altText": "'Filtro per anno di riferimento'",
            }),
            # actionButton con visualLink e senza tooltip: tooltip generabile
            ("btnVaiDettaglio", "actionButton", 20, 420, 4, {
                "buttonText": "'Vai al dettaglio'",
                "navigationSection": "'pagDettaglio'",
            }),
            # anomalia: actionButton senza visualLink (tipicamente un pulsante segnalibro)
            ("btnReimposta", "actionButton", 200, 420, 4, {
                "buttonText": "'Reimposta filtri'",
            }),
            # tooltipText nel posto sbagliato: va spostato in visualLink.tooltip
            ("btnHome", "actionButton", 380, 420, 4, {
                "buttonText": "'Home'",
                "navigationSection": "'pagPanoramica'",
                "tooltipText": "'Torna alla home'",
            }),
        ],
    },
    "pagDettaglio": {
        "displayName": "Dettaglio per impianto",
        "visuals": [
            ("txtTitoloDettaglio", "textbox", 20, 16, 1, {
                "objects": {"general": [{"properties": {"paragraphs": [
                    {"textRuns": [{"value": "Dettaglio per impianto"}]}]}}]},
            }),
            # matrix: prima del 2026-08-31 finiva fra le anomalie
            ("mtxEsitiPerProvincia", "matrix", 20, 80, 3, {
                "title": "'Esiti per provincia'",
                "measures": ["Numero controlli"], "dims": ["Provincia", "Esito"],
            }),
            # tableEx con altText scritto a mano: non va sovrascritto
            ("tblImpianti", "tableEx", 460, 80, 4, {
                "title": "'Elenco impianti'", "dims": ["Impianto"],
                "altText": "'Elenco degli impianti con esito del controllo'",
            }),
            # pivotTable senza altText: alt text derivabile
            ("pvtConformita", "pivotTable", 900, 80, 4, {
                "title": "'Conformita per tipologia'",
                "measures": ["Percentuale conformi"], "dims": ["Tipologia"],
            }),
            # anomalia: visual custom da marketplace, capabilities non garantite
            ("vizCustomMappa", "Mappa_ER_1A2B3C4D5E6F7890ABCDEF12345678", 20, 420, 3, {
                "title": "'Mappa province'",
            }),
            # gruppo: l'altText va sul contenitore, i figli non ricevono tabOrder proprio
            ("grpEsiti", "visualGroup", 460, 420, 2, {"groupDisplayName": "Esiti"}),
            ("donutEsiti", "donutChart", 470, 430, 3, {
                "title": "'Distribuzione esiti'",
                "measures": ["Numero controlli"], "dims": ["Esito"],
                "parentGroupName": "grpEsiti",
            }),
            ("lineTrend", "lineChart", 700, 430, 3, {
                "title": "'Andamento controlli'",
                "measures": ["Numero controlli"], "dims": ["Mese"],
                "parentGroupName": "grpEsiti",
            }),
        ],
    },
}


def literal(value: str) -> dict:
    return {"expr": {"Literal": {"Value": value}}}


def projections(names: list[str], kind: str) -> list[dict]:
    out = []
    for n in names:
        field = {"Measure": {"Property": n}} if kind == "m" else {"Column": {"Property": n}}
        out.append({"field": field, "queryRef": f"{kind}.{n}"})
    return out


def build_visual(name: str, vtype: str, x: int, y: int, z: int, extra: dict) -> dict:
    doc: dict = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": z, "width": 260, "height": 200},
    }
    if extra.get("parentGroupName"):
        doc["parentGroupName"] = extra["parentGroupName"]

    if vtype == "visualGroup":
        doc["visualGroup"] = {"displayName": extra.get("groupDisplayName", name)}
        return doc

    visual: dict = {"visualType": vtype}
    vco: dict = {}

    if extra.get("title"):
        vco["title"] = [{"properties": {"text": literal(extra["title"])}}]

    general_props: dict = {}
    if extra.get("altText"):
        general_props["altText"] = literal(extra["altText"])
    if extra.get("tooltipText"):
        # proprieta' inesistente in Power BI: la skill la sposta in visualLink.tooltip
        general_props["tooltipText"] = literal(extra["tooltipText"])
    if general_props:
        vco["general"] = [{"properties": general_props}]

    if extra.get("navigationSection"):
        vco["visualLink"] = [{"properties": {
            "show": literal("true"),
            "type": literal("'PageNavigation'"),
            "navigationSection": literal(extra["navigationSection"]),
        }}]

    if vco:
        visual["visualContainerObjects"] = vco

    if extra.get("objects"):
        visual["objects"] = extra["objects"]
    elif extra.get("buttonText"):
        visual["objects"] = {"text": [{"properties": {"text": literal(extra["buttonText"])}}]}

    qs: dict = {}
    if extra.get("measures"):
        qs["Values"] = {"projections": projections(extra["measures"], "m")}
    if extra.get("dims"):
        qs["Category"] = {"projections": projections(extra["dims"], "d")}
    if qs:
        visual["query"] = {"queryState": qs}

    visual["drillFilterOtherVisuals"] = True
    doc["visual"] = visual
    return doc


def write_json(path: str, data: dict) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    # UTF-8 senza BOM, con newline finale: le stesse convenzioni che la skill preserva
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
        fh.write("\n")


def main() -> int:
    dest = sys.argv[1] if len(sys.argv) > 1 else "."
    root = os.path.join(os.path.abspath(dest), "demo-pbip")
    report = os.path.join(root, "Demo Controlli.Report")

    write_json(os.path.join(root, "Demo Controlli.pbip"), {
        "version": "1.0",
        "artifacts": [{"report": {"path": "Demo Controlli.Report"}}],
    })
    write_json(os.path.join(report, "definition.pbir"), {
        "version": "4.0",
        "datasetReference": {"byPath": {"path": "../Demo Controlli.SemanticModel"}},
    })
    write_json(os.path.join(report, "definition", "pages", "pages.json"), {
        "pageOrder": list(PAGES.keys()),
        "activePageName": next(iter(PAGES)),
    })

    n = 0
    for page_name, page in PAGES.items():
        base = os.path.join(report, "definition", "pages", page_name)
        write_json(os.path.join(base, "page.json"), {
            "name": page_name,
            "displayName": page["displayName"],
            "width": 1280,
            "height": 720,
            "displayOption": "FitToPage",
        })
        for name, vtype, x, y, z, extra in page["visuals"]:
            write_json(os.path.join(base, "visuals", name, "visual.json"),
                       build_visual(name, vtype, x, y, z, extra))
            n += 1

    print(f"Fixture creato: {root}")
    print(f"{len(PAGES)} pagine, {n} visual")
    print()
    print("Provalo con:")
    print(f'  python3 scripts/pbip_a11y.py scan "{root}"')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
