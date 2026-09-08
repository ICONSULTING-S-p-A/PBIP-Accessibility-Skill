# Privacy notice

**Plugin:** Power BI Accessibility (PBIP/PBIR)
**Publisher:** Iconsulting S.p.A. — TODO registered address
**Contact:** TODO@iconsulting.biz
**Last updated:** TODO

## Summary

This plugin performs no data collection. It contains no network calls, no telemetry, no
analytics and no remote service. Iconsulting receives no data of any kind from users of
this plugin.

## What the plugin reads

Only files inside the Power BI project folder the user explicitly points it at:

- `*.pbip`
- `<report>.Report/definition.pbir`
- `<report>.Report/definition/pages/**/page.json`
- `<report>.Report/definition/pages/**/visuals/**/visual.json`

It reads report *structure and metadata* (visual types, titles, positions, existing
accessibility properties). It does not read the semantic model, query results, or any
report data.

The plugin does not access the user's conversation history, Claude's memory,
conversation summaries, or any file outside the project path provided.

## What the plugin writes

1. **In place, inside the project**: the four accessibility properties documented in the
   README, in the `visual.json` files concerned, and only after explicit confirmation.
2. **An Excel report** (`accessibilita_<Report>_<timestamp>.xlsx`) next to the project,
   or in `--output-dir` if specified.
3. **A scan cache** under the user's home directory (`~/.pbip-accessibility-mcp`),
   deliberately outside the project so the repository is not polluted. It contains the
   findings of the last scan and can be deleted at any time.

Nothing is transmitted anywhere. All processing happens on the machine running the
skill, within the user's session.

## Data retention

Iconsulting stores nothing. The only persistent artefacts are the local files listed
above, under the user's own control.

## Third-party dependencies

Python standard library plus `openpyxl` (Excel generation, local only). No other
runtime dependency.

## Changes

Any change to this notice will be published in this file, in the repository, with an
updated date.
