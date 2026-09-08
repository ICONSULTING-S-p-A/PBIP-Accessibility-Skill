# Power BI Accessibility (PBIP/PBIR)

Audits and remediates the accessibility metadata that screen readers and keyboard
navigation actually consume in Power BI reports: **alternative text**, **navigation
tooltips** on buttons, and **tab order**.

Works on reports saved in **PBIP / PBIR** format (Power BI Desktop, *Save as project*),
where these properties are declarative JSON inside each `visual.json` — and therefore
machine-auditable, unlike the binary `.pbix` format.

Maintained by [Iconsulting S.p.A.](https://www.iconsulting.biz)

---

## Why this exists

Accessibility tooling covers websites, mobile apps and PDFs. It does not cover
business-intelligence reports — which are the main channel through which public
administrations and regulated operators publish data to citizens.

Under Directive (EU) 2019/882 (European Accessibility Act), transposed in Italy by
Legislative Decree 82/2022 alongside Law 4/2004, in-scope digital services must meet
EN 301 549 (WCAG level AA). A dashboard embedded in a compliant portal is not itself
compliant: a missing `altText` is an unlabelled element for a screen reader, and an
unset `tabOrder` means keyboard focus follows z-order rather than reading order.

On a real report — hundreds of `visual.json` files — this is neither feasible nor safe to
do by hand. The rules are subtle (which visual types support `altText`, where a button
tooltip lives, when *nothing* should be written), and applying them by eye across
hundreds of files produces silent errors that break the report in Power BI Desktop.

## What it does

| Check | Property written | Applies to |
|---|---|---|
| Alternative text | `visual.visualContainerObjects.general[0].properties.altText` | core visual types, decorative elements excluded |
| Group alternative text | `visualGroup.objects.general[0].properties.altText` | `visualGroup` only |
| Navigation tooltip | `visual.visualContainerObjects.visualLink[0].properties.tooltip` | `actionButton` only |
| Tab order | `position.tabOrder` | all visuals, decorative ones excluded from focus |

Three commands, in a read-then-confirm-then-write sequence:

```bash
# 1. read-only analysis: Excel report + on-screen summary, writes nothing to the project
python3 scripts/pbip_a11y.py scan "<project path>"

# 2. inspect specific findings
python3 scripts/pbip_a11y.py issues "<project path>" --category altText_mancante --limit 20

# 3. apply (refuses to run without an explicit confirmation flag)
python3 scripts/pbip_a11y.py apply "<project path>" --confirm-all
python3 scripts/pbip_a11y.py apply "<project path>" --issue-ids id1,id2
```

## Safety properties

These are enforced by the script, not by convention:

- **Never overwrites** an existing `altText` or `tooltip`, whatever its content —
  re-checked at apply time, not only at scan time.
- **`apply` without `--confirm-all` or `--issue-ids` writes nothing** and returns
  `needs_confirmation`.
- **Stale-scan guard**: if project files changed after the scan, `apply` returns
  `stale_scan` and writes nothing.
- **Minimal textual edits**: indentation, key order, line endings and every uninvolved
  byte are preserved, so the git diff shows only the lines that matter.
- **UTF-8 without BOM**, verified by re-reading the first bytes after each write (a BOM
  makes Power BI Desktop fail to parse the file).
- **`position.z` is never modified.**
- Findings that are not safely automatable are always reported as anomalies and never
  guessed at — see below.
- Nothing is written outside the project except the Excel report and a scan cache in the
  user's home directory. No files are created inside the project folder.

## Anomalies (deliberately not automated)

| Anomaly | Why it is not automatable |
|---|---|
| `actionButton` with no `visualLink` block | The action cannot be determined (typically a bookmark button) |
| Insufficient semantic context | No descriptive title and no data fields: an invented alt text would misinform the very users it targets |
| Unrecognised visual type (custom / marketplace) | Custom visual capabilities are not guaranteed; adding `altText` can break the report |
| Textbox with no text runs | No visible text to reuse |
| Incompatible `schemaVersion` | `tabOrder` was silently ignored in Power BI Desktop 2.147.x |

The skill surfaces these, explains them, and asks the user for the missing information.

## Requirements

- Python 3 with `openpyxl` (`pip install openpyxl`)
- A report in PBIP/PBIR format with a `definition/pages` folder
  (`definition.pbir` version >= 4.0)

## Try it on the bundled example

`examples/demo-pbip/` is a synthetic two-page report containing deliberate violations
plus cases that must *not* be touched. No real data, no semantic model required.

```bash
cp -r examples/demo-pbip /tmp/demo
python3 skills/pbip-accessibility/scripts/pbip_a11y.py scan /tmp/demo
```

Expected: 27 findings — 13 tab order, 9 missing alt text, 1 missing tooltip,
1 misplaced property, 3 anomalies. Then:

```bash
python3 skills/pbip-accessibility/scripts/pbip_a11y.py apply /tmp/demo --confirm-all
```

Expected: 24 applied, 3 skipped (the anomalies). Verify that `slicerAnno` still has its
hand-written alt text, that `imgLogoArpa` received `tabOrder: -9999000` (excluded from
keyboard focus), and that no `position.z` value changed.

## Privacy

See [PRIVACY.md](PRIVACY.md). Everything runs locally; the skill has no network calls
and no telemetry.

## Support

- Issues: `https://github.com/TODO-ORG/powerbi-accessibility/issues`
- Contact: `SkillSupport@iconsulting.com`

## License

Apache-2.0 — see [LICENSE](LICENSE).
