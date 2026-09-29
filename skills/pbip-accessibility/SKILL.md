---
name: pbip-accessibility
description: Audits and remediates the accessibility metadata that screen readers and keyboard navigation consume in Power BI reports saved in PBIP/PBIR format - alternative text, navigation tooltips on buttons, and tab order - against WCAG / EN 301 549. Produces an Excel audit report and edits visual.json files in place as UTF-8 without BOM. Use this skill whenever the user points at, uploads or mentions a Power BI project folder (a .pbip file, a *.Report folder, definition/pages, visual.json) together with accessibility, alt text, testo alternativo, descrizione comando, tooltip, tab order, ordine di tabulazione, screen reader, WCAG, EN 301 549, or a compliance audit - and also when they only ask to "fix", "check", "sistemare", "controllare" or "rendere accessibile" a Power BI report without naming the properties. Does NOT apply to binary .pbix files, which must first be re-saved from Power BI Desktop as a project.
---

# Accessibilità PBIP/PBIR — alt text, tooltip, tab order

Questa skill porta un report Power BI in formato PBIP/PBIR a norma sui tre aspetti che
uno screen reader e la navigazione da tastiera usano davvero: il testo alternativo dei
visual, la descrizione comando dei pulsanti di navigazione e l'ordine di tabulazione.

Il lavoro è fatto da uno script deterministico in `scripts/pbip_a11y.py`, non a mano
visual per visual. Il motivo è concreto: un report reale ha centinaia di `visual.json`,
e le regole (quali tipi supportano `altText`, dove va il tooltip, quando *non* si deve
scrivere niente) sono abbastanza sottili che applicarle a occhio su 400 file produce
errori silenziosi che poi rompono il report in Power BI Desktop.

Il tuo ruolo è: capire quale progetto trattare, far girare l'analisi, **mostrare a chi
usa la skill cosa cambierebbe e farsi confermare**, applicare, e infine spiegare cosa è
cambiato e cosa resta da valutare a mano.

## Requisiti

Serve Python 3 con `openpyxl` (per l'Excel di riepilogo). Se manca:

```bash
pip install openpyxl --break-system-packages
```

## Passo 1 — Individua il progetto

Il percorso da passare allo script è la cartella che contiene il file `.pbip` e le
cartelle `*.Report` / `*.SemanticModel`, oppure direttamente la cartella `*.Report`.

Se l'utente indica una cartella generica, cercala:

```bash
find "<cartella indicata>" -maxdepth 3 \( -name "*.pbip" -o -name "definition.pbir" \)
```

**Lavora sempre in-place, nella posizione originale del progetto.** Non copiare il
progetto in una cartella di lavoro, non creare cartelle nuove, non spostare file: chi
usa questa skill ha quel progetto aperto in Power BI Desktop e sotto controllo di
versione, e una copia altrove significa modifiche che poi vanno riportate a mano.

Se il progetto è sotto git, vale la pena verificare che l'albero di lavoro sia pulito
prima di modificare (`git status`), così le modifiche della skill restano isolate in un
diff leggibile. La skill non fa commit: quello resta una decisione dell'utente.

## Passo 2 — Scansione (sola lettura)

```bash
python3 <cartella-skill>/scripts/pbip_a11y.py scan "<percorso progetto>"
```

Lo scan non modifica **nulla**. Produce:

- un **Excel** di riepilogo accanto al progetto (`accessibilita_<Report>_<timestamp>.xlsx`)
  con tre fogli: *Riepilogo*, *Dettaglio issue*, *Anomalie*;
- una **cache di scan** fuori dal progetto (per non sporcare il repo), che `apply` usa
  per sapere cosa applicare;
- un riepilogo a schermo con i conteggi per categoria e per pagina e l'elenco completo
  delle anomalie.

Opzioni utili:

- `--output-dir DIR` — salva l'Excel altrove (utile se il progetto è read-only).
- `--print full` — stampa anche il dettaglio di ogni singola modifica proposta. Su
  report grandi sono centinaia di righe: usalo se l'utente vuole ispezionare tutto,
  altrimenti resta sul riepilogo e interroga le issue in modo mirato con `issues`.

Per guardare solo una parte:

```bash
python3 <cartella-skill>/scripts/pbip_a11y.py issues "<percorso>" \
  --category altText_mancante --page "Homepage" --limit 20
```

Categorie: `altText_mancante`, `tooltip_mancante`, `tabOrder_da_assegnare`,
`proprieta_da_spostare`, `anomalia`.

## Passo 3 — Mostra e fai confermare

Prima di scrivere sul progetto, riporta all'utente:

- quante issue sono state trovate, divise per categoria e per pagina;
- qualche esempio concreto di alt text e tooltip generati (non solo i numeri: è il
  contenuto dei testi che l'utente deve poter contestare);
- le anomalie, cioè i casi che lo script **non** correggerà da solo;
- dove è stato salvato l'Excel.

Poi chiedi conferma. Se `apply` viene lanciato senza `--confirm-all` e senza
`--issue-ids`, non modifica niente e risponde `needs_confirmation`: è una rete di
sicurezza, non un passaggio da aggirare.

Se l'utente vuole applicare solo una parte (per esempio solo il tab order, o solo una
pagina), raccogli gli ID con `issues` e passali a `--issue-ids`.

## Passo 4 — Applica

```bash
# tutto ciò che è auto-fixable
python3 <cartella-skill>/scripts/pbip_a11y.py apply "<percorso progetto>" --confirm-all

# solo issue selezionate (la selezione esplicita vale come conferma)
python3 <cartella-skill>/scripts/pbip_a11y.py apply "<percorso>" --issue-ids id1,id2,id3
```

Cosa garantisce la scrittura:

- **Mai sovrascrivere** un `altText` o un `tooltip` già presente. La verifica è rifatta
  al momento dell'apply, non solo allo scan: nel frattempo l'utente potrebbe aver
  scritto qualcosa a mano.
- **Edit testuali minimi**: indentazione, ordine delle chiavi, fine riga e tutto il JSON
  non coinvolto restano identici. Il diff mostra solo le righe che contano.
- **UTF-8 senza BOM**, verificato leggendo i primi byte dopo il salvataggio. Il BOM fa
  fallire il parsing in Power BI Desktop.
- **`position.z` non viene mai toccato.**
- Le issue non auto-fixable finiscono sempre in `skipped`, mai applicate a caso.
- Le colonne *Stato* e *Motivo* dell'Excel dello scan vengono aggiornate
  (`Applicato` / `Saltato` / `Fallito`), così quel file resta il registro nel tempo.

Se i file del progetto sono cambiati dopo lo scan, `apply` risponde `stale_scan` e non
modifica niente: rifai lo scan. È il caso tipico di chi salva da Power BI Desktop tra
un comando e l'altro.

Una verifica che vale la pena suggerire all'utente dopo l'apply: riaprire il report in
Power BI Desktop e controllare che si carichi senza errori. Serve in generale, ma in
particolare sui progetti con molte immagini e forme, perché la regola che scrive
`altText: ""` sugli elementi decorativi anche quando la proprietà è assente non è ancora
stata confermata su Desktop — il riquadro "Verifica ancora aperta" in
`references/regole-accessibilita-pbir.md` spiega cosa fare se dovesse dare problemi.

## Passo 5 — Riepilogo finale

Chiudi con un riepilogo in italiano che contenga:

- elenco dei file modificati (o il conteggio, se sono centinaia, con l'Excel come
  riferimento per il dettaglio);
- i valori `tabOrder` assegnati, nell'ordine di lettura, almeno per le pagine principali;
- gli alt text generati (il testo prodotto, non solo il numero);
- l'`altText` generato per ogni gruppo;
- i pulsanti aggiornati con il `tooltip` di navigazione;
- immagini e forme decorative escluse dalla tastiera;
- gli elementi **non** modificati perché non interattivi o perché avevano già la
  proprietà;
- le anomalie, con il motivo, perché sono l'unica parte che richiede una decisione umana.

Presenta anche il file Excel all'utente perché possa aprirlo.

## Anomalie: cosa sono e come trattarle

Lo script non inventa mai un testo quando il contesto non basta. I casi ricorrenti:

| Anomalia | Perché non è automatizzabile | Cosa proporre |
|---|---|---|
| `actionButton` senza `visualLink` | Non c'è modo di sapere che azione fa (spesso è un pulsante legato a un segnalibro) | Chiedere all'utente la destinazione, poi scrivere `visualLink[0].properties.tooltip` a mano |
| Contesto semantico insufficiente | Nessun titolo descrittivo né campi dati: un alt text inventato sarebbe fuorviante per chi non vede il visual | Chiedere all'utente cosa mostra il visual |
| Tipo di visual non presente nell'elenco dei supportati | Copre due casi diversi: i visual custom/marketplace, le cui capabilities non sono garantite, e i tipi core non ancora inseriti in `ALT_TEXT_SUPPORTED` (`pieChart`, `gauge`, `treemap`, `map`, `multiRowCard`…). Aggiungere `altText` a un tipo che non lo supporta rompe il report | Per un custom: verificare le capabilities a mano. Per un tipo core: verificarlo in Power BI Desktop e aggiungerlo a `ALT_TEXT_SUPPORTED` e `ITALIAN_TYPE_NAMES` in `scripts/pbip_rules.py`, così la segnalazione non si ripresenta al prossimo scan |
| Textbox senza testo nei `textRuns` | Non c'è testo visibile da riusare | Chiedere all'utente |
| `schemaVersion` incompatibile | `tabOrder` era ignorato silenziosamente in Power BI Desktop 2.147.x | Segnalare e far verificare la versione di Desktop |

Quando l'utente fornisce l'informazione mancante, applicala tu con `Edit` sul singolo
`visual.json`, rispettando i percorsi e il formato `Literal` descritti in
`references/regole-accessibilita-pbir.md`. Ricordati di salvare senza BOM e di non
toccare `position.z`.

## Le regole di dominio

`references/regole-accessibilita-pbir.md` contiene le regole complete: quali tipi di
visual supportano `altText` e dove, il caso particolare degli `actionButton`, i percorsi
JSON esatti, il criterio dell'ordine di lettura, come si costruiscono i testi.

Leggilo quando devi intervenire a mano su un visual, quando l'utente chiede *perché* la
skill ha fatto una certa scelta, o quando qualcosa nel report non rientra nei casi che
lo script gestisce.
