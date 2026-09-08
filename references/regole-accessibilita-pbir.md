# Regole di accessibilità PBIR — riferimento completo

Regole implementate da `scripts/pbip_a11y.py`. Leggile quando devi intervenire a mano su
un visual (tipicamente per risolvere un'anomalia) o spiegare una scelta dello script.

## Indice

1. [Proprietà modificabili e percorsi JSON](#1-proprietà-modificabili-e-percorsi-json)
2. [Quali tipi di visual supportano altText](#2-quali-tipi-di-visual-supportano-alttext)
3. [Tab order](#3-tab-order)
4. [Alt text](#4-alt-text)
5. [Pulsanti: altText e tooltip](#5-pulsanti-alttext-e-tooltip)
6. [Gruppi](#6-gruppi)
7. [Codifica dei file](#7-codifica-dei-file)
8. [Verifica di compatibilità](#8-verifica-di-compatibilità)

---

## 1. Proprietà modificabili e percorsi JSON

Si modificano **esclusivamente** queste proprietà, e solo dove supportate:

| Proprietà | Percorso nel `visual.json` | Si applica a |
|---|---|---|
| alt text | `visual.visualContainerObjects.general[0].properties.altText` | tutti i visualType core, esclusi i decorativi |
| alt text di gruppo | `visualGroup.objects.general[0].properties.altText` | solo `visualGroup` |
| tooltip di navigazione | `visual.visualContainerObjects.visualLink[0].properties.tooltip` | solo `actionButton` |
| ordine di tabulazione | `position.tabOrder` | tutti |

Formato dei valori testuali (`Literal` PBIR):

```json
{
  "altText": {
    "expr": { "Literal": { "Value": "'testo alternativo in italiano'" } }
  }
}
```

Gli apici singoli interni vanno raddoppiati (`''`). Lo script preferisce sostituirli con
l'apostrofo tipografico `’`, così il `Literal` resta leggibile e non ci sono ambiguità di
escaping.

**Cose che non si fanno mai:**

- non si modifica `position.z` — vedi [Tab order](#3-tab-order) per il perché;
- non si usa `tooltipText`: è un nome di proprietà che in Power BI **non esiste**
  (probabile refuso storico per `tooltip`);
- non si sovrascrive un `altText` o un `tooltip` già presente, qualunque sia il contenuto:
  chi l'ha scritto sapeva cosa stava descrivendo;
- non si rinominano i visual, non si toccano posizioni, dimensioni, query, dati, modello
  semantico;
- non si riformatta JSON non coinvolto nelle modifiche;
- non si creano file nuovi dentro il progetto.

---

## 2. Quali tipi di visual supportano altText

`altText` in `visualContainerObjects.general` è supportato dalla maggior parte dei visual
core di Power BI. Le eccezioni verificate empiricamente:

**Tipi decorativi — mai un `altText` descrittivo, sempre la stringa vuota:**

- `image` — immagini e loghi
- `shape` — forme decorative
- `basicShape` — forme base decorative

Per questi: `tabOrder` = `-9999000`, `altText` = `""` **sempre** — anche quando la
proprietà non è presente nel file — e `z` invariato.

La stringa vuota va scritta anche se manca, e la ragione è che i due meccanismi di
accesso sono distinti. `tabOrder = -9999000` toglie l'elemento dall'ordine di tabulazione,
quindi il Tab non ci arriva mai. Ma uno screen reader non naviga solo col Tab: in modalità
di lettura percorre la pagina elemento per elemento, decorativi compresi. Un logo senza
`altText` in quella modalità viene comunque annunciato, tipicamente leggendo il nome del
file dell'immagine — `LOGO_ARL2734654422224455.png` letto ad alta voce. `altText: ""` è il
modo dichiarativo di dire "elemento decorativo, ignoralo", ed è la stessa convenzione di
`alt=""` in HTML.

> **Verifica ancora aperta (al 2026-09-02).** Una versione precedente di queste regole
> elencava `image`, `shape` e `basicShape` fra i tipi che *non* supportano `altText` in
> `general`, con la motivazione che l'aggiunta causa errori in Power BI Desktop. Quella
> affermazione non è mai stata documentata con una prova, e la stessa lista si era già
> rivelata imprecisa per `actionButton` (vedi sotto), quindi la regola attuale scrive
> `altText: ""` sui decorativi. Resta però da confermare aprendo in Power BI Desktop un
> report modificato: se Desktop segnala un errore di parsing su un `image` o uno `shape`,
> la regola va rivista e la riga corrispondente in `pbip_rules.py` (il confronto
> `if v.alt_text != "":`) va riportata a `if v.has_alt_text and v.alt_text != "":`.
> Finché la verifica non è fatta, vale la pena controllare il report in Desktop dopo un
> `apply` su progetti che contengono molte immagini o forme.

**Attenzione:** una `shape` o `basicShape` con un `visualLink` (navigazione o altra azione)
è un pulsante mascherato. Va trattata come elemento interattivo e **non** esclusa dal focus.

**Percorsi con logica diversa:**

- `textbox` — vedi [Alt text](#4-alt-text), l'alt text si ricava dal testo visibile
- `visualGroup` — percorso `visualGroup.objects.general[0].properties.altText`

**`actionButton` — caso particolare (verificato empiricamente il 2026-08-04):** supporta
`altText` in `visualContainerObjects.general.properties.altText`, con lo stesso formato
degli altri visual. Un `actionButton` può avere **entrambe** le proprietà, con scopi
diversi: l'`altText` descrive il pulsante, il `tooltip` in `visualLink` è la descrizione
comando annunciata al momento dell'azione. Non è un duplicato.

**Visual custom (marketplace/AppSource):** se il `visualType` è una stringa lunga simile a
un GUID o comunque non riconoscibile come visual Power BI standard, **non aggiungere
`altText` in automatico**. Le capabilities dei visual custom non sono garantite e
un'aggiunta non supportata rompe il report. Segnalalo come anomalia da verificare a mano.

**Tutti gli altri visualType** (grafici, card, KPI, matrici, tabelle, slicer, mappe…)
supportano `altText` in `visualContainerObjects.general.properties`.

### Limite noto dell'implementazione: elenco chiuso di tipi

La regola qui sopra dice "tutti gli altri", ma `pbip_rules.py` non funziona così: tiene un
**elenco esplicito** (`ALT_TEXT_SUPPORTED`) dei tipi su cui scrive l'`altText`, e ogni tipo
fuori da quell'elenco finisce fra le anomalie con il messaggio *"tipo non presente negli
elenchi dei tipi supportati"*.

La scelta è prudenziale: aggiungere `altText` a un tipo che non lo supporta rompe il report
in Power BI Desktop, e un falso negativo (una segnalazione da sistemare a mano) costa molto
meno di un falso positivo (un report che non si apre più). Ma ha un prezzo: i tipi core non
ancora inseriti nell'elenco vengono segnalati come anomalie pur essendo perfettamente
trattabili.

Tipi attualmente nell'elenco: `actionButton`, `card`, `cardVisual`, `barChart`,
`clusteredBarChart`, `clusteredColumnChart`, `columnChart`, `donutChart`, `slicer`,
`textbox`, `lineChart`, `lineStackedColumnComboChart`, `lineClusteredColumnComboChart`,
`hundredPercentStackedColumnChart`, `pageNavigator`, `shapeMap`, `tableEx`, `pivotTable`,
`matrix`.

Tipi core diffusi **non** ancora coperti, che compariranno fra le anomalie: `pieChart`,
`areaChart`, `stackedAreaChart`, `scatterChart`, `funnel`, `gauge`, `waterfallChart`,
`treemap`, `map`, `filledMap`, `multiRowCard`, `kpi`, `ribbonChart`,
`decompositionTreeVisual`, `keyDriversVisual`, `qnaVisual`.

Quando incontri un'anomalia di questo tipo, la strada corretta è verificare il tipo in
Power BI Desktop e aggiungerlo a `ALT_TEXT_SUPPORTED` e a `ITALIAN_TYPE_NAMES` — non
bypassare la segnalazione scrivendo l'`altText` a mano su decine di file, perché così il
prossimo scan la ritroverà identica.

---

## 3. Tab order

### Ordine di lettura

I visual di ogni pagina si ordinano **da sinistra a destra, dall'alto verso il basso**,
sulla base delle coordinate `x` e `y`.

Le righe si individuano per sovrapposizione verticale: due visual stanno sulla stessa riga
se le loro estensioni verticali si sovrappongono per almeno metà dell'altezza del più
basso. Dentro ogni riga si ordina per `x`. Questo criterio è più robusto del semplice
ordinamento per `y`, perché nei report reali gli elementi di una stessa fascia non sono
mai allineati al pixel.

Il **primo elemento a ricevere il focus** è il titolo della pagina, se presente: una
`textbox` nella prima riga in alto.

### Gruppi

Un gruppo (`visualType: "visualGroup"`, cioè un `visual.json` con la chiave `visualGroup`)
conta come **unità singola**:

- nell'ordinamento si usa la posizione del gruppo, non quella dei figli;
- il `tabOrder` numerico si assegna **solo** al gruppo;
- i figli (quelli con `parentGroupName`) mantengono il `tabOrder` che hanno.

L'unica eccezione: un elemento decorativo va nascosto dalla tastiera anche se è figlio di
un gruppo. Quella è una questione di stato nascosto, non di ordine numerico.

### Valori numerici

`tabOrder` funziona in **ordine decrescente**: il valore più alto riceve il focus per
primo. Si assegnano multipli di 1000, così restano spazi liberi per inserimenti futuri.

Esempio con 5 visual:

| Ordine di lettura | Visual | tabOrder |
|---|---|---|
| 1° | Titolo pagina | 5000 |
| 2° | KPI principale | 4000 |
| 3° | Grafico | 3000 |
| 4° | Tabella | 2000 |
| 5° | Filtro | 1000 |

### Esclusione dalla tastiera

Per gli elementi decorativi si usa `tabOrder: -9999000`, il valore sentinella con cui
Power BI Desktop esclude davvero un visual dall'ordine di tabulazione (pannello Selezione
→ Ordine di tabulazione → icona occhio barrato). Verificato empiricamente il 2026-08-05.

`tabOrder: 0` **non** basta: lascia l'elemento raggiungibile da tastiera, semplicemente
per ultimo.

### Perché `z` non va toccato

`position.z` determina l'ordine di sovrapposizione visiva. L'ordine di tabulazione in
Power BI **non dipende da `z`**: correggere `tabOrder` non richiede alcuna modifica a `z`.

Riassegnare `z` "in sincrono" con `tabOrder` inverte lo stacking degli elementi sovrapposti
e produce difetti visivi difficili da individuare. Casi osservati su report reali:

- una textbox sovrapposta a un `actionButton` (pattern overlay: etichetta sopra pulsante
  con fill) finisce **sotto** il pulsante e il testo scompare;
- un pulsante o un'etichetta finisce sotto un altro visual sovrapposto e diventa
  invisibile, anche senza un `fill` esplicito.

L'ordine di lettura di una pagina in generale **non coincide** con l'ordine di stacking
progettato: qualunque riassegnazione meccanica di `z` è potenzialmente distruttiva.

---

## 4. Alt text

Tutti i testi alternativi vanno scritti **in italiano**, max **150 caratteri** (oltre, si
troncano ai primi 147 più `...`).

### Visual grafici (grafici, KPI, matrici, tabelle, scatter…)

Il testo nomina il tipo di grafico in italiano, la misura principale e la dimensione, e
dice a quale domanda il visual risponde.

Il soggetto semantico si ricava con questa gerarchia:

1. titolo esplicito e descrittivo — non un UUID, non `visualContainer1`, non un nome
   interno con underscore tipo `Testo_Donazioni`, non `Group 2`;
2. nomi di misure e dimensioni dal blocco `query.queryState` (ruoli misura: `Values`, `Y`,
   `Y2`, `Value`, `Data`, `Size`; ruoli dimensione: `Category`, `Axis`, `Rows`, `Series`,
   `Columns`, `Details`, `Legend`);
3. se non si ricava nulla: **anomalia**, nessun alt text inventato.

Il verbo cambia con il tipo, perché descrive come si legge il visual: i grafici di
confronto *confrontano*, quelli temporali *mostrano un andamento*, le ciambelle *mostrano
una distribuzione*, e i tipi tabellari (`matrix`, `tableEx`, `pivotTable`) *riepilogano* —
chi usa lo screen reader li percorre cella per cella, non ne ricava una forma d'insieme,
quindi "riepiloga" descrive meglio cosa troverà.

Esempi corretti:

- `Grafico a barre che confronta il fatturato per regione nell'anno FY2025.`
- `Grafico a linee che mostra l'andamento mensile delle vendite nel tempo.`
- `Matrice che riepiloga il numero di clienti per categoria di prodotto.`
- `Tabella che riepiloga il dettaglio degli impianti controllati.`
- `KPI che mostra il margine lordo complessivo del periodo selezionato.`

Da evitare: `Grafico con dati` (generico), `Visual 1` (non descrittivo), descrizioni oltre
i 150 caratteri.

### Slicer

Il titolo di uno slicer è **inaffidabile**: è spesso un residuo di copia-incolla non
aggiornato dopo il rebind del campo. Caso reale verificato il 2026-08-05: lo stesso titolo
"Filtro tipo contratto" su slicer bindati a `sesso_label`, `eta_label`, `TIPO_CONTRATTO`.

Per gli slicer si ignora sempre il titolo e si usa il campo dati effettivamente bindato:
`Filtro per selezionare <campo>.`

### Caselle di testo (textbox)

Le textbox devono **sempre** avere un alt text. Il testo si ricava concatenando, separati
da uno spazio, tutti i valori:

```
visual.objects.general[0].properties.paragraphs[].textRuns[].value
```

Questo è il testo che l'utente vede. **Non** si legge `visualContainerObjects.title`:
contiene un nome interno come `'Testo_Donazioni'`, non il testo visibile.

L'alt text va in `visual.visualContainerObjects.general[0].properties.altText`.

Se una textbox non ha testo leggibile nei `textRuns`: anomalia.

### pageNavigator

Testo fisso: `Barra di navigazione tra le pagine del report.` (elemento di navigazione,
non ha dati sottostanti).

### Immagini decorative e loghi

Non ricevono il focus (vedi [Tab order](#3-tab-order)). L'`altText` va portato a stringa
vuota `""` — sempre, sia che sia già presente con un contenuto, sia che manchi del tutto —
così lo screen reader li ignora anche in modalità di lettura, dove il `tabOrder` non ha
effetto. La motivazione completa è nella sezione
[Quali tipi di visual supportano altText](#2-quali-tipi-di-visual-supportano-alttext).
Non si aggiunge nessun tooltip.

---

## 5. Pulsanti: altText e tooltip

Un `actionButton` ha due proprietà distinte, generate con logiche diverse.

### Alt text del pulsante

`visual.visualContainerObjects.general[0].properties.altText`

- se il pulsante ha un testo visibile in `visual.objects.text[].properties.text`:
  `Pulsante <testo visibile>.`
- altrimenti si usa la destinazione ricavata dal `visualLink`: `Pulsante: <destinazione>.`
- se non si determina né testo visibile né destinazione: anomalia.

### Tooltip di navigazione (descrizione comando)

`visual.visualContainerObjects.visualLink[0].properties.tooltip` — dentro lo stesso blocco
`properties` che contiene già `show`, `type` e `navigationSection`, **non** in `general`.

Struttura di un pulsante di navigazione:

```json
{
  "visualLink": [
    {
      "properties": {
        "show": { "expr": { "Literal": { "Value": "true" } } },
        "type": { "expr": { "Literal": { "Value": "'PageNavigation'" } } },
        "navigationSection": { "expr": { "Literal": { "Value": "'ReportSectionXXXXX'" } } },
        "tooltip": { "expr": { "Literal": { "Value": "'Vai alla pagina Donazioni'" } } }
      }
    }
  ]
}
```

Il valore di `navigationSection` è l'ID della pagina di destinazione: il nome leggibile si
trova nel `displayName` del `page.json` della cartella corrispondente (o in
`pages/pages.json`).

Il `title` del pulsante in `visualContainerObjects.title[0].properties.text` contiene un
nome interno (es. `'Pulsante_Donazioni'`): **non usarlo** come testo del tooltip.

Testi per tipo di azione:

| `type` del visualLink | Tooltip |
|---|---|
| `PageNavigation` | `Vai alla pagina <nome pagina>` |
| `PageNavigation` verso la pagina corrente | `<nome pagina> (pagina corrente)` |
| `Back` | `Torna alla pagina precedente` |
| `Bookmark` | `Attiva la vista <testo visibile>` |
| `WebUrl` / `WebHyperlink` | `Apri il collegamento <testo visibile>` |
| assente o non riconosciuto | anomalia, nessun tooltip |

### `tooltipText` nel posto sbagliato

Se trovi `tooltipText` in `visualContainerObjects.general` su un `actionButton`: è una
proprietà inesistente. Va rimossa e la descrizione spostata in
`visualLink[0].properties.tooltip`.

Un `altText` presente in `general` su un `actionButton` è invece **valido**: non si tocca.

---

## 6. Gruppi

L'`altText` di un gruppo va in `visualGroup.objects.general[0].properties.altText` e deve
descrivere lo scopo o il contenuto del gruppo, **non** limitarsi al suo nome. Si ricava
leggendo i visual figli (`parentGroupName` uguale al `name` del gruppo) e rispondendo alla
domanda: cosa comunica questo gruppo all'utente?

Gli elementi decorativi dentro il gruppo non contribuiscono alla descrizione.

Esempi corretti:

- `Sezione KPI: riepilogo di fatturato, margine e numero ordini per il periodo selezionato.`
- `Pannello filtri: consente di selezionare anno, regione e categoria di prodotto.`
- `Sezione andamento: grafico a linee e tabella di dettaglio delle vendite mensili.`

Da evitare: `Gruppo 1`, `group_abc123` — nomi autogenerati, privi di significato.

Se i figli non forniscono informazioni utilizzabili: anomalia.

---

## 7. Codifica dei file

Tutti i JSON modificati vanno salvati in **UTF-8 senza BOM**.

Con il BOM, Power BI Desktop può non riconoscere il file o introdurre caratteri invisibili
a inizio file, causando errori di parsing o di visualizzazione del report. Se lo strumento
usato per scrivere aggiunge un BOM di default, va rimosso esplicitamente prima del
salvataggio, e verificato dopo controllando che i primi byte non siano `EF BB BF`.

Indentazione, ordine delle chiavi, fine riga e tutto il JSON non coinvolto nelle modifiche
vanno preservati: un diff pulito è l'unico modo per far rivedere le modifiche a un umano.

---

## 8. Verifica di compatibilità

Prima di modificare, si legge il campo `version` di `definition.pbir` (nella cartella
`*.Report`).

- `version` assente o non leggibile → anomalia.
- `version < 4.0` → anomalia. In Power BI Desktop 2.147.x la proprietà `tabOrder` era
  ignorata silenziosamente: va verificata la versione di Desktop in uso.
- Formato PBIR-legacy (`report.json` monolitico, senza `definition/pages`) → non
  supportato dallo scan per-visual. Il progetto va risalvato da Power BI Desktop nel
  formato PBIR avanzato.

Le anomalie di compatibilità non bloccano lo scan: vengono registrate e riportate.
