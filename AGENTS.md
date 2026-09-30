# AGENTS.md

Contesto operativo per il repository Word2Chm.

## Ambiente

Il .NET 8 SDK è installato in `/workspace/.dotnet` e non è nel `PATH`.
`libicu` non è disponibile, quindi il runtime gira in invariant globalization:
va impostato `DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1`.

Attenzione: l'ambiente viene azzerato periodicamente e `$HOME` non è stabile, quindi
l'SDK va tenuto fuori da `$HOME` (in `/workspace/.dotnet`, che sopravvive al reset)
e reinstallato con:

```bash
curl -sSL https://dot.net/v1/dotnet-install.sh | bash -s -- --channel 8.0 --install-dir /workspace/.dotnet --no-path
```

```bash
export PATH="/workspace/.dotnet:$PATH"
export DOTNET_ROOT=/workspace/.dotnet
export DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1
```

## Comandi

```bash
dotnet build Word2Chm.sln -c Release   # build completa
dotnet test                            # test della pipeline
```

I warning `NETSDK1188` sul locale `cs`/`de`/`it` ecc. provengono dalle risorse del
pacchetto `Microsoft.TestPlatform.TestHost` e non sono problemi del codice:
si filtrano con `grep -v NETSDK1188`.

## Vincoli di piattaforma

- `src/Word2Chm.App` è WinForms (`net8.0-windows`). Compila su Linux grazie a
  `EnableWindowsTargeting`, ma non può essere eseguito qui.
- La compilazione del `.chm` richiede `hhc.exe` (solo Windows). Senza di esso la
  pipeline genera comunque tutti i file di progetto.

## Note di progettazione

- **Invariant globalization**: `CultureInfo.GetCultures` e la normalizzazione
  Unicode (`String.Normalize`) non sono affidabili in questo ambiente. Per gli
  slug usare `Slugger.Deaccent`, che usa una mappa di accenti esplicita; per le
  lingue usare la mappa LCID in `ChmProjectGenerator`.
- **ID di contesto**: definiti dall'utente nei titoli come `Titolo {#IDH_NOME}`
  (o `{#IDH_NOME=1234}`). Vanno mantenuti stabili: sono il contratto verso il
  codice C++ che chiama `HtmlHelp`.
- **ID di contesto su sottotitoli**: `[ALIAS]` accetta solo un file, non
  `file.htm#anchor`: `hhc.exe` cercherebbe un file con quel nome letterale e
  segnala `HHC3015 ... the file does not exist`, senza però fallire la
  compilazione. Un marcatore su un titolo sotto il livello di pagina finisce
  quindi in `HelpPage.Anchors` e riceve una piccola pagina di reindirizzamento
  (`<nome>-<ancora>-id.html`) che porta all'ancora; lo stub va elencato in
  `[FILES]`.
- **Codifica dei file letti da `hhc.exe`**: `.hhc`, `.hhk` e anche i file HTML
  vanno ASCII puri. `hhc.exe` li legge come ANSI, quindi i caratteri UTF-8
  (apostrofi tipografici, trattini lunghi) vanno emessi come entità numeriche.
  In HTML lo fa `HtmlGenerator.ToAsciiSafe`.
- **Riga `[WINDOWS]` del `.hhp`**: i campi sono posizionali e il viewer li legge
  senza margini di errore. `WindowStyles` deve restare al campo 9 e
  `NavigationPaneStyle` all'11; una virgola in piu' sposta `0x23520` sul campo
  della geometria della finestra e il CHM non si apre piu' ("There is not enough
  memory available for this task"), pur compilando senza errori. Il test
  `HhpKeepsWindowsFieldsAligned` blinda l'allineamento.
- **Campi XE**: Word spezza l'istruzione di un campo su più run
  (` XE "` + parola chiave + `" `), spesso annidata in un campo `HYPERLINK`.
  Vanno concatenati i `FieldCode` tra `fldChar begin/end`; leggere un solo run
  produce zero voci di indice.
- **File di indice**: `.hhk` va dichiarato sia come `Index file=` sia nella lista
  `[FILES]` del `.hhp`, altrimenti le parole chiave non entrano nel CHM.
- **`Window Styles` del `.hhc`**: il riquadro di navigazione è una tree view
  Win32, quindi il valore va letto come stile `TVS_*`. Servono `TVS_HASBUTTONS`
  (0x1), `TVS_HASLINES` (0x2) e `TVS_LINESATROOT` (0x4); `TVS_CHECKBOXES` (0x100)
  va lasciato spento. Il vecchio valore `0x23520` azzerava i primi tre e accendeva
  quello delle caselle: nel CHM comparivano checkbox e nessun pulsante `+`.
  Il valore corretto è `0x27`, come nei progetti reali. Il test
  `HhcEnablesTreeLinesInsteadOfCheckBoxes` blinda i quattro bit.
- **Suddivisione in pagine**: `BuildOptions.PageLevel` (default 6, cioè
  `DefaultPageLevel`) è il livello di titolo che apre una nuova pagina. Il default
  punta a "una pagina per ogni voce del sommario": ogni titolo presente nel menu ha
  il proprio file. Con 1 l'intero capitolo finisce in una sola pagina HTML lunghissima
  e il menu resta l'unico modo per muoversi dentro. Attenzione: la GUI salvava
  `PageLevel` in `settings.json`, quindi un valore 1 scritto da una build precedente
  continuava a vincere sul nuovo default finché `AppSettings` non ha introdotto un
  numero di versione che lo scarta.
- **Un titolo che non ha testo proprio** (etichetta di sezione, forma tipica di Word
  quando numera ogni livello come capitolo) non diventa una pagina ma un'ancora della
  pagina padre (`CollapseContentlessPages`), così il menu non si riempie di pagine di
  una riga; resta invece una pagina se ha un ID di contesto, una voce di indice o un
  segnalibro, perché qualcosa lo collega per nome.
- **Template WinCHM**: con `ConversionOptions.TemplateDirectory` le pagine vengono
  avvolte nello skin (`($title$)`, `($content$)`, `($navigation$)`, `($footer$)`,
  pulsanti prev/next). Gli asset del template vanno copiati accanto ai topic e
  aggiunti a `[FILES]`, altrimenti i pulsanti non compaiono nel CHM. Lo skin linka
  solo il proprio CSS, quindi `help.css` viene iniettato a parte per non perdere la
  formattazione di tabelle e codice.
- I test costruiscono il `.docx` a runtime (`DocxFixture`), quindi non esistono
  fixture binarie da mantenere. Il documento reale del cliente è `sample/SkipperQt_IT.docx`
  (285 pagine, 96 media di cui 45 in VML): è la fixture di riferimento per le prove end-to-end.
- **Elenchi**: paragrafi consecutivi con lo stesso `numId` vanno uniti in un unico
  `ListBlock`, altrimenti ogni voce genera un `<ol>` separato e la numerazione riparte
  da 1. Un cambio di `numId` chiude il blocco: liste diverse non devono condividere il
  contatore. I paragrafi con campi `XE` o segnalibri restano fuori dal raggruppamento,
  perché altrimenti perdono l'indice e le destinazioni dei link.
- **Immagini VML**: oltre ad `a:blip` (DrawingML) vanno lette `v:imagedata` dentro
  `w:object` e `w:pict` (immagini incollate e OLE); sono la maggioranza in `SkipperQt_IT.docx`.
  La ricerca va fatta sull'ambito del *figlio* del run (`child.Descendants<...>`), non
  dell'intero run: `RunProperties` cade nel ramo `default` e una scansione su `run`
  emetteva la stessa immagine due volte.
- **Titolo di pagina duplicato**: lo skin mostra `($title$)` nell'header, quindi il
  primo blocco della pagina (l'`HeadingBlock` che l'ha aperta) va saltato e sostituito
  con uno `<span id="...">` vuoto, perché sommario e indice continuano a linkare quell'ancora.
  Si confronta la *posizione* (`index == 0`), non `page.Anchor`, che non è popolato.
- **Icone del sommario**: senza `ImageNumber` esplicito hh.exe usa il punto interrogativo
  per le foglie. `1` = libro/contenitore, `11` = documento/foglia (confermato da `.hhc`
  reali Doxygen). Va emesso sia nel `.hhc` sia nel `.hhk` (anche per i nodi padre).
- **Ordine prev/next**: `FlattenPages` deve deduplicare sulla chiave *file* (senza
  `#anchor`). Usando `node.Local` grezzo ogni pagina entrava due volte nell'ordine
  (570 voci per 285 pagine) e la prima pagina puntava all'ultima.
- **Piè di pagina**: `BuildOptions.Footer` → `HelpDocument.Footer` (default
  `HelpDocument.DefaultFooter`, `© 2026 FARO srl. All rights reserved.`), configurabile
  dalla GUI e codificato in entità ASCII prima di finire nell'HTML.
- **Bordi**: le tabelle usano `#000000` fisso (non `--border`, che è grigio chiaro) per
  rispecchiare lo stile griglia di Word. Il CSS dello skin azzera il bordo delle `<img>`
  dei pulsanti, che Internet Explorer (il motore del visualizzatore CHM) disegna sui link.
