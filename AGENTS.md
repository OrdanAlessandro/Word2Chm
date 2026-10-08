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
- **`hhc.exe` esce con codice 1 quando la compilazione riesce** (il suo flag interno
  "file scritto" finisce nell'exit code senza essere invertito), quindi l'esito **non**
  si decide con `ExitCode == 0`: `ChmCompiler` cancella il `.chm` precedente e considera
  riuscita la compilazione se il file esiste dopo l'esecuzione. Il test
  `CopiesTheChmEvenWhenHhcExitsWithANonZeroCode` blinda questo comportamento usando un
  finto `hhc` che esce con 1 ma scrive comunque il file. Per lo stesso motivo un `.chm`
  stantio nell'output non deve far sembrare riuscita una compilazione fallita: da qui la
  cancellazione preventiva e il test `IgnoresAStaleChmLeftByAnEarlierRun`.

## Note di progettazione

- **Invariant globalization**: `CultureInfo.GetCultures` e la normalizzazione
  Unicode (`String.Normalize`) non sono affidabili in questo ambiente. Per gli
  slug usare `Slugger.Deaccent`, che usa una mappa di accenti esplicita; per le
  lingue usare la mappa LCID in `ChmProjectGenerator`.
- **ID di contesto**: l'utente scrive il simbolo nel titolo (`Titolo {#IDH_NOME}`) e
  fornisce con la GUI (`ContextIdHeaderPath`, campo "File ID (.h)") l'header C++ che
  definisce i numeri. Il `.h` è l'unica fonte dei valori: così un solo header serve
  tutte le edizioni linguistiche e gli ID restano stabili, che è il contratto verso il
  codice C++ che chiama `HtmlHelp`. Un vecchio `{#IDH_NOME=1234}` è ancora riconosciuto
  e rimosso dal testo visibile, ma il valore inline è ignorato. Un simbolo assente
  dall'header non ferma la conversione: finisce in `BuildOptions.MissingSymbols` e viene
  elencato in `id-mancanti.txt`, senza voce `[MAP]` (il compilatore la accetta).
- **L'header degli ID può usare `#define` o `enum`**: `ContextIdHeader.Parse` legge
  entrambi e i due stili possono convivere nello stesso file; vince l'ultimo valore visto
  in ordine di file (come per il compilatore). In un `enum` (anche `enum class`) un
  enumeratore senza `=` vale il precedente più uno, a partire da 0 (il primo enumeratore
  implicito vale quindi 0, come in C++); sono valutate le
  espressioni costanti usuali (letterali dec/hex/bin/ottali, riferimenti a simboli già
  noti, `<< >> & ^ | + - * / % ~` e parentesi) con la giusta precedenza. Un'espressione
  non risolvibile (cast, chiamate, simboli definiti più avanti) **non** viene indovinata:
  il simbolo resta assente e finisce tra gli ID mancanti; anche il successivo run di
  enumeratori impliciti resta ignoto finché un valore esplicito non lo riavvia. I commenti
  `//` e `/* ... */` sono rimossi prima dell'analisi, quindi un esempio dentro un commento
  non conta come definizione. Un valore che non entra in `int` lancia `OverflowException`.
- **Intervallo ammesso per un ID di contesto HTML Help**: la specifica HTML Help archivia
  gli ID come interi senza segno a 4 byte, quindi l'intervallo valido è **1–4294967295**;
  lo **0 non è un ID valido** e non viene risolto dal visualizzatore (`HH_HELP_CONTEXT`
  con 0 non apre nulla). Il tool non impone questo limite: uno 0 definito nell'header viene
  scritto in `[MAP]` come qualunque altro valore. Se un `enum` parte da 0, il primo simbolo
  ottiene 0 e la sua chiamata `HtmlHelp` non funzionerà, quindi conviene far partire gli ID
  da 1 o più (il caso tipico negli esempi del progetto è 1000).
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
- **Nomi dei file di output**: il nome del `.chm` (`SkipperQtHelp_IT.chm`) è
  parametrizzato (`ConversionOptions.ChmFileName`, campo "File CHM" nella GUI) e **non**
  deriva dal nome base: il programma host carica il CHM per nome, quindi non deve cambiare
  a ogni conversione. Separatori di percorso e caratteri non validi diventano `_`, così un
  nome tipo `../fuori.chm` resta dentro la cartella di output; se manca l'estensione viene
  aggiunta. L'header `.h` non è più generato: è un **input** dell'utente, e il suo percorso
  è obbligatorio nella GUI (vedi la nota sugli ID di contesto).
- **Copia del `.chm` fuori dalla cartella di output**: `ConversionOptions.ChmCopyPath`
  (campo "Copia il .chm in" nella GUI) copia il file compilato dove lo legge
  l'applicazione host. Una destinazione che è già una cartella riceve il file con il suo
  nome, altrimenti il valore è il percorso del file e l'estensione `.chm` viene aggiunta se
  manca. La copia avviene **dopo** la compilazione e solo se la compilazione è riuscita; un
  errore di copia non annulla nulla (il `.chm` resta nella cartella di output e il problema
  finisce in `HelpDocument.Warnings`, mostrato nel log). Il percorso effettivo è in
  `ConversionResult.CopiedChmPath`. La GUI rifiuta la copia con "Compila il .chm" disattivato,
  perché non ci sarebbe alcun file da copiare.
- **Corsivo: lo stile carattere è un *toggle*, non una dichiarazione**. Un `w:rStyle`
  che porta `<w:i/>` non significa "corsivo": *inverte* il corsivo ereditato dal
  paragrafo. In `SkipperQt_IT.docx` i paragrafi `Didascalia` (corsivo) con run `Enfasicorsivo`
  risultano *diritti*, e un run con `<w:i/>` proprio li rende di nuovo corsivi. Trattando
  lo stile come "italic = true" si corsivano 154 paragrafi che Word mostra in tondo
  (es. "Quando vengono modificati..." in "Salvare i parametri nella memoria del CN").
  Precedenza effettiva: paragrafo → stile carattere (toggle, oppure off se `<w:i w:val="0"/>`)
  → `<w:i>`/`<w:i w:val="0"/>` diretto sul run. `w:rPr` dentro `w:pPr` non è ereditato
  dai run, quindi non va usato come corsivo del paragrafo.
- **Come verificare la formattazione senza Word**: LibreOffice è installato e genera un
  render di riferimento. Serve `LD_LIBRARY_PATH=/usr/lib/libreoffice/program` (senza,
  `soffice.bin` non trova `libreglo.so`) e un profilo utente dedicato:
  `LD_LIBRARY_PATH=/usr/lib/libreoffice/program soffice.bin -env:UserInstallation=file:///tmp/louser --headless --convert-to html --outdir /tmp/lorender sample/SkipperQt_IT.docx`.
  Nel render, `<em class="western"><span style="font-style: normal">` = corsivo annullato.
  Confrontare i modelli sui paragrafi del documento reale (non su casi inventati) è ciò
  che ha smascherato il toggle: 62% di accordo per la lettura solo-run, 97% per lo stile
  come verità, 99,6% per il toggle.
