# AGENTS.md

Contesto operativo per il repository Word2Chm.

## Ambiente

Il .NET 8 SDK è installato in `/home/openhands/.dotnet` e non è nel `PATH`.
`libicu` non è disponibile, quindi il runtime gira in invariant globalization:
va impostato `DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1`.

```bash
export PATH="$HOME/.dotnet:$PATH"
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
- **Suddivisione in pagine**: `BuildOptions.PageLevel` (default 3) è il livello di
  titolo che apre una nuova pagina. A 1 un intero capitolo finisce in un'unica
  pagina HTML lunghissima; a 3 ogni argomento ha la sua pagina. Un titolo che non
  ha testo proprio (etichetta di sezione, forma tipica di Word quando numera ogni
  livello come capitolo) non diventa una pagina ma un'ancora della pagina padre
  (`CollapseContentlessPages`), così il menu non si riempie di pagine di una riga;
  resta invece una pagina se ha un ID di contesto, una voce di indice o un
  segnalibro, perché qualcosa lo collega per nome.
- **Template WinCHM**: con `ConversionOptions.TemplateDirectory` le pagine vengono
  avvolte nello skin (`($title$)`, `($content$)`, `($navigation$)`, `($footer$)`,
  pulsanti prev/next). Gli asset del template vanno copiati accanto ai topic e
  aggiunti a `[FILES]`, altrimenti i pulsanti non compaiono nel CHM. Lo skin linka
  solo il proprio CSS, quindi `help.css` viene iniettato a parte per non perdere la
  formattazione di tabelle e codice.
- I test costruiscono il `.docx` a runtime (`DocxFixture`), quindi non esistono
  fixture binarie da mantenere.
