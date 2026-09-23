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
- I test costruiscono il `.docx` a runtime (`DocxFixture`), quindi non esistono
  fixture binarie da mantenere.
