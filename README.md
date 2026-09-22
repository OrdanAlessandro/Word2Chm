# Word2Chm

Converte un documento Word (`.docx`) in una guida HTML Help (`.chm`) con WinForms
(C# / .NET 8). La struttura del documento viene preservata: titoli, elenchi,
tabelle, immagini, collegamenti e segnalibri.

## Come funziona

- **Una pagina HTML per ogni Titolo 1.** Il livello che genera una nuova pagina è
  configurabile (1-6).
- **ID di contesto espliciti.** Gli ID vengono scritti nel titolo con un marcatore
  `Titolo {#IDH_NOME}` (oppure `Titolo {#IDH_NOME=1234}` per forzare il numero).
  Il marcatore viene rimosso dal testo visibile. Poiché gli ID sono definiti a mano,
  restano stabili anche quando la struttura o la lingua del documento cambiano.
- **Indice `.hhk` dalle voci di indice di Word.** I campi `XE` del documento
  alimentano il file di indice, incluse le sottovoci (`XE "parola" \t "sottovoce"`).
- **Header C++ `.h`.** Ogni simbolo diventa una costante di compilazione, quindi il
  codice chiamante non dipende mai dai valori numerici generati.

## Output prodotti

Per un documento `guida.docx` vengono generati, nella cartella di output:

| File | Contenuto |
| --- | --- |
| `NNN-titolo.html` | Una pagina per ogni Titolo 1 |
| `index.html` | Pagina iniziale con l'indice |
| `help.css` | Foglio di stile condiviso |
| `assets/*` | Immagini estratte dal documento |
| `guida.hhp` | Progetto di HTML Help Workshop, con `[FILES]`, `[ALIAS]`, `[MAP]` |
| `guida.hhc` | Indice (Contents) |
| `guida.hhk` | Indice analitico (Index), solo se ci sono voci `XE` |
| `guida.h` | Header C++ con gli help ID |
| `guida.chm` | Guida compilata, solo se `hhc.exe` è disponibile |

## Requisiti

- .NET 8 SDK
- `hhc.exe` per compilare effettivamente il `.chm` (HTML Help Workshop, oppure il
  Windows SDK). È facoltativo: senza di esso vengono comunque generati tutti i file
  di progetto.

Il percorso di `hhc.exe` può essere passato esplicitamente oppure viene ricercato
automaticamente nella variabile d'ambiente `HHC_PATH`, nelle cartelle di
installazione di HTML Help Workshop, nel Windows SDK e accanto all'eseguibile.

## Uso

GUI:

```
dotnet run --project src/Word2Chm.App
```

Da codice:

```csharp
var result = new ConversionPipeline().Run(new ConversionOptions
{
    DocxPath = @"C:\docs\guida.docx",
    OutputDirectory = @"C:\docs\guida-chm",
    Build = new BuildOptions { DefaultContextId = 1000, PageLevel = 1 },
    Compile = new CompileOptions { HhcPath = @"C:\Program Files (x86)\HTML Help Workshop\hhc.exe" },
});
```

Nel codice C++:

```cpp
#include "guida.h"
HtmlHelp(hwnd, L"guida.chm", HH_HELP_CONTEXT, IDH_INSTALLAZIONE);
```

## Struttura del repository

- `src/Word2Chm.Core` — parser DOCX, modello intermedio, generatori (HTML, CSS,
  `.hhp`, `.hhc`, `.hhk`, `.h`) e invocazione di `hhc.exe`.
- `src/Word2Chm.App` — interfaccia WinForms.
- `tests/Word2Chm.Core.Tests` — test sulla pipeline completa; il documento di prova
  viene costruito a runtime, quindi non serve alcun binario di fixture.

## Test

```
dotnet test
```
