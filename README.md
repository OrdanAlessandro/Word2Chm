# Word2Chm

Converte un documento Word (`.docx`) in una guida HTML Help (`.chm`) con WinForms
(C# / .NET 8). La struttura del documento viene preservata: titoli, elenchi,
tabelle, immagini, collegamenti e segnalibri.

## Come funziona

- **Una pagina HTML per ogni Titolo 1.** Il livello che genera una nuova pagina è
  configurabile (1-6).
- **ID di contesto espliciti.** Gli ID vengono scritti nel titolo con un marcatore
  `Titolo {#IDH_NOME}`. Il marcatore viene rimosso dal testo visibile. Poiché gli ID
  sono definiti a mano, restano stabili anche quando la struttura o la lingua del
  documento cambiano.
- **Valori numerici dal file `.h` dell'utente.** Il numero non è generato: viene letto
  dal `#define` corrispondente nel header C++ indicato nella GUI (campo "File ID (.h)").
  Lo stesso file può definire gli ID con `#define IDH_X 1000` oppure con un `enum`
  (anche `enum class`), e i due stili possono convivere; in un `enum` un enumeratore
  senza `=` vale il precedente più uno, come per il compilatore C++ (il primo, se non ha
  valore, vale 0), e sono accettate espressioni costanti (`1 << 4`, `IDH_BASE | 0x2`,
  `(2 + 3) * 4`). Un unico header può
  quindi servire tutte le edizioni linguistiche mantenendo gli ID
  identici. I simboli usati nel documento ma assenti dall'header vengono elencati in
  `id-mancanti.txt` e la conversione si completa comunque, così si sistemano tutti in
  un solo passaggio. Un marcatore `{#IDH_NOME=1234}` di vecchi documenti è ancora
  riconosciuto e rimosso, ma il valore inline viene ignorato.
  Attenzione al valore di partenza: HTML Help accetta ID da **1** in su (sono interi
  senza segno a 4 byte) e lo **0 non è valido**, quindi un `enum` che parte da 0 assegna
  0 al primo simbolo e la relativa chiamata `HtmlHelp` non aprirà nulla; conviene farlo
  partire da 1 o più (negli esempi di questo progetto da 1000).
- **Indice `.hhk` dalle voci di indice di Word.** I campi `XE` del documento
  alimentano il file di indice, incluse le sottovoci (`XE "parola" \t "sottovoce"`).

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
| `SkipperQtHelp_IT.chm` | Guida compilata, solo se `hhc.exe` è disponibile |
| `id-mancanti.txt` | Simboli assenti dall'header `.h`, solo se ce ne sono |

Il nome del `.chm` è configurabile (campo "File CHM") e non deriva dal nome base,
così il file caricato dall'applicazione host conserva sempre lo stesso nome.

Con il campo "Copia il .chm in" si indica dove copiare il `.chm` compilato al termine
della conversione: una cartella esistente (o un percorso che termina con il separatore)
riceve il file con il suo nome, altrimenti il valore è il percorso del file di
destinazione e, se manca, l'estensione `.chm` viene aggiunta. Il `.chm` resta comunque
nella cartella di output. Se la copia non riesce (destinazione non scrivibile, disco
pieno) la conversione non viene persa: il file compilato resta nella cartella di output
e il problema compare tra gli avvisi. La copia ha senso solo insieme alla compilazione,
quindi la GUI rifiuta la combinazione "copia impostata" e "compila disattivato".

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
    ContextIdHeaderPath = @"C:\src\helpId.h",          // obbligatorio: fornisce i numeri
    Build = new BuildOptions { PageLevel = 6 },
    TemplateDirectory = @"C:\docs\template\fixedtop",   // opzionale
    ChmCopyPath = @"C:\app\Help\",                       // opzionale: copia il .chm qui
    Compile = new CompileOptions { HhcPath = @"C:\Program Files (x86)\HTML Help Workshop\hhc.exe" },
});
```

`PageLevel` è il livello di titolo che apre una nuova pagina HTML. Il default 6
significa "una pagina per ogni voce del sommario": ogni titolo che compare nel menu
ha il proprio file, quindi il viewer non mostra più un capitolo intero come un unico
documento lungo. Abbassandolo si accorpano i sottotitoli nella pagina del titolo
superiore; con 1 un intero capitolo finisce in un'unica pagina. I titoli che non
hanno testo proprio (etichette di sezione) non generano una pagina ma un'ancora
della pagina padre, così il menu non si riempie di voci di una riga.

`TemplateDirectory` è opzionale: se indicata, le pagine vengono avvolte nello skin
WinCHM (`fixedtop.htm`) con menu di navigazione e pulsanti avanti/indietro. Nella
GUI il percorso viene rilevato automaticamente accanto all'eseguibile.

Nel codice C++:

```cpp
#include "helpId.h"   // lo stesso header fornito alla conversione
HtmlHelp(hwnd, L"SkipperQtHelp_IT.chm", HH_HELP_CONTEXT, IDH_INSTALLAZIONE);
```

## Struttura del repository

- `src/Word2Chm.Core` — parser DOCX, modello intermedio, generatori (HTML, CSS,
  `.hhp`, `.hhc`, `.hhk`) e invocazione di `hhc.exe`.
- `src/Word2Chm.App` — interfaccia WinForms.
- `tests/Word2Chm.Core.Tests` — test sulla pipeline completa; il documento di prova
  viene costruito a runtime, quindi non serve alcun binario di fixture.

## Test

```
dotnet test
```
