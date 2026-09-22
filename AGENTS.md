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
- I test costruiscono il `.docx` a runtime (`DocxFixture`), quindi non esistono
  fixture binarie da mantenere.
