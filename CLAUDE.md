# CLAUDE.md

Knowledge-projekt: avigord
Dokumentation, beslut och agentlogg finns i Knowledge — se ~/.claude/CLAUDE.md. Paketet hör
till Avig.Ord-projektet, eftersom det bröts ut ur Avig.Kalkyl för Avig.Ords skull (beslut
"Avig.Ord — Beslut: Gemensamt paket Avig.OpenXml för koden som delas med Avig.Kalkyl").

## Vad det här är

NuGet-paketet Avig.OpenXml: paketlagret (OPC) under Avig.Kalkyl och Avig.Ord. Läser och skriver
zip-paket och Flat OPC: delar, content types och relationer, med gränser mot fientliga filer.
Det vet inget om kalkylblad eller dokument.

## Kommandon

```powershell
dotnet build Avig.OpenXml.slnx
dotnet test Avig.OpenXml.slnx
./publish.ps1 -NoBump          # testar, packar, formkontrollerar, lägger i G:\LocalNugetRepo
```

## Regler

- **Inga beroenden utom Avig.Result.** `publish.ps1` failar annars.
- **Ändra inget utan att köra Avig.Kalkyls och Avig.Ords tester mot det nya paketet.** Båda
  paketens felkoder läser av meddelandena härifrån: `OpcErrors.DamagedPrefix` ("The file is
  damaged") och `OpcErrors.LimitMarker` ("the limit set in "). De får inte ändras.
- **Förväntade fel är `Result`**, aldrig null (AVIG0007). Saknas något: `Try…`-mönster eller
  `Result.NotFound`.
- **XML utan DTD och med djupgräns**: allt går via `OpcXml.CreateReader`.
- **Varje del som läses går via `OpcZip.OpenEntryAsync`** (eller `OpcPackage.OpenPartAsync`), så
  att den hålls till storleken i sitt zip-huvud.
- **Zip-filen ska säga en sak** (`ZipStructure`, `OpcOpenOptions.StrictZip`): lokala huvuden som
  katalogen, inga byte före, mellan eller efter posterna. Namn som inget paket får ha avgörs av
  `OpcPath.IsSafe`, för både zip och Flat OPC, och dubbletter jämförs med `OpcPath.ComparisonKey`.
  Skärp aldrig reglerna utan att köra dem över riktiga filer (Kalkyls kompatibilitetsfiler,
  Ords WordFiles och gärna filer på disken) så att inget som Office skrivit avvisas.
- **Fientliga filer blir tester** i `HostileFilesTests.cs`.
- Version 1.0.0 släpps samma dag som Avig.Kalkyl 1.0.0. Till dess 0.x i `G:\LocalNugetRepo`.

## Konventioner

Som i Avig.Kalkyl: explicita typer (aldrig `var`), filomfattande namnrymder, klammrar alltid,
kommentarer på engelska, CRLF, Avig.Analyzers på `error`.
