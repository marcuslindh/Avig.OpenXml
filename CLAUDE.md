# CLAUDE.md

Knowledge-projekt: avigord
Dokumentation, beslut och agentlogg finns i Knowledge — se ~/.claude/CLAUDE.md. Paketet hör
till Avig.Ord-projektet, eftersom det bröts ut ur Avig.Kalkyl för Avig.Ords skull (beslut
"Avig.Ord — Beslut: Gemensamt paket Avig.OpenXml för koden som delas med Avig.Kalkyl").

## Vad det här är

NuGet-paketet Avig.OpenXml: paketlagret (OPC) under Avig.Kalkyl, Avig.Ord och Avig.Bildspel. Läser
och skriver zip-paket och Flat OPC: delar, content types och relationer, med gränser mot
fientliga filer. Det vet inget om kalkylblad, dokument eller presentationer.

Från 0.2.0 finns också det som de tre delar:

- **`Length`** (namnrymden `Avig.OpenXml`): mått lagrade som EMU, flyttad hit från Avig.Ord.
- **Namnrymder och content types** i `OpcNamespaces`, också DrawingML (a:, c:, asvg:) och
  PresentationML (presentation, bilder, layouter, mallar, anteckningar, båda sorternas
  kommentarer, presProps, viewProps, tableStyles och .pptm, .potx, .potm, .ppsx, .ppsm).
- **DrawingML-lagret** (namnrymden `Avig.OpenXml.Drawing`), flyttat hit från Avig.Kalkyl (beslut
  i Knowledge: "Avig.Bildspel — Beslut: DrawingML och Length flyttas till Avig.OpenXml"):
  former och förinställd geometri (`Shape`, `ShapeKind`, `Shape.FromPreset`, `ShapeAdjustment`),
  fyllningar (`Fill`: färg, gradient, mönster, bild), linjer (`Outline`), skuggor och mjuka kanter
  (`Shadow`), text i former (`ShapeText`, `TextRun`, `TextFont`, `TextAlignment`, `TextAnchor`),
  grupper (`ShapeGroup`, `GroupItem`, också nästlade), bilder (`Picture`, PNG/JPEG/GIF/BMP/TIFF och SVG med
  rasterbild som reserv, storleken ur filhuvudet), länkar (`Hyperlink`), platser i `Length`
  (`DrawingBounds`) och diagramdelen (`ChartSpace`, `ChartSpaceSeries`, `ChartData`,
  `ChartAxis`, `DataLabels` …). `DrawingWriter` och `ChartWriter` skriver, `DrawingReader` och
  `ChartReader` läser tillbaka (på lokalt namn, värden utanför typernas gränser klämms).

  **Lagret vet inte var det ligger.** Formträdets namnrymd (`xdr:` i ett blad, `p:` på en bild) och
  dess två skillnader (Excels tomma `macro`, bildens `nvPr`) ges av värden i `ShapeTreeNamespace`.
  Ankare, `p:spTree`, delarna och relationerna skriver värden; `DrawingContext` samlar id:n och
  vad relationerna ska peka på (`Targets`), och `DrawingReadContext` ger läsaren relationerna och
  bilderna. Diagrammens formler (`ChartData.Formula`) skrivs och läses som de är; vad de pekar på
  bestämmer värden (Kalkyls `Chart`/`ChartSeries` gör om cellområden till formler).

  Från 0.2.1 (för Avig.Bildspels läsare): `DrawingColors` löser temafärger och färgmodifierare,
  `DrawingReadContext.SchemeColors` låter läsaren använda dem (null, som förut, lämnar dem utanför),
  `DrawingReader.Read` läser utan I/O med en `DrawingReadContext` som har bilderna i minnet, och
  `Hyperlink.Read` och `DrawingBounds.Read` (negativa positioner) är publika för läsare.

  Från 0.2.3 (för Avig.Bildspels skrivare): skrivaren är synkron (`DrawingWriter.WriteShape` …,
  `ChartWriter.WriteChartSpace`) och skriver i vilken `XmlWriter` som helst, också en utan `Async`
  eller en i ett `XElement`. `…Async`-formerna finns kvar, skriver samma XML i skrivarens buffert
  och ger en färdig uppgift.

  Från 0.3.0 (för Avig.Bildspels milstolpe 6, formerna): egna geometrier (`CustomGeometry`,
  `Shape.FromGeometry`), kopplingar som `p:cxnSp` med limmade ändar (`Shape.IsConnector`,
  `ConnectionSite`), alla 187 förinställda former (`Shape.PresetNames`), övertoningar efter bana,
  rutade bildfyllningar, gruppens och bakgrundens fyllning (`Fill.PathGradient`, `TiledPicture`,
  `Group`, `Background`), linjer med sammansättning, ändar, hörn, egen streckning, pilspetsar och
  övertoning (`Outline`), sken, reflektion, oskärpa, förinställd skugga, överlägg och perspektivskugga
  (`GlowEffect`, `ReflectionEffect` …), 3D (`Scene3D`, `Shape3D`, `Bevel`, alla kameror, ljus och
  material), temats formformat (`ShapeStyle`), låsningar, dold, dekorativ och titel. Det som läggs
  till skrivs bara när det är satt, så att Kalkyls XML är densamma byte för byte.

  Också från 0.3.0 (för Avig.Bildspels diagram, milstolpe 6): alla diagramtyper PowerPoint 365 gör.
  `ChartSpace` har fått `Grouping` (`ChartGrouping`, `Stacked` är samma inställning),
  `ThreeD` (`ChartView3D`), `BarShape`, `RadarStyle`, `Wireframe`, `ScatterLines`, bubbel-,
  börs- och cirkel-av-cirkel-inställningar, `SeriesAxis`, `DataTable`, `Style` och
  `ExternalData` (relationen till arbetsboken); `ChartType` har fått `Radar`, `Bubble`,
  `Stock`, `Surface`, `PieOfPie` och `BarOfPie`; serierna `BubbleSizes`, `Explosion`,
  `Trendlines` och `ErrorBars`; axlarna `MinorGridlines`, `Reversed`, `LogBase`, `MinorUnit`.
  Den nya diagrammotorn (`cx:`) är `ChartExSpace` med `ChartExWriter`/`ChartExReader`, och
  `ChartWorkbook` gör den minsta .xlsx ett diagram behåller sina värden i. Det som fanns skrivs
  som förut, utom att ett punktdiagram där alla serier är utjämnade får
  `scatterStyle="smoothMarker"` (annars läser PowerPoint och Excel typen som raka linjer).
  `cx:axisId`, `cx:binSize` och `cx:binCount` skrivs med `val`-attribut, som PowerPoint 365
  kräver; Open XML SDK:s schema har texten som innehåll, så de fynden filtreras i testerna.

  Från 0.3.1 (för Avig.Bildspels milstolpe 6, länkarnas ljud): `Hyperlink.HighlightClick`,
  `StopsSound` och `Sound` (`HyperlinkSound`): `highlightClick` och `endSnd` skrivs och läses på
  `a:hlinkClick`; ljudet (`a:snd`) lagrar och skriver ägaren. Inget skrivs när de inte är satta.

  En ändring i lagret ska klara `DrawingTests`/`ChartTests` (schemavalidering med
  DocumentFormat.OpenXml i både `xdr:` och `p:`, och rundtur: skriv, läs, skriv samma XML), och
  Kalkyls tester, ExcelCheck och samples (samma XML byte för byte), innan den släpps.

## Kommandon

```powershell
dotnet build Avig.OpenXml.slnx
dotnet test Avig.OpenXml.slnx
./publish.ps1 -NoBump          # testar, packar, formkontrollerar, lägger i G:\LocalNugetRepo
```

## Regler

- **Inga beroenden utom Avig.Result.** `publish.ps1` failar annars. Testerna får använda
  DocumentFormat.OpenXml, för schemavalideringen av DrawingML och för att jämföra PresentationML-
  namnen med Open XML SDK:s.
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
