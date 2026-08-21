# Lovion Coordinate Tool

Deze tool verrijkt een Lovion Excel-lijst met coordinaten en schrijft vier outputs weg:

* `xlsx` met extra kolommen
* `csv` met extra kolommen
* `geojson` in `EPSG:28992`
* `shapefile` (`.shp/.shx/.dbf/.prj/.cpg`)

Hij kan nu twee invoertypes verwerken:

* `Excel`: Lovion-lijst met adressen
* `TXT`: Lovion-tekstexport per LS-aansluiting, inclusief `Overdrachtspunt`

Daarnaast is er nu een GUI-workbench:

* meerdere bestanden tegelijk laden
* geplakte tekst toevoegen via een knop of direct vanuit het klembord importeren, met eigen `SourceFile`-naam
* meerdere stations achter elkaar openen en alle LS-aansluitingen per station via de knop `Stations` importeren
* adressen in de grid direct geocoderen via `PDOK`
* rijen daarna direct snappen naar de officiële Enexis `WFS`
* street-furniture termen beheren via een eigen termen-menu
* alle aansluitingen bewerken in een grid
* exporteren naar `csv`, `geojson` en `shapefile`

## Werking

1. Leest het opgegeven Excel-werkblad in.
2. Bouwt per rij een adresquery op basis van `Straatnaam`, `Huisnummer`, `Postcode` en `Woonplaats`.
3. Haalt coordinaten op via PDOK Locatieserver.
4. Voor normale entries accepteert hij PDOK-hits die exact overeenkomen op straat, huisnummer en postcode, plus tolerant `house_suffix_mismatch` zoals `5` versus `5A`.
5. Voor `Gebruiksdoel = Straatmeubilair` gebruikt hij het adres alleen als zoekanker en probeert hij daarna een WFS-aansluiting op dezelfde straat of in de buurt te kiezen.
6. Probeert coordinaten optioneel te verfijnen met Enexis WFS-features uit `Opendata:asm_e_lv_service_connection`.
7. Schrijft verrijkte outputbestanden weg.

Voor tekstexporten:

1. Leest per `LS Aansluiting ...` blok alle key/value-regels in.
2. Parseert `Overdrachtspunt` als RD-coordinaat uit `linksonder` en `rechtsboven`.
3. Gebruikt de puntlocatie van `Overdrachtspunt` als shapefile-geometry.
4. Zet alle velden uit de tekstexport als attributen in `csv`, `geojson` en `shapefile`.
5. Schrijft een extra `fieldmap.csv` weg, omdat shapefile-attribuutnamen maximaal 10 tekens lang mogen zijn.

## Gebruik

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Get-LovionCoordinates.ps1 `
  -InputExcelPath "C:\pad\naar\Alle stations nieuw.xlsx" `
  -WorksheetName "Alle stations nieuw"
```

Met een lokale GeoJSON in plaats van live WFS:

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Get-LovionCoordinates.ps1 `
  -InputExcelPath "C:\pad\naar\Alle stations nieuw.xlsx" `
  -WorksheetName "Alle stations nieuw" `
  -WfsGeoJsonPath "C:\pad\naar\asm_e_lv_service_connection.geojson"
```

Alleen adressen geocoden, zonder WFS:

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Get-LovionCoordinates.ps1 `
  -InputExcelPath "C:\pad\naar\Alle stations nieuw.xlsx" `
  -WorksheetName "Alle stations nieuw" `
  -SkipWfs
```

Kleine proefrun op bijvoorbeeld 25 rijen:

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Get-LovionCoordinates.ps1 `
  -InputExcelPath "C:\pad\naar\Alle stations nieuw.xlsx" `
  -WorksheetName "Alle stations nieuw" `
  -SkipWfs `
  -MaxRows 25
```

Tekstexport direct naar shapefile:

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Get-LovionCoordinates.ps1 `
  -InputTextPath "C:\pad\naar\tekstfile van de excel gelimiteerde tot 100 LS aansluitingen.txt"
```

## EXE Bouwen

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Build-LovionCoordinateToolExe.ps1
```

Daarna staat de launcher hier:

* `.\dist\LovionCoordinateTool.exe`

Voorbeeld:

```powershell
.\dist\LovionCoordinateTool.exe `
  -InputExcelPath "C:\pad\naar\Alle stations nieuw.xlsx" `
  -WorksheetName "Alle stations nieuw" `
  -WfsGeoJsonPath "C:\pad\naar\asm_e_lv_service_connection.geojson"
```

Voor tekstexport:

```powershell
.\dist\LovionCoordinateTool.exe `
  -InputTextPath "C:\pad\naar\tekstfile van de excel gelimiteerde tot 100 LS aansluitingen.txt"
```

## C++ CLI

Er staat ook een native C++ CLI-port in `.\cpp`.

Build:

```powershell
cmake -S .\cpp -B .\cpp\build
cmake --build .\cpp\build --config Release
```

Daarna staat de C++ exe hier:

* `.\cpp\build\Release\LovionCoordinateToolCpp.exe` (Visual Studio-generator)

De C++ versie ondersteunt nu `txt` en `csv` invoer, PDOK, WFS/lokale GeoJSON en export naar `csv`, `geojson` en `shapefile`. Directe `.xlsx`-invoer en de GUI blijven voorlopig in de PowerShell-versie; sla Excel eerst op als CSV als je de C++ CLI wilt gebruiken.

Voorbeeld:

```powershell
.\dist_cpp\LovionCoordinateToolCpp.exe `
  -InputCsvPath "C:\pad\naar\lovion.csv" `
  -SkipWfs
```

## Lovion-batchautomatisering

De knop `Lovion 100-batches` in de PowerShell-GUI automatiseert tabelkopieen uit Lovion/Citrix:

* start met de eerste geselecteerde rij in Lovion
* leest `Aantal geladen`, `Gefilterd` en `Geselecteerd` uit de Lovion-weergave
* probeert per batch exact `100` rijen te kopieren via `Exporteren > Kopieer naar klembord`
* corrigeert automatisch met `Shift + Pijl omhoog/omlaag` als er te veel of te weinig rijen in het klembord terechtkomen
* importeert elke batch direct terug in de workbench
* gebruikt bij een actieve filter het zichtbare aantal `Gefilterd`; anders wordt `Aantal geladen` gebruikt
* kan optioneel stoppen op een handmatig ingevuld totaal, of automatisch op het uit Lovion gelezen totaal
* laat de cursor op de laatst gebruikte Lovion-positie staan en stopt de automatisering zodra de gebruiker de muis zelf beweegt

## GUI

Script starten:

```powershell
PowerShell -ExecutionPolicy Bypass -File .\LovionCoordinateTool.Gui.ps1
```

GUI EXE bouwen:

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Build-LovionCoordinateToolGuiExe.ps1
```

Daarna staat de GUI-launcher hier:

* `.\dist\LovionCoordinateWorkbench.exe`

Wat de GUI kan:

* meerdere `txt`, `xlsx` en `csv` bestanden tegelijk laden
* eerste kolom is altijd `SourceFile`
* `Klembord Import` om de huidige clipboardtekst direct toe te voegen, met een eigen `SourceFile`-naam
* `Plakvenster` voor grote Lovion-tekstexports die je eerst wilt bekijken of handmatig wilt aanpassen, ook met een eigen `SourceFile`-naam
* `Stations` vraagt een lijst stationnummers, start vanuit de Lovion-lijst `LS stroomtransformatorgroep`, opent voor iedere query eerst de bovenste tab `EXPLORE`, klikt daarna op `Start query`, opent het resultaat via `Open > VIEW` en importeert alle aansluitingen met de bestaande 100-regelmethode; `SourceFile` is exact het stationnummer, zonder batchnummer erachter
* `Geocodeer PDOK` om rijen zonder coordinaten op adres te verrijken
* `Geocodeer PDOK` gebruikt nu standaard parallelle requests voor unieke adresqueries, waardoor grote lijsten merkbaar sneller lopen
* `Geocodeer PDOK` gebruikt nu een veldgerichte query op `straatnaam`, `huisnummer`, `postcode` en `woonplaatsnaam`, gefilterd op `type:adres` en `bron:BAG`
* als die strakke PDOK-query niets oplevert, probeert de workbench automatisch eenvoudigere fallback-queries zonder toevoeging of als losse adresregel
* `Snap Enexis` om alleen specifieke `Gebruiksdoelen` naar `street_furniture` te zetten en de rest naar `service_connection`
* `Snap Enexis` gebruikt nu unieke Enexis-punten; een punt of feature wordt dus maar aan één rij toegewezen
* `Bouwaansluiting` wordt altijd als `service_connection` behandeld en krijgt bij hetzelfde adres pas na de andere aansluitingen een vrije `service_connection`
* `Snap Enexis` gebruikt nu ook een ruimtelijke index, zodat grote lijsten veel sneller snappen dan bij een volledige scan over alle features per rij
* `Street Terms` of `Instellingen > Street Terms` om de woordenlijst voor `street_furniture` te beheren
* bewerkbare grid voor alle aansluitingen
* `Herlees Coords` om `Overdrachtspunt` of `FinalRdX/FinalRdY` opnieuw te verwerken
* export van de huidige grid naar `csv`, `geojson` en `shapefile`
* shapefile gebruikt voor tekstexports expliciet `OverdrachtspuntRdX1` en `OverdrachtspuntRdY1`
* shapefile-export schrijft ook een `*_fieldmap.csv` weg

De GUI-build embedt het PowerShell-script in de exe. Op een Windows-machine met PowerShell start je dus in de praktijk met alleen `LovionCoordinateWorkbench.exe`.

## Belangrijk

* `CoordinateSource = pdok` betekent: coordinaat komt uit een strikte PDOK-match op het adres.
* `CoordinateSource = pdok_tolerant` betekent: coordinaat komt uit een getolereerde PDOK-match met `PdokMatchReason = house_suffix_mismatch`.
* `CoordinateSource = wfs` betekent: coordinaat is gesnapt naar een WFS-feature.
* `CoordinateSource = overdrachtspunt_text` betekent: punt komt direct uit het `Overdrachtspunt` in de tekstexport.
* `PdokStrictMatch = False` met `PdokMatchReason = house_suffix_mismatch` is geen perfecte match, maar wordt wel als tolerante adresmatch gebruikt.
* `service_connection_point_already_used_keep_seed` of `street_furniture_point_already_used_keep_seed` betekent: er was wel een kandidaat, maar die lag al vast op een eerdere rij; de rij houdt dan zijn seed-coordinaat.
* In de GUI geldt nu: alleen `Gebruiksdoelen` die matchen op de ingestelde street-furniture termen zoals `riool`, `camera`, `verlichting`, `container` gaan naar `Enexis_Opendata:asm_e_lv_street_furniture`.
* Alle andere `Gebruiksdoelen` gaan naar `Enexis_Opendata:asm_e_lv_service_connection`.
* Die termen zijn in de GUI aanpasbaar en worden opgeslagen in `street_furniture_terms.json` naast de exe of naast het script.
* Matching op die termen is niet hoofdlettergevoelig en mag overal in `Gebruiksdoel` voorkomen.
* `Bouwaansluiting` is hiervan uitgezonderd en blijft altijd `service_connection`, ook als iemand die term in de termenlijst probeert te zetten.
* Voor `street_furniture` matcht de tool eerst op `omschrijving ~= gebruiksdoel` en daarna op afstand.
* Als meerdere aansluitingen op hetzelfde adres zitten, wordt `Bouwaansluiting` pas na de andere aansluitingen op dat adres gekoppeld aan de eerstvolgende vrije `service_connection`.
* PDOK-geocoding draait standaard met meerdere gelijktijdige unieke adresqueries in plaats van volledig serieel.
* PDOK zoekt nu niet meer alleen op losse vrije tekst, maar met een expliciete adresquery zoals `straatnaam:"..." and huisnummer:... and postcode:... and woonplaatsnaam:"..."`, plus `fq=type:adres` en `fq=bron:BAG`.
* als zo'n strakke query niets vindt, probeert de workbench automatisch een eenvoudiger PDOK-adres zonder toevoeging en daarna een losse tekstquery, zodat lege coordinaatvelden zoveel mogelijk worden voorkomen.
* Enexis-snapping gebruikt een ruimtelijke bucket-index rond de seed-coordinaten in plaats van per rij alle WFS-features opnieuw te sorteren.
* Controleer `UsageCategory`, `WfsLayer`, `WfsMatchMode` en `WfsDistanceM` voordat je de uitkomst als definitieve netlocatie gebruikt.

## Outputkolommen

* `UsageCategory`, `StreetKey`
* `PdokRdX`, `PdokRdY`, `PdokLon`, `PdokLat`
* `PdokStrictMatch`, `PdokMatchReason`
* `WfsFeatureId`, `WfsMatchMode`, `WfsDistanceM`
* `FinalRdX`, `FinalRdY`, `FinalLon`, `FinalLat`
* `CoordinateSource`, `MatchStatus`
* Bij tekstexport ook: `RecordTitle`, `ExportedBy`, `ExportedAt`, `OverdrachtspuntRdX1`, `OverdrachtspuntRdY1`, `OverdrachtspuntRdX2`, `OverdrachtspuntRdY2`, `OverdrachtspuntIsPoint`
