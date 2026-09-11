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
* Lovion-macro-informatie tonen en het bestand `Trafo zoeken macro rapport.xml` downloaden voor import via Lovion `Rapporten importeren`
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

De knop `Stations` gebruikt voor de navigatie nu OCR op het actuele Lovion-venster. De tabs, `Start query`, `VIEW`, `Schakelaar`, `Aansluitingen` en de bijbehorende `Netwerk`-opties worden pas aangeklikt nadat het zichtbare label in de juiste schermregio is gevonden. Hierdoor wordt een verschoven of nog niet volledig geladen menu niet meer blind op een vaste positie aangeklikt. De workflow wacht langer op trage Lovion-schermen en stopt met een duidelijke foutmelding als een vereist label niet verschijnt. Voor het trafo-aantal is `Aantal geladen` leidend: 1 betekent één trafo en 2 of meer betekent dat elke geladen rij afzonderlijk wordt verwerkt. De trafo-rij wordt via OCR verankerd aan de kolomkop `Master Asset ID` in de resultaatlijst rechtsonder, zodat de klik niet in de navigatieboom links kan belanden. Als de eerste stationquery meer dan één resultaat oplevert, worden de resultaten één voor één verwerkt: iedere trafo krijgt een eigen selectie-, `Open`-, netwerk- en exportcyclus, terwijl alle geëxporteerde rijen dezelfde stationcode als `SourceFile` houden. Als de kaart geen centrale marker heeft maar wel twee gele resultaten toont, wordt het station ook bewust overgeslagen en aan het eind gemeld voor handmatige toevoeging aan de Lovion Coordinate Workbench-lijst.

Voor grote aansluitingenlijsten zet de batch de grid eerst met `Ctrl+Home` op de eerste rij, houdt één fysieke `Shift` continu vast tijdens de reeks `↓`-toetsen en controleert zowel de teller als de klembordexport. Na een batch wordt `Shift` losgelaten en volgt precies één gewone `↓` voor de volgende startregel.

## Stationsworkflow in Lovion/Citrix

De knop `Stations` is bedoeld voor de native Lovion-applicatie die via Citrix op het bureaublad draait. Lovion moet volledig zichtbaar zijn op het meest linkse van drie schermen; het is geen browserworkflow. Tijdens de automatisering mag de gebruiker de muis niet bewegen en niet in Lovion typen. De muisbewaking stopt het proces bewust wanneer handmatige beweging wordt gedetecteerd.

Startvoorwaarden:

1. Open in Lovion de lijst `LS stroomtransformatorgroep`.
2. Laat Lovion zichtbaar op het linker scherm staan.
3. Geef de stationslijst op in de `Stations`-dialoog, bijvoorbeeld:

   ```text
   017.584, 017.525, 017.523, 017.503, 017.519, 017.545,
   017.562, 017.571, 017.577, 017.512, 017.526, 017.568
   ```

Per station werkt de workflow als volgt:

1. Klikt de stationslijst aan, selecteert de bestaande stationwaarde met `Ctrl+A` en typt het stationnummer.
2. Klikt alleen op gevonden OCR-labels in de actuele schermregio. De workflow gebruikt dus niet blind één vaste knoppositie.
3. Klikt achtereenvolgens op `EXPLORE`, `Start query`, `Open` en `VIEW`. `Start query` wordt maximaal 45 seconden gezocht; als Citrix traag is, wordt gewacht totdat het zichtbare label en de verwachte volgende schermtoestand aanwezig zijn.
4. Na `Start query` leest de workflow `Aantal geladen` en `Gefilterd`. De OCR moet eerst van het vorige scherm veranderen en daarna stabiel zijn. `Aantal geladen` is de leidende teller; `Gefilterd` wordt alleen gebruikt als Lovion de geladen-teller niet toont.
5. De teller van de eerste stationquery bepaalt hoeveel trafos bij het station horen. Bij exact één resultaat wordt de rij niet aangeklikt: Lovion gaat direct door naar `Open` en `VIEW`. Bij meerdere resultaten wordt via OCR de kolomkop `Master Asset ID` in de rechter resultaatlijst gezocht en wordt daar per cyclus precies één rij geselecteerd. Daarna leest de workflow `Geselecteerd` met een volledige schermscan en wacht op stabiele tellerwaarden voordat `Open` en `VIEW` worden uitgevoerd. Daardoor wordt een geslaagde klik niet door een te kleine OCR-crop als mislukte selectie gezien en wordt de klik niet verward met de boomregel `LS Stroomtransformatorgroep` links. Na het sluiten van de resultaat-tabs keert Lovion terug naar de al geladen eerste resultatenlijst van hetzelfde station; voor trafo 2 en volgende wordt daar direct de volgende rij aangeklikt, zonder het station opnieuw te zoeken of `Start query` opnieuw uit te voeren. Zo blijven meerdere trafos aan hetzelfde station gekoppeld zonder dat selecties of netwerkresultaten door elkaar lopen.
6. Voor de geselecteerde trafo wordt `VIEW` geopend. Daarna worden in de netwerkweergave de gevonden labels gebruikt voor `Netwerk`, `Schakelaars`, `Rekening houden met standen van schakelaars`, `Geen objecten` en `Aansluitingen`.
7. Na de kaart-/netwerkselectie wordt `Exporteren` via OCR gevonden en wordt gewacht op de LS-aansluitingenlijst. De tabelteller wordt pas geaccepteerd als de kolomkoppen en de aantallen betrouwbaar zichtbaar zijn.

Meerdere eerste zoekresultaten zijn geen oversla-conditie meer. Tijdens het testen was dat het geval voor `017.545`: Lovion gaf `geladen=2` en `gefilterd=2`. De workflow gebruikt dit aantal nu om trafo 1 en trafo 2 afzonderlijk te verwerken. Beide importbatches gebruiken exact `017.545` als `SourceFile`; in de samenvatting worden ze herkenbaar weergegeven als `017.545 trafo 1/2` en `017.545 trafo 2/2`.

### Selecteren en kopiëren van aansluitingen

Voor een aansluitingenlijst met meer dan 100 rijen:

1. De grid krijgt focus en wordt met `Ctrl+Home` teruggezet naar de eerste rij. Dit voorkomt dat een eerder horizontaal of verticaal gescrolde grid maar de zichtbare 14 rijen selecteert.
2. De eerste rij wordt enkel geselecteerd.
3. Voor iedere volgende rij wordt dezelfde fysieke `Shift` ingedrukt gehouden. Terwijl Shift continu ingedrukt blijft, wordt `↓` één voor één verstuurd. Er zit een pauze tussen de toetsen omdat Citrix toetsen vertraagd kan verwerken.
4. De teller `Geselecteerd` wordt via OCR gelezen. De workflow wacht op twee stabiele metingen voordat de selectie verdergaat.
5. De selectie wordt via `Exporteren > Kopieer naar klembord` gekopieerd. De tekstexport wordt meerdere keren gecontroleerd totdat het exacte verwachte aantal herkenbare rijen aanwezig is.
6. Na de eerste 100 rijen wordt de laatste geselecteerde rij opnieuw gefocust, wordt Shift losgelaten en wordt precies één gewone `↓` verstuurd. De teller moet dan 1 zijn; dat is de start van de volgende batch.

Shift mag niet tijdens het muiswiel-scrollen worden vastgehouden: in Lovion veroorzaakt dat horizontaal scrollen. In de huidige batchworkflow is de pijltjesmethode daarom de primaire methode. Bij een trage verwerking wordt gewacht en niet direct op een tussentijdse OCR-teller gecorrigeerd.

### Foutafhandeling en logging

Alle belangrijke stappen worden naar `LovionBatch.log` geschreven, waaronder OCR-klikken, wachttijden, eerste zoekresultaten, trafo-indexen, batchaantallen, klembordcontroles en handmatige kaartgevallen. Als een vereist label niet binnen de timeout verschijnt, wordt de workflow gestopt met een foutmelding in plaats van op een vermoedelijke positie door te klikken.

De bestaande kaartmarkerfallback kan na een mislukte centrale kaartklik nog gele markers proberen en heeft een aparte handmatige-melding voor ambigue kaartresultaten. Die situatie staat los van de trafo-verwerking: meerdere resultaten in de eerste stationquery worden nu wel afzonderlijk geopend, terwijl ambigue kaartmarkers nog handmatige toevoeging aan de Workbench kunnen vereisen.

## Lovion-macro

De knop `Lovion macro` toont de installatie-informatie voor de macro `Trafo zoeker Macro` en biedt `XML downloaden`. De knop slaat de meegeleverde macro op als `Trafo zoeken macro rapport.xml`; de XML zit bij een standalone EXE ingebouwd, zodat de knop ook werkt zonder een los bronbestand naast de EXE.

Installeer de macro in Lovion/Citrix als volgt:

1. Klik in de Workbench op `Lovion macro` en daarna op `XML downloaden`.
2. Open in Lovion de functie `Rapporten importeren`.
3. Selecteer het opgeslagen XML-bestand en importeer/installeer de macro.
4. Controleer in Lovion onder `GEN > Elektriciteit > LS stroomtransformatorgroep` of `Trafo zoeker Macro` zichtbaar is.
5. Gebruik de parameter `Station` om een stationnummer in te vullen.

De macro vereist Lovion release `7.2.1` of nieuwer. De daadwerkelijke import vindt plaats in de native Lovion/Citrix-app; de Workbench levert het gevalideerde XML-bestand en de instructies.

### Getest gedrag

De basisroute is live op de Citrix-Lovion-app getest met de opgegeven stations. In de brede test werden 1.735 aansluitingen verwerkt; onder andere `017.525` als `100 + 21`, `017.519` als `100 + 100 + 22` en `017.577` als `100 + 100 + 75`. Daarbij leverde `017.545` twee eerste zoekresultaten op (`geladen=2`, `gefilterd=2`); de oude skipcontrole is met deze wijziging vervangen door twee afzonderlijke trafo-cycli. De nieuwe multi-trafo-lus is met een gecontroleerde PowerShell-mocktest gevalideerd: twee `Open`-cycli, twee imports met dezelfde `SourceFile` `017.545`, 84 mock-aansluitingen totaal en twee sluitcycli.

Daarnaast zijn de PowerShell-smoketest, `git diff --check`, de standalone GUI-EXE-build en een EXE-smoketest uitgevoerd. De EXE-build eindigde met exitcode 0. Aparte regressietests bevestigen dat bij `Aantal geladen=1` geen trafo-rijselectie of geselecteerd-tellercontrole wordt uitgevoerd, terwijl bij `Aantal geladen=2` wel precies één rijselectie en tellercontrole plaatsvinden, en dat trafo 2 de al geladen stationresultaten hergebruikt zonder opnieuw zoeken of `Start query`.

## Onderhoud en toekomstige aanpassingen

### Belangrijke bestanden

| Bestand | Rol |
| --- | --- |
| `LovionCoordinateTool.Gui.ps1` | Hoofdbronbestand van de GUI, OCR, muis-/toetsenbordinvoer en stationsworkflow. |
| `assets/Trafo zoeken macro rapport.xml` | Bronbestand van de Lovion-macro `Trafo zoeker Macro`; wordt bij de GUI-build in de EXE ingebouwd. |
| `Get-LovionCoordinates.ps1` | Losse CLI voor adresverrijking en export; dit bestand wordt niet gebruikt door de knop `Stations`. |
| `Build-LovionCoordinateToolGuiExe.ps1` | Bouwt de standalone `LovionCoordinateWorkbench.exe` en embedt de GUI-broncode. |
| `Build-LovionCoordinateToolExe.ps1` | Bouwt de oudere CLI-EXE voor de niet-GUI workflow. |
| `tests/Test-LovionStationWorkflow.ps1` | Lokale regressietests voor OCR-ankering, tellers en hergebruik van stationresultaten. |
| `.github/workflows/release-v1.2.0.yml` | Bouwt en publiceert de standalone release-EXE wanneer de release-tag wordt gepusht; de workflownaam/tagversie wordt bij iedere release bijgewerkt. |
| `LovionBatch.log` | Runtime-log naast de GUI of EXE; bevat OCR-resultaten, trafo-indexen, wachttijden en batchstatussen. |

### Stationsworkflow als toestandsmachine

De workflow bestaat uit twee lussen. De buitenste lus verwerkt de opgegeven stations in volgorde. De binnenste lus verwerkt alle trafo-resultaten van één station. De belangrijkste regel is dat alleen trafo-index `0` een stationzoekactie uitvoert.

1. `Start-LovionStationsImport` leest de stationlijst, vraagt bevestiging, minimaliseert de Workbench en geeft vijf seconden om Lovion/Citrix actief te laten worden.
2. `Invoke-LovionStationsImport` start de muisbewaking en ruimt modifiers op in een `finally`-blok.
3. `Invoke-LovionStationsImportCore` maakt per station één exact `SourceFile`-prefix en zet `TransformerIndex=0` en `TransformerCount=1`.
4. `Open-LovionConnectionsForStation` controleert of `LS Stroomtransformatorgroep` zichtbaar is. Bij de eerste trafo wordt het stationnummer ingevoerd, daarna worden `EXPLORE` en `Start query` via OCR aangeklikt.
5. `Wait-LovionInitialStationQueryResults` wacht totdat de post-query OCR veranderd en stabiel is. `Aantal geladen` bepaalt het trafo-aantal; `Gefilterd` is alleen fallback als `Aantal geladen` niet beschikbaar is.
6. Bij exact één geladen resultaat wordt geen rij aangeklikt. Lovion gaat direct naar `Open` en `VIEW`.
7. Bij meerdere resultaten zoekt `Wait-LovionTransformerResultRowPoint` de OCR-kop `Master Asset ID` in de rechterondertabel. Daaruit wordt de rijpositie met de rij-index en een pitch van 22 pixels berekend. Daarna wordt één rij aangeklikt en wordt `Geselecteerd=1` met een volledige schermscan gecontroleerd.
8. De netwerkstappen klikken `Netwerk`, `Schakelaar`, de schakelaarinstelling, `Geen objecten` en `Aansluitingen` via OCR. Vervolgens wordt het eerste netwerkresultaat geopend, wordt `Exporteren` gevonden en wordt gewacht op de aansluitingenlijst.
9. Na het importeren sluit `Close-LovionStationResultTabs` de resultaat-tabs en controleert hij de terugkeer naar de eerste stationresultatenlijst.
10. Als er nog een trafo-index over is, geeft `Invoke-LovionStationsImportCore` `ReuseCurrentStationResults` en het eerder gelezen `KnownTransformerCount` mee. De tweede trafo wordt dan direct in dezelfde geladen lijst aangeklikt; stationnummer, `EXPLORE` en `Start query` worden niet opnieuw uitgevoerd.
11. Pas na de laatste trafo gaat de buitenste lus verder naar het volgende station en wordt een nieuwe stationquery uitgevoerd.

Wijzig nooit de algemene `Invoke-LovionClickRow`-coördinaten om de trafo-resultaten te repareren. Die functie is voor de aansluitingenlijst en kan op een ander grid werken. Trafo-resultaten moeten via `Wait-LovionTransformerResultRowPoint` aan `Master Asset ID` worden geankerd. Alle opgenomen Lovion-punten zijn relatief aan het Lovion-venster, niet aan een willekeurig scherm.

### Onderhoudsregels

* Voeg voor nieuwe Lovion-knoppen eerst OCR met een schermregio en een timeout toe; gebruik een vaste positie alleen als expliciete guarded fallback.
* Houd de volledige schermscan aan voor tellers onderin rechts, vooral `Geselecteerd`. De bovenste rechter crop is daarvoor niet voldoende.
* Laat bij één trafo de rijselectie weg. Een extra klik kan in Citrix de navigatieboom selecteren.
* Laat bij trafo 2 en volgende de bestaande resultatenlijst staan. Voeg geen tweede `Start query` toe zolang `Close-LovionStationResultTabs` naar dezelfde lijst terugkeert.
* Houd `SourceFile` exact gelijk aan het stationnummer. Gebruik trafo-indexen alleen in statusmeldingen en logregels.
* Behoud de muisguard en modifier-opruiming. Handmatige muisbeweging moet de automatisering veilig stoppen.
* Pas timeouts pas aan nadat de log heeft aangetoond welke schermtoestand traag is. De huidige richtwaarden zijn: `Start query` 45 s, initiële resultaten 45 s, trafo-kop 35 s, geselecteerde teller 20 s, netwerklabels 30–45 s en aansluitingenlijst 60 s.

### Lokaal testen vóór een release

Voer vanuit de repository uit:

```powershell
PowerShell -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-LovionStationWorkflow.ps1
PowerShell -NoProfile -Sta -ExecutionPolicy Bypass -File .\tests\Test-LovionMacro.ps1
PowerShell -NoProfile -Sta -ExecutionPolicy Bypass -File .\tests\Test-GuiSurface.ps1
PowerShell -NoProfile -Sta -ExecutionPolicy Bypass -File .\LovionCoordinateTool.Gui.ps1 -SmokeTest
PowerShell -NoProfile -ExecutionPolicy Bypass -File .\Build-LovionCoordinateToolGuiExe.ps1 -OutputDirectory .\dist-test
.\dist-test\LovionCoordinateWorkbench.exe -SmokeTest
.\dist-test\LovionCoordinateWorkbench.exe -MacroSmokeTest
git diff --check
```

Voor een handmatige Citrix-test moet Lovion volledig zichtbaar op het meest linkse scherm staan. Test minimaal `017.584` (één resultaat: geen rijselectie) en `017.545` (twee resultaten: trafo 1 openen, terug naar dezelfde eerste lijst, trafo 2 direct selecteren). Beweeg tijdens deze test de muis niet en controleer `LovionBatch.log` na afloop.

### Releaseproces

1. Werk eerst code, README en regressietests bij op `main`.
2. Voer de lokale tests hierboven uit en controleer de diff.
3. Verhoog in `.github/workflows/release-v1.2.0.yml` de workflownaam, tag en release-notities naar de nieuwe versie.
4. Commit en push `main`.
5. Push exact dezelfde versie-tag, bijvoorbeeld `v1.2.0`. GitHub Actions bouwt dan de standalone EXE, voert de EXE-smoketest uit en publiceert alleen `LovionCoordinateWorkbench.exe`.
6. Controleer na afloop de GitHub-release, de assetnaam en de release-notities. Een release mag niet verwijzen naar een oudere tag of een oude EXE.

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

### Verschil tussen `Klembord` en `Plakken`

* `Klembord` leest de huidige tekst op het Windows-klembord direct uit en importeert die na het opgeven van een `SourceFile`-naam. Dit is de snelle route wanneer de Lovion-export al klaarstaat.
* `Plakken` opent eerst een groot bewerkvenster. Daar kun je tekst handmatig plakken of met `Haal Uit Klembord` ophalen, controleren en aanpassen voordat je op `Toevoegen` klikt en een `SourceFile`-naam opgeeft. Dit is de veilige route voor grote exports of tekst die eerst opgeschoond moet worden.

Beide routes gebruiken daarna dezelfde Lovion-tekstparser en voegen de herkende rijen aan de huidige grid toe. Geen van beide knoppen wist bestaande rijen automatisch.

Wat de GUI kan:

* meerdere `txt`, `xlsx` en `csv` bestanden tegelijk laden
* eerste kolom is altijd `SourceFile`
* `Klembord Import` om de huidige clipboardtekst direct toe te voegen, met een eigen `SourceFile`-naam
* `Plakvenster` voor grote Lovion-tekstexports die je eerst wilt bekijken of handmatig wilt aanpassen, ook met een eigen `SourceFile`-naam
* `Stations` vraagt een lijst stationnummers, zoekt ieder station één keer vanuit de Lovion-lijst `LS stroomtransformatorgroep`, opent de eerste trafo via `EXPLORE > Start query > Open > VIEW` en importeert alle aansluitingen met de bestaande 100-regelmethode; meerdere trafos worden daarna uit dezelfde geladen eerste resultatenlijst geopend zonder opnieuw te zoeken, terwijl `SourceFile` exact het stationnummer blijft zonder batchnummer erachter
* `Lovion macro` toont de installatie-informatie voor `Trafo zoeker Macro` onder `GEN > Elektriciteit > LS stroomtransformatorgroep` en slaat met `XML downloaden` het gevalideerde XML-bestand op voor Lovion `Rapporten importeren`
* bewerkbare grid voor alle aansluitingen
* export van de huidige grid naar `csv`, `geojson` en `shapefile`
* shapefile gebruikt voor tekstexports expliciet `OverdrachtspuntRdX1` en `OverdrachtspuntRdY1`
* shapefile-export schrijft ook een `*_fieldmap.csv` weg

De GUI-build embedt het PowerShell-script in de exe. Op een Windows-machine met PowerShell start je dus in de praktijk met alleen `LovionCoordinateWorkbench.exe`.

De GUI-build embedt ook de macro-XML in de exe. Daardoor is voor de knop `Lovion macro` geen los XML-bestand naast de release-EXE nodig. In Lovion/Citrix moet de gebruiker de opgeslagen XML daarna nog zelf via `Rapporten importeren` installeren.

## Belangrijk

* `CoordinateSource = pdok` betekent: coordinaat komt uit een strikte PDOK-match op het adres.
* `CoordinateSource = pdok_tolerant` betekent: coordinaat komt uit een getolereerde PDOK-match met `PdokMatchReason = house_suffix_mismatch`.
* `CoordinateSource = wfs` betekent: coordinaat is gesnapt naar een WFS-feature.
* `CoordinateSource = overdrachtspunt_text` betekent: punt komt direct uit het `Overdrachtspunt` in de tekstexport.
* `PdokStrictMatch = False` met `PdokMatchReason = house_suffix_mismatch` is geen perfecte match, maar wordt wel als tolerante adresmatch gebruikt.
* `service_connection_point_already_used_keep_seed` of `street_furniture_point_already_used_keep_seed` betekent: er was wel een kandidaat, maar die lag al vast op een eerdere rij; de rij houdt dan zijn seed-coordinaat.
* In de parser geldt: alleen `Gebruiksdoelen` die matchen op de ingestelde street-furniture termen zoals `riool`, `camera`, `verlichting`, `container` gaan naar `Enexis_Opendata:asm_e_lv_street_furniture`.
* Alle andere `Gebruiksdoelen` gaan naar `Enexis_Opendata:asm_e_lv_service_connection`.
* Die termen worden gelezen uit `street_furniture_terms.json` naast de exe of naast het script. De voormalige GUI-knoppen voor PDOK, Enexis, Terms en Herlees zijn verwijderd; de onderliggende functies blijven in de broncode staan voor compatibiliteit.
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
