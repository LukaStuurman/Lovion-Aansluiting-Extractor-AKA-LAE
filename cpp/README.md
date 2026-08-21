# Lovion Coordinate Tool C++

Native C++ CLI-port van de Lovion Coordinate Tool.

## Status

Deze versie ondersteunt:

* Lovion TXT-export met `Overdrachtspunt`
* Lovion CSV-lijsten met adreskolommen
* PDOK-geocoding
* Enexis WFS-snapping of lokale WFS GeoJSON
* export naar CSV, GeoJSON en shapefile

Deze versie ondersteunt nog niet:

* directe `.xlsx`-invoer en `.xlsx`-terugschrijven
* de Windows Forms workbench/GUI

Gebruik daarvoor voorlopig de bestaande PowerShell-versie. Voor de C++ CLI kun je een Excel-werkblad opslaan als CSV.

## Build

De build vereist Visual Studio Build Tools (`cl.exe`) of MinGW-w64 (`g++.exe`) in `PATH`.

Met CMake:

```powershell
cmake -S .\cpp -B .\cpp\build
cmake --build .\cpp\build --config Release
```

Met de Visual Studio-generator staat de output daarna doorgaans hier:

```text
.\cpp\build\Release\LovionCoordinateToolCpp.exe
```

## Gebruik

TXT-export:

```powershell
.\dist_cpp\LovionCoordinateToolCpp.exe `
  -InputTextPath "C:\pad\naar\lovion-export.txt" `
  -OutputDirectory ".\output_cpp"
```

CSV met PDOK en WFS:

```powershell
.\dist_cpp\LovionCoordinateToolCpp.exe `
  -InputCsvPath "C:\pad\naar\lovion.csv" `
  -OutputDirectory ".\output_cpp"
```

Alleen PDOK, zonder WFS:

```powershell
.\dist_cpp\LovionCoordinateToolCpp.exe `
  -InputCsvPath "C:\pad\naar\lovion.csv" `
  -SkipWfs
```

Lokale WFS GeoJSON:

```powershell
.\dist_cpp\LovionCoordinateToolCpp.exe `
  -InputCsvPath "C:\pad\naar\lovion.csv" `
  -WfsGeoJsonPath "C:\pad\naar\asm_e_lv_service_connection.geojson"
```

Kleine test:

```powershell
.\dist_cpp\LovionCoordinateToolCpp.exe `
  -InputCsvPath "C:\pad\naar\lovion.csv" `
  -SkipWfs `
  -MaxRows 25
```
