# ExcelToJsonParser

Parser minimal pentru fișiere RACASDIA INTERSOFT -> JSON

Ce face:
- Detectează automat rândul de header (primele 15 rânduri)
- Extrage coloanele: număr ordine, simbol, denumire, um
- Extrage cantități și prețuri pentru materiale, manoperă, utilaje și transport (dacă sunt prezente)
- Calculează totaluri și validează față de totalul găsit în sheet (dacă există)
- Scrie un fișier JSON cu structura normalizată

Prerechizite:
- .NET 7 SDK

Cum se folosește (PowerShell):

```powershell
cd "c:\Users\alexa\OneDrive\Desktop\proiecte\TEST\test\ExcelToJsonParser"
# restore
dotnet restore
# build
dotnet build -c Release
# run
dotnet run -- "C:\cale\către\RACASDIA_INTERSOFT.xlsx" "C:\cale\către\ieșire.json"
```

Observații:
- În funcție de structura exactă a fișierului, s-ar putea să fie nevoie să ajustăm candidatul de header sau mapările.
- După ce îmi trimiți fișierul, îl rulez local și adaptez mapările dacă e cazul.
