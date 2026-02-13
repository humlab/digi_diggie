
## Download the placename database

### Python scripts

### Prerequistes

 - Account and credentials to https://geotorget.lantmateriet.se/
 - Authority to download [Ortnamn, vector](https://geotorget.lantmateriet.se/geodataprodukter/ortnamn-nedladdning-vektor-api)
  
### Steps
 - Log in to https://geotorget.lantmateriet.se/
 - Open 
 - 


---

## 🇸🇪 **Lantmäteriets koordinatsystem**

Lantmäteriets geografiska data (t.ex. Ortnamn, GSD, Topografisk webbkarta, m.m.) levereras **nästan alltid i SWEREF 99 TM**, som har:

| Egenskap          | Värde                                  |
| ----------------- | -------------------------------------- |
| **System**        | SWEREF 99 TM                           |
| **EPSG / SRID**   | **3006**                               |
| **Enhet**         | meter                                  |
| **Koordinater**   | x ≈ 250000–850000, y ≈ 6100000–7700000 |
| **Arealreferens** | Plan (projektion)                      |
| **Typ**           | Projekterat koordinatsystem            |

Detta är **det svenska nationella referenssystemet** för kartdata.

---

## 🌍 **EPSG 4326 (WGS 84)**

| Egenskap        | Värde                                  |
| --------------- | -------------------------------------- |
| **System**      | WGS 84 (lat/long)                      |
| **SRID**        | 4326                                   |
| **Enhet**       | grader (decimal degrees)               |
| **Koordinater** | lat ≈ 55–69, lon ≈ 11–24               |
| **Typ**         | Geografiskt koordinatsystem (sfäriskt) |

Detta används främst för GPS, webbkartor (t.ex. Leaflet, OpenStreetMap, Google Maps), och datautbyte.

---

## 🔍 **Hur du ser vilket system din data använder**

Titta på dina koordinater (`nkoordinat`, `ekoordinat`):

| Kolumn       | Exempelvärde | Tolkning                                              |
| ------------ | ------------ | ----------------------------------------------------- |
| `nkoordinat` | 6588464      | → tydligt **Y i meter** (≈ 6,588 km norr om ekvatorn) |
| `ekoordinat` | 583500       | → **X i meter** (≈ 583 km östligt)                    |

Eftersom de är **sex–sju siffror långa** och i **meter**, är detta **SWEREF 99 TM (EPSG 3006)**, inte grader.
Om det vore 4326, skulle värdena ligga runt `lat ~ 59.3`, `lon ~ 17.0`.

---

## ✅ **Slutsats för Ortnamn**

För Lantmäteriets *Ortnamn*:

> Använd alltid **SRID 3006** (SWEREF 99 TM).

---

## 💡 **Exempel på korrekt PostGIS-hantering**

Efter att du importerat CSV utan `geom`-kolumnen:

```sql
ALTER TABLE placenames
  ADD COLUMN geom geometry(Point, 3006);

UPDATE placenames
  SET geom = ST_SetSRID(ST_MakePoint(ekoordinat, nkoordinat), 3006);
```

Vill du använda dem i t.ex. webbkarta (WGS84/4326):

```sql
SELECT ST_AsGeoJSON(ST_Transform(geom, 4326)) FROM placenames;
```

---
