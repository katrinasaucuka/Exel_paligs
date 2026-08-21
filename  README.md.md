# Excel Palīgs

Neliela darbvirsmas lietotne, kas paātrina un padara precīzāku datu ievadi Excel failos — lietotājam vairs nav jāatceras, kurā rindā vai kolonnā ierakstīt informāciju.

## Ko tas dara

- Atver vienkāršu ievades formu ar kategorijas izvēli un laukiem: Cena, Kas, Datums.
- Pēc datu ievades un pogas "Pievienot ierakstu" nospiešanas automātiski atrod pareizo vietu Excel failā.
- Ieraksta datus, sakārtojot pēc datuma (vecākais ieraksts augšā, jaunākais apakšā).

## Kāpēc šis rīks noderīgs

Manuāli strādājot ar Excel, jāatceras kurā kolonnā ir konkrētās kategorijas dati, kurā rindā jāieraksta jaunais ieraksts, un jāseko, lai datumi paliktu pareizā secībā. Šis rīks visu šo dara automātiski — pietiek ievadīt informāciju un nospiest vienu pogu.

## Izmantotās tehnoloģijas

| Komponente | Tehnoloģija |
|---|---|
| Valoda | Python |
| Grafiskā saskarne | `tkinter` |
| Excel apstrāde | `openpyxl` |

**Datu drošība:** programma darbojas pilnībā lokāli — visi dati paliek lietotāja datorā, nekāda informācija netiek sūtīta internetā.

## Uzstādīšana

```bash
pip install openpyxl
```

## Lietošana

1. Atver failu `excel_paligs.py` un norādi ceļu uz savu Excel failu (11. rindā).
2. Palaid skriptu:

```bash
python excel_paligs.py
```

3. Izvēlies kategoriju, ievadi datus (Cena, Kas, Datums) un nospied "Pievienot ierakstu".

## Ierobežojumi

- Paredzēts lokālai lietošanai, vienam lietotājam vienlaikus — nav paredzēts vairāku cilvēku vienlaicīgai piekļuvei tam pašam failam.
- Nav iebūvētas datu validācijas (piem., datuma formāta pārbaudes).

## Iespējamie uzlabojumi

- Pārveidot par patstāvīgu `.exe` failu (ar `PyInstaller`), lai nav nepieciešams Python.
- Pievienot ievades datu validāciju.
- Pāriet uz koplietojamu risinājumu (piem., Google Sheets), lai vairāki cilvēki var lietot vienlaicīgi.

---

**Autors:** Katrīna Saučuka
