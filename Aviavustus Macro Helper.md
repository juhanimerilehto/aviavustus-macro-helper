# Aviavustus Macro Helper
Version 1.0.0
### Creator: Juhani Merilehto - @juhanimerilehto - Jyväskylä University of Applied Sciences (JAMK), Likes Institute

![JAMK Likes Logo](jamklikes.png)

# Excel Sheet Combiner for Aviavustukset.fi | Excel-taulukoiden yhdistäjä Aviavustukset.fi:lle

A VBA-based tool to combine multiple municipality sheets from aviavustukset.fi Excel exports into a single, analyzable dataset.

VBA-pohjainen työkalu, joka tarvittaessa yhdistää aviavustukset.fi-sivuston Excel-viennin useat kuntataulukot yhdeksi analysoitavaksi tietojoukoksi.

## Background | Tausta

**English:**  
When downloading grant data from [aviavustukset.fi](https://aviavustukset.fi/), the Excel file can some times be formatted to contains separate sheets for i.e., each municipality. This tool helps combine all sheets into a single list while preserving the municipality information.

**Suomi:**  
Kun avustustietoja ladataan [aviavustukset.fi](https://aviavustukset.fi/)-sivustolta, voi Excel-tiedosto sisältää erilliset välilehdet esim. jokaiselle kunnalle. Tämä työkalu auttaa yhdistämään kaikki välilehdet yhdeksi listaksi säilyttäen kuntatiedot.

## Features | Ominaisuudet

**English:**
- Automatically combines all municipality sheets into one consolidated view
- Adds a "Municipality" column to track the source of each entry
- Handles large datasets (tested with 200+ municipality sheets)
- Preserves original data columns (change as required for you case):
  - Toteuttaja (Implementer)
  - Hankkeen nimi (Project name)
  - Avustusmuoto (Type of grant)
  - Myöntövuosi (Year granted)
  - Myönnetty avustus (Granted amount)

**Suomi:**
- Yhdistää automaattisesti kaikki kuntataulukot yhteen näkymään
- Lisää "Municipality"-sarakkeen jokaisen merkinnän lähteen seuraamiseksi
- Käsittelee suuria tietojoukkoja (testattu yli 200 kuntataulukolla)
- Säilyttää alkuperäiset tietosarakkeet (vaihda oman käyttötarpeesi mukaan):
  - Toteuttaja
  - Hankkeen nimi
  - Avustusmuoto
  - Myöntövuosi
  - Myönnetty avustus

## Prerequisites | Edellytykset

**English:**
- Microsoft Excel (any modern version)
- Excel file from aviavustukset.fi with municipality sheets

**Suomi:**
- Microsoft Excel (mikä tahansa nykyaikainen versio)
- Excel-tiedosto aviavustukset.fi-sivustolta kuntataulukoineen

## Installation | Asennus

**English:**
1. Open your Excel file containing the municipality sheets
2. Press `Alt + F11` to open the VBA editor
3. Insert > Module
4. Copy and paste the following code (also as separate file "aviavustus-macro.md"):

**Suomi:**
1. Avaa Excel-tiedosto, joka sisältää kuntataulukot
2. Paina `Alt + F11` avataksesi VBA-editorin
3. Insert > Module
4. Kopioi ja liitä seuraava koodi (myös erilisenä tiedostona "aviavustus-macro.md"):

```vba
Sub CombineSheets()
    Dim ws As Worksheet
    Dim combinedSheet As Worksheet
    Dim lastRow As Long
    Dim copyRange As Range
    Dim pasteRow As Long
    
    ' Add a new sheet for combined data
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Combined").Delete
    Application.DisplayAlerts = True
    Set combinedSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    combinedSheet.Name = "Combined"
    On Error GoTo 0
    
    ' Add headers
    With combinedSheet
        .Cells(1, 1) = "Toteuttaja"
        .Cells(1, 2) = "Hankkeen nimi"
        .Cells(1, 3) = "Avustusmuoto"
        .Cells(1, 4) = "Myöntövuosi"
        .Cells(1, 5) = "Myönnetty avustus"
        .Cells(1, 6) = "Municipality"
    End With
    
    ' Start pasting row
    pasteRow = 2
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Sheets
        ' Skip the combined sheet
        If ws.Name <> "Combined" Then
            ' Find last row in current sheet
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' If sheet has data (more than just headers)
            If lastRow > 1 Then
                ' Copy data range
                Set copyRange = ws.Range("A2:E" & lastRow)
                
                ' Paste to combined sheet
                copyRange.Copy combinedSheet.Cells(pasteRow, 1)
                
                ' Fill municipality name
                combinedSheet.Range("F" & pasteRow & ":F" & (pasteRow + lastRow - 2)).Value = ws.Name
                
                ' Update paste row for next iteration
                pasteRow = pasteRow + lastRow - 1
            End If
        End If
    Next ws
    
    ' Format as table (optional)
    combinedSheet.UsedRange.Select
    combinedSheet.ListObjects.Add(xlSrcRange, combinedSheet.UsedRange, , xlYes).Name = "CombinedTable"
    
    ' Autofit columns
    combinedSheet.UsedRange.Columns.AutoFit
    
    MsgBox "Sheets have been combined!", vbInformation
End Sub
```

## Usage | Käyttö

**English:**
1. Save your aviavustukset.fi Excel file as `.xlsm` (Excel Macro-Enabled Workbook)
2. Press `Alt + F8` to open the Macro dialog
3. Select "CombineSheets" and click "Run"
4. A new sheet named "Combined" will be created at the beginning of your workbook with all data consolidated

**Suomi:**
1. Tallenna aviavustukset.fi Excel-tiedosto `.xlsm`-muodossa (Excel-makrot käyttävä työkirja)
2. Paina `Alt + F8` avataksesi Makro-valintaikkunan
3. Valitse "CombineSheets" ja klikkaa "Suorita"
4. Uusi "Combined"-niminen taulukko luodaan työkirjasi alkuun kaikilla yhdistetyillä tiedoilla

## Important Notes | Tärkeitä huomioita

**English:**
- Always keep a backup of your original file before running the macro
- The macro assumes your data starts from row 2 in each sheet (with row 1 being headers)
- The combined data will be formatted as a table for easy filtering and sorting

**Suomi:**
- Pidä aina varmuuskopio alkuperäisestä tiedostosta ennen makron suorittamista
- Makro olettaa tietojen alkavan riviltä 2 jokaisella välilehdellä (rivin 1 ollessa otsikkorivi)
- Yhdistetyt tiedot muotoillaan taulukoksi helppoa suodatusta ja lajittelua varten



## Contributing | Osallistuminen

**English:**  
If you encounter any issues or have suggestions for improvements, please create an issue in this repository.

**Suomi:**  
Jos kohtaat ongelmia tai sinulla on parannusehdotuksia, ole hyvä ja luo issue tähän repositorioon.

## License | Lisenssi

**English:**
This project is licensed for free use under the condition that proper credit is given to Juhani Merilehto (@juhanimerilehto) and JAMK Likes institute. You are free to use, modify, and distribute this project, provided that you mention the original author and institution and do not hold them liable for any consequences arising from the use of the algorithm.

**Suomi:**
Tämä projekti on lisensoitu vapaaseen käyttöön sillä ehdolla, että asianmukainen kunnia annetaan Juhani Merilehdolle (@juhanimerilehto) ja JAMK Likes-instituutille. Voit vapaasti käyttää, muokata ja jakaa tätä projektia, kunhan mainitset alkuperäisen tekijän ja instituution, etkä pidä heitä vastuussa algoritmin käytöstä aiheutuvista seurauksista.

## Acknowledgments | Kiitokset

**English:**
- [Aviavustukset.fi](https://aviavustukset.fi/) for providing the grant data
- JAMK University of Applied Sciences
- Likes institute

**Suomi:**
- [Aviavustukset.fi](https://aviavustukset.fi/) avustustietojen tarjoamisesta
- Jyväskylän ammattikorkeakoulu
- Likes-tutkimuskeskus

## Author | Tekijä

Juhani Merilehto - [@juhanimerilehto](https://github.com/juhanimerilehto)