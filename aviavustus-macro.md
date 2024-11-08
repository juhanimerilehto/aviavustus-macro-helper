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