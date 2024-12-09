Attribute VB_Name = "Module1"
Sub ticker()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim i As Long
    Dim exists As Boolean
    Dim checkRow As Long
    Dim nextRow As Long
    
    Set ws = ThisWorkbook.Sheets("A")
    
    If ws.Cells(1, 11).Value <> "Ticker Symbols" Then
        ws.Cells(1, 11).Value = "Ticker Symbols"
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    nextRow = 2
    
    For i = 2 To lastRow
        tickerSymbol = ws.Cells(i, 1).Value
        
        If tickerSymbol <> "" Then
            exists = False
            
            For checkRow = 2 To nextRow - 1
                If ws.Cells(checkRow, 11).Value = tickerSymbol Then
                    exists = True
                    Exit For
                End If
            Next checkRow
            
            If Not exists Then
                ws.Cells(nextRow, 11).Value = tickerSymbol
                nextRow = nextRow + 1
            End If
        End If
    Next i
End Sub
Sub convert()
    Dim cell As Range
    Dim lastRow As Long
   
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Each cell In Range("A1:A" & lastRow)
        
        If IsDate(cell.Value) Then
            
            cell.Value = CDate(cell.Value)
        End If
    Next cell
End Sub


