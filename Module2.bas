Attribute VB_Name = "Module2"
Sub Yellow()
    Dim lastRow As Long
    'Dim txEndRow As Long
    Dim rng As Range
    Dim ws As Worksheet
    Dim rs As Worksheet
    Dim totalTrnNumber As Long
    Dim totalTrnAmount As Double
    Dim bigTrnNumber As Long
    Dim bigTrnAmount As Double
    Dim j As Integer
    
    ' Set the range of the table (adjust the sheet name and range as per your needs)
    Set ws = ThisWorkbook.Worksheets("Sheet (1)")
    Set rng = ws.UsedRange
    'Range ("B2:E21")
    lastRow = rng.Rows.Count + rng.Row - 1
    'txEndRow = lastRow - 2
    
    ' Initialize variables
    totalTrnNumber = 0
    totalTrnAmount = 0
    bigTrnNumber = 0
    bigTrnAmount = 0
    
    Dim rheck As Boolean
    For Each Sheet In Worksheets
    If Sheet.Name Like "REMOVED" Then rheck = True: Exit For
    Next
    
    If rheck <> True Then
    Sheets.Add.Name = "REMOVED"
    Set rs = ThisWorkbook.Worksheets("REMOVED")
    ws.Rows(1).Copy rs.Rows(1)
    End If
    
    Set rs = ThisWorkbook.Worksheets("REMOVED")
    
           
    ' Loop through each row in the range
    For i = 2 To lastRow - 2
        If rng.Cells(i, 2).Interior.Color = RGB(255, 255, 0) Then
                j = (rs.UsedRange.Rows.Count) + 1
                rng.Rows(i).Copy rs.UsedRange.Rows(j)
                rng.Rows(i).Delete
                'txEndRow = txEndRow - 1
                i = i - 1 ' Adjust the index since rows have shifted up
        End If
    Next i
        
    lastRow = rng.Rows.Count + rng.Row - 1
        
    ' Loop through each row in the range
    For i = 2 To lastRow - 2
                ' Check if the REPORTING AMOUNT is blank
                If rng.Cells(i, 2).Value <> "" Then
                    ' Update the segment totals
                    totalTrnNumber = totalTrnNumber + 1
                    totalTrnAmount = totalTrnAmount + rng.Cells(i, 2).Value
                Else
                    ' Update Total Trn Number and Total Trn Amount for the above segment
                    rng.Cells(i, 3).Value = totalTrnNumber
                    rng.Cells(i, 4).Value = totalTrnAmount
                    ' Update sheet total
                    bigTrnNumber = bigTrnNumber + totalTrnNumber
                    bigTrnAmount = bigTrnAmount + totalTrnAmount
                    ' Clear before moving onto next segment
                    totalTrnNumber = 0
                    totalTrnAmount = 0
                End If
    Next i
   
    rng.Cells(lastRow, 3).Value = bigTrnNumber
    rng.Cells(lastRow, 4).Value = bigTrnAmount

    ' Loop through each row in the range
        For i = 2 To lastRow - 2
                    ' Check if the REPORTING AMOUNT is blank
                    If rng.Cells(i, 3).Value <> "" Then
                        If rng.Cells(i, 3).Value = 0 Then
                        rng.Rows(i).Delete
                        i = i - 1
                        lastRow = lastRow - 1
                        End If
                    rng.Cells(i, 5).Value = rng.Cells(i, 4).Value / bigTrnAmount
                    End If
        Next i

    Sheets("REMOVED").Select
    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Cells.EntireColumn.AutoFit

End Sub
