Sub Analysis():
'MsgBox (Cells(2, 1))

endRow = Cells(Rows.Count, "A").End(xlUp).Row
'Cells(1, 1).Interior.ColorIndex = 4
summary_table = 2

Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double

For row_idx = 2 To endRow:
'iterator
' This is for adding the ticker to the same row
'Cells(row_idx, 9) = Cells(row_idx, 1)

' This is for adding the ticker to a different row
If Cells(row_idx, 1) <> Cells(row_idx - 1, 1) Then
    Cells(summary_table, 9) = Cells(row_idx, 1)
    
    openingPrice = Cells(row_idx, 3)
End If

If Cells(row_idx, 1) <> Cells(row_idx + 1, 1) Then
    closingPrice = Cells(row_idx, 6)
    yearlyChange = closingPrice - openingPrice
    Cells(summary_table, 10) = yearlyChange
    
    If yearlyChange >= 0 Then
        Cells(summary_table, 10).Interior.ColorIndex = 4
    Else
        Cells(summary_table, 10).Interior.ColorIndex = 3
    End If
    
    summary_table = summary_table + 1
End If

Next row_idx

End Sub

