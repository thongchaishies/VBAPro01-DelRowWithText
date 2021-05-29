# Example Code


Sub DeleteRowsWithText()

Const MyTarget = "Break Line" ' <-- change to suit
  
  Dim rng As Range
  Dim i As Long, j As Long
  
  ' Calc last row number
  j = Range("B" & Rows.Count).End(xlUp).Row 'Cells.SpecialCells(xlCellTypeLastCell).Row  'can be: j = Range("C" & Rows.Count).End(xlUp).Row
  
  ' Collect rows with MyTarget   taget is column A
  For i = j To 1 Step -1
    If WorksheetFunction.CountIf(Range("B" & i), MyTarget) > 0 Then
      If rng Is Nothing Then
        Set rng = Rows(i)
      Else
        Set rng = Union(rng, Rows(i))
      End If
    End If
  Next
  
  ' Delete rows with MyTarget
  If Not rng Is Nothing Then rng.Delete
  
  ' Update UsedRange
  With ActiveSheet.UsedRange: End With

End Sub
