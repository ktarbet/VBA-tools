Attribute VB_Name = "Interpolate"
' TableInterp to interpolate in two directions (given x and y find z)
' Karl Tarbet  
' September 2003
Function TableInterp(vertical_value As Double, horizontal_value As Double, _
               tbl As Range) As Double
Attribute TableInterp.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim row As Integer
    Dim col As Integer
    
    If vertical_value < tbl.Cells(2, 1) Or vertical_value > tbl.Cells(tbl.Cells.Rows.count, 1) Then
            TableInterp = CVErr("Error: with TableInterp")
       Exit Function
       End If
       
    If horizontal_value < tbl.Cells(1, 2) Or horizontal_value > tbl.Cells(1, tbl.Cells.Columns.count) Then
            TableInterp = CVErr("Error: with TableInterp")
      Exit Function
    End If
    
    ' find row and column to begin interpolation
    For col = 2 To tbl.Cells.Columns.count
       If (tbl.Cells(1, col + 1) >= horizontal_value) Then
          Exit For
       End If
    Next
    For row = 2 To tbl.Cells.Rows.count
       If (tbl.Cells(row + 1, 1) >= vertical_value) Then
          Exit For
       End If
    Next
    Dim verticalDistance, verticalpercent As Double
    Dim horizontalDistance, horizontalPercent As Double
    Dim vertInterp1, vertInterp2 As Double
    
    verticalDistance = tbl.Cells(row + 1, 1) - tbl.Cells(row, 1)
    If (verticalDistance <= 0) Then
      TableInterp = CVErr("Error: with TableInterp")
      Exit Function
    End If

    verticalpercent = (vertical_value - tbl.Cells(row, 1)) / verticalDistance

    horizontalDistance = tbl.Cells(1, col + 1) - tbl.Cells(1, col)
    If (horizontalDistance <= 0) Then
      TableInterp = CVErr("Error: with TableInterp")
      Exit Function
    End If

    horizontalPercent = (horizontal_value - tbl.Cells(1, col)) / horizontalDistance
    
    vertInterp1 = (1# - verticalpercent) * tbl.Cells(row, col) _
                 + verticalpercent * tbl.Cells(row + 1, col)
    vertInterp2 = (1# - verticalpercent) * tbl.Cells(row, col + 1) _
                 + verticalpercent * tbl.Cells(row + 1, col + 1)

    TableInterp = (1# - horizontalPercent) * vertInterp1 _
                + horizontalPercent * vertInterp2

End Function

Function Interp(x As Variant, Table As Object, _
  YCol As Integer) As Variant

  Dim TableRow As Integer, Temp As Variant
  Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double
  Dim d As Double

  On Error Resume Next
  Temp = Application.Match(x, Table.Resize(, 1), 1)
  If IsError(Temp) Then
    Interp = CVErr(Temp)
  Else
    TableRow = CInt(Temp)
    x0 = Table(TableRow, 1)
    y0 = Table(TableRow, YCol)
    If x = x0 Then
      Interp = y0
    Else
      x1 = Table(TableRow + 1, 1)
      y1 = Table(TableRow + 1, YCol)
      Interp = (x - x0) / (x1 - x0) * (y1 - y0) + y0
    End If
  End If
End Function

Function HInterp(y As Variant, Table As Object, _
  XRow As Integer) As Variant

  Dim TableCol As Integer, Temp As Variant
  Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double
  Dim d As Double
  Dim numColumns As Integer
  
  On Error Resume Next
  numColumns = Table.Columns.count
  Temp = Application.Match(y, Table.Resize(1, numColumns), 1)
  If IsError(Temp) Then
    HInterp = CVErr(Temp)
  Else
    TableCol = CInt(Temp)
    y0 = Table(1, TableCol)
    x0 = Table(XRow, TableCol)
    If y = y0 Then
      HInterp = x0
    Else
      y1 = Table(1, TableCol + 1)
      x1 = Table(XRow, TableCol + 1)
      HInterp = (y - y0) / (y1 - y0) * (x1 - x0) + x0
    End If
  End If
End Function
