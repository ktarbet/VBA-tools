' class TextFile
' basic text file Class
' Karl Tarbet  October 2013

Option Explicit



Private data() As String
'Private count As Integer

Public Property Get Lines(index As Integer) As String
  Lines = data(index)
End Property

Public Property Get Size() As Integer
   Size = UBound(data) - LBound(data) + 1
End Property

'Public Property Let Vec(index As Long, Value As  )
'  v(index) = Value
'End Property

Public Function IndexOfBoth(str As String, str2 As String) As Integer

Dim i As Integer
Dim idx As Integer
Dim idx2 As Integer
IndexOfBoth = -1

For i = 1 To Size()
   idx = InStr(1, data(i), str)
   idx2 = InStr(1, data(i), str2)
   If idx > 0 And idx2 > 0 Then
   IndexOfBoth = i
   Exit For
   End If
Next

End Function

Public Property Get IndexOf(str As String, Optional startIndex = 1) As Integer
Dim i As Integer
Dim idx As Integer

IndexOf = -1

For i = startIndex To Size()
   idx = InStr(1, data(i), str)
   
   If idx > 0 Then
   IndexOf = i
   Exit For
   End If
Next


End Property

Public Sub ReadFile(fileName As String)

    Dim intFileNum As Integer
    Dim buff As String
    intFileNum = FreeFile
    Open fileName For Input As #intFileNum
    buff = Input(LOF(1), intFileNum)
    data = Split(buff, vbLf)
    Close #intFileNum
    
End Sub

