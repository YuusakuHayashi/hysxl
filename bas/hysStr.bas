Attribute VB_Name = "hysStr"
Option Explicit


Function zeroPadding(str As String, lng As Integer) As String
    Dim s As Integer
    s = lng - Len(str)
    If s > 0 Then
        str = String(s, "0") & str
    End If
    zeroPadding = str
End Function


Function spacePadding(str As String, lng As Integer, Optional t As String = "H") As String
    Dim s As Integer
    s = lng - Len(str)
    
    If s > 0 Then
        If t = "H" Then
            str = Space(s) & str
        Else
            str = str & String(s, "�@")
        End If
    End If
    spacePadding = str
End Function


Function cutRight(str As String, lng As Integer) As String
    Dim s As Integer
    s = Len(str) - lng
    If s > 0 Then
        str = Left(str, lng)
    End If
    cutRight = str
End Function
<!--stackedit_data:
eyJoaXN0b3J5IjpbMTgzNTI2Mzk1Nl19
-->
