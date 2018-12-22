Attribute VB_Name = "hysStr"
Option Explicit

Function RemoveElement(el, list() As String)
    '•¡”—v‘f‚É‚à‘Î‰ž‚³‚¹‚½‚¢‚ªA¡‰ñ‚Í•Û—¯
    Dim i As Integer
    Dim flg As Boolean: flg = False
    
    Select Case TypeName(el)
        Case "String"
    End Select

    For i = el To UBound(list) - 1
        'If list(i) = list(el) Then
            list(i) = list(i + 1)
            'flg = True
        'End If
    Next
    
    If flg Then
        ReDim Preserve list(UBound(list) - 1)
    End If
    RemoveElement = list
    
End Function


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
            str = str & String(s, "ï¿½@")
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
