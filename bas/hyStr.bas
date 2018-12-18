Attribute VB_Name = "hyStr"
Option Explicit

Function removeElementOfList(el As String, list() As String)

    '•¡”—v‘f‚É‚à‘Î‰‚³‚¹‚½‚¢‚ªA¡‰ñ‚Í•Û—¯
    
    Dim i As Integer
    Dim flg As Boolean: flg = False
    For i = LBound(list) To UBound(list)
        If list(i) = el Then
            list(i) = list(i + 1)
            i = i + 1
            flg = True
        End If
    Next
    
    If flg Then
        ReDim Preserve list(UBound(list) - 1)
    End If
    
    removeElementOfList = list
    
End Function

Function getBaseName(mdl As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    getBaseName = fso.getBaseName(mdl)
    Set fso = Nothing
End Function
