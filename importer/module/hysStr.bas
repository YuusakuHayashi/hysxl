Attribute VB_Name = "hysStr"
Option Explicit

Function ExcludeElementsOfList(elements, list)

    Dim i As Integer
    Dim j As Integer
    Dim idx As Integer: idx = 0
    Dim flag As Boolean: flag = False
    
    For i = LBound(elements) To UBound(elements)
        For j = LBound(list) To UBound(list)
            If list(j) = elements(i) Then
                idx = j
                flag = True
                Exit For
            End If
        Next
        If flag Then
            For j = idx To UBound(list) - 1
                list(j) = list(j + 1)
            Next
            ReDim Preserve list(UBound(list) - 1)
            flag = False
        End If
    Next
    ExcludeElementsOfList = list
End Function

Function GetListOfFileLines(FileName)
    '“K“–‚È‚Ì‚Å—vC³
    Dim fso As Object
    Dim f As Object
    Dim buf As String
    Dim eol As Integer
    Dim cnt As Integer
    Dim lns
    Dim i As Integer: i = 0
    Dim j As Integer: j = 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(FileName, 1)
    buf = f.ReadAll
    lns = Split(buf, vbCrLf)
    GetListOfFileLines = lns
    
    f.Close
    Set f = Nothing
    Set fso = Nothing
End Function

Function AddElementsOfList(elements, list)

    Dim el_cnt As Integer
    Dim ls_cnt As Integer
    Dim i As Integer: i = 0
    Dim j As Integer: j = 0
    
    ls_cnt = UBound(list)
    el_cnt = UBound(elements)
    
    ReDim Preserve list(ls_cnt + el_cnt)
    
    For i = ls_cnt + 1 To UBound(list)
        list(i) = elements(j)
        j = j + 1
    Next
    
    AddElementsOfList = list
    
End Function

Function GetFileListOfFolder(Folder)
    Dim fso As Object
    Dim fs As Object
    Dim f As Object
    Dim i As Integer: i = 0
    Dim cnt As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.GetFolder(IMPORT_PATH).files

    ReDim IMPORT_LIST(fs.Count - 1)
    
    For Each f In fs
        IMPORT_LIST(i) = f.Name
        i = i + 1
    Next
    
    Set fso = Nothing
    Set fs = Nothing
    Set f = Nothing
End Function

Function RemoveElement(el, list() As String)
    '•¡”—v‘f‚É‚à‘Î‰ž‚³‚¹‚½‚¢‚ªA¡‰ñ‚Í•Û—¯
    Dim i As Integer
    
    Select Case TypeName(el)
        Case "String"
    End Select

    For i = el To UBound(list) - 1
        list(i) = list(i + 1)
    Next
    
    ReDim Preserve list(UBound(list) - 1)

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
