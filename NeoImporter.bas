Attribute VB_Name = "NeoImporter"
Option Explicit

Dim IMPORT_PATH As String
Dim MY_LIST()
Dim MODULE_PATH As String
Dim KESU_LIST As Variant
Dim IGNORE_FILE As String

Sub Bundler()
    MODULE_PATH = ThisWorkbook.PATH & "\module"
    IMPORT_PATH = MODULE_PATH & "\mylist"
    IGNORE_FILE = MODULE_PATH & "\.listignore"
    KESU_LIST = Array("Washoi")
    
    Call GetImportList
    Call ExcludeElementsOfList
    
    Call AddExclusionList
    Call Main
    
    IMPORT_PATH = vbNullString
End Sub

Sub ExcludeElementsOfList()
    Dim i As Integer
    Dim j As Integer
    Dim idx As Integer: idx = 0
    
    For i = LBound(KESU_LIST) To UBound(KESU_LIST)
        For j = LBound(MY_LIST) To UBound(MY_LIST)
            If MY_LIST(j) = KESU_LIST(i) Then
                idx = i
                Exit For
            End If
            Debug.Print j & " " & MY_LIST(j)
        Next
        For j = idx + 1 To UBound(MY_LIST)
            MY_LIST(j) = MY_LIST(j + 1)
            Debug.Print j & " " & MY_LIST(j)
        Next
        ReDim Preserve MY_LIST(UBound(MY_LIST) - 1)
    Next
End Sub

Sub GetImportList()
    Dim fso As Object
    Dim fs As Object
    Dim f As Object
    Dim i As Integer
    Dim cnt As Integer
    Dim mn As String
    Dim mbn As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.GetFolder(IMPORT_PATH).files
    
    cnt = fs.Count

    ReDim MY_LIST(cnt)
    
    For Each f In fs
        MY_LIST(i) = f.Name
    Next
    
    Set fso = Nothing
    Set fs = Nothing
    Set f = Nothing
End Sub

Sub AddExclusionList()
    'ìKìñÇ»ÇÃÇ≈óvèCê≥
    Dim fso As Object
    Dim f As Object
    Dim buf As String
    Dim eol As Integer
    Dim cnt As Integer
    Dim lns
    Dim i As Integer: i = 0
    Dim j As Integer: j = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(IGNORE_FILE, 1)
    buf = f.ReadAll
    lns = Split(buf, vbCrLf)
    eol = UBound(lns)
    
    cnt = UBound(KESU_LIST)
    ReDim Preserve KESU_LIST(cnt + eol)
    
    For i = cnt + 1 To UBound(KESU_LIST)
        KESU_LIST(i) = lns(j)
        j = j + 1
    Next
    
    f.Close
    Set f = Nothing
    Set fso = Nothing
End Sub

Sub Main()
    Dim fso As Object
    Dim fs As Object
    Dim f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.GetFolder(IMPORT_PATH).files
    Dim i As Integer
    Dim mn As String
    Dim mbn As String

    For Each f In fs
        mn = f.Name
        mbn = fso.getBaseName(mn)
        If Not ExclusionCheck(mbn) Then
            ThisWorkbook.VBProject.VBComponents.IMPORT IMPORT_PATH & "\" & mn
        End If
    Next
    Set fso = Nothing
    Set fs = Nothing
    Set f = Nothing
End Sub

Function ExclusionCheck(ModuleName As String) As Boolean
    Dim i As Integer
    For i = LBound(KESU_LIST) To UBound(KESU_LIST)
        If ModuleName = KESU_LIST(i) Then
            ExclusionCheck = True
            Exit Function
        End If
    Next
    ExclusionCheck = False
End Function
