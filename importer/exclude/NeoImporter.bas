Attribute VB_Name = "NeoImporter"
Option Explicit

Dim IMPORT_PATH As String
Dim MODULE_PATH As String
Dim EXCLUSION_LIST As Variant
Dim IGNORE_FILE As String

Sub Bundler()
    MODULE_PATH = ThisWorkbook.PATH & "\importer"
    IMPORT_PATH = MODULE_PATH & "\bas"
    IGNORE_FILE = MODULE_PATH & "\.importignore"
    EXCLUSION_LIST = Array("NeoImporter")
    Call AddExclusionList

    Call Main
    
    IMPORT_PATH = vbNullString
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
    
    cnt = UBound(EXCLUSION_LIST)
    ReDim Preserve EXCLUSION_LIST(cnt + eol)
    
    For i = cnt + 1 To UBound(EXCLUSION_LIST)
        EXCLUSION_LIST(i) = lns(j)
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
    For i = LBound(EXCLUSION_LIST) To UBound(EXCLUSION_LIST)
        If ModuleName = EXCLUSION_LIST(i) Then
            ExclusionCheck = True
            Exit Function
        End If
    Next
    ExclusionCheck = False
End Function
