Attribute VB_Name = "NeoImporter"
Option Explicit

Dim IMPORT_PATH As String
Dim IMPORT_LIST()
Dim MODULE_PATH As String
Dim EXCLUSION_LIST As Variant
Dim IGNORE_FILE As String

Sub Bundler()
    MODULE_PATH = ThisWorkbook.PATH & "\importer"
    IMPORT_PATH = MODULE_PATH & "\bas"
    IGNORE_FILE = MODULE_PATH & "\.importignore"
    EXCLUSION_LIST = Array("NeoImporter.bas")
    
    Call GetImportList
    Call AddElementOfExclusionList
    Call ExcludeElementsOfList
    
    Call Main
    
    IMPORT_PATH = vbNullString
End Sub

Sub AddElementOfExclusionList()
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

Sub ExcludeElementsOfList()

    Dim i As Integer
    Dim j As Integer
    Dim idx As Integer: idx = 0
    Dim flag As Boolean: flag = False
    
    For i = LBound(EXCLUSION_LIST) To UBound(EXCLUSION_LIST)
        For j = LBound(IMPORT_LIST) To UBound(IMPORT_LIST)
            If IMPORT_LIST(j) = EXCLUSION_LIST(i) Then
                idx = j
                flag = True
                Exit For
            End If
        Next
        If flag Then
            For j = idx To UBound(IMPORT_LIST) - 1
                IMPORT_LIST(j) = IMPORT_LIST(j + 1)
            Next
            ReDim Preserve IMPORT_LIST(UBound(IMPORT_LIST) - 1)
            flag = False
        End If
    Next
End Sub

Sub GetImportList()
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
End Sub

'Sub AddExclusionList()
'    ìKìñÇ»ÇÃÇ≈óvèCê≥
'    Dim fso As Object
'    Dim f As Object
'    Dim buf As String
'    Dim eol As Integer
'    Dim cnt As Integer
'    Dim lns
'    Dim i As Integer: i = 0
'    Dim j As Integer: j = 0
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set f = fso.OpenTextFile(IGNORE_FILE, 1)
'    buf = f.ReadAll
'    lns = Split(buf, vbCrLf)
'    eol = UBound(lns)
'
'    cnt = UBound(EXCLUSION_LIST)
'    ReDim Preserve EXCLUSION_LIST(cnt + eol)
'
'    For i = cnt + 1 To UBound(EXCLUSION_LIST)
'        EXCLUSION_LIST(i) = lns(j)
'        j = j + 1
'    Next
'
'    f.Close
'    Set f = Nothing
'    Set fso = Nothing
'End Sub

Sub Main()
    Dim i As Integer
    For i = LBound(IMPORT_LIST) To UBound(IMPORT_LIST)
        ThisWorkbook.VBProject.VBComponents.IMPORT IMPORT_PATH & "\" & IMPORT_LIST(i)
    Next
End Sub

'Function ExclusionCheck(ModuleName As String) As Boolean
'    Dim i As Integer
'    For i = LBound(EXCLUSION_LIST) To UBound(EXCLUSION_LIST)
'        If ModuleName = EXCLUSION_LIST(i) Then
'            ExclusionCheck = True
'            Exit Function
'        End If
'    Next
'    ExclusionCheck = False
'End Function
