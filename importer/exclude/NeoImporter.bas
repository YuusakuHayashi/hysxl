Attribute VB_Name = "NeoImporter"
Option Explicit

Dim IMPORT_PATH As String
Dim IMPORT_LIST()
Dim MODULE_PATH As String
Dim EXCLUSION_LIST As Variant
Dim IGNORE_FILE As String
Dim FORCE_FILE_PATH As String
Dim FORCED_LIST
Dim FORCED_MODULE
Dim FORCED_XL

Sub Bundler()
    MODULE_PATH = ThisWorkbook.path & "\importer"
    IMPORT_PATH = MODULE_PATH & "\module"
    IGNORE_FILE = MODULE_PATH & "\.importignore"
    EXCLUSION_LIST = Array("NeoImporter.bas")
    FORCE_FILE_PATH = MODULE_PATH & "\force"
    
    Call GetImportList
    Call AddElementsOfExclusionList
    Call ExcludeElementsOfList
    Call Main
    
'    Call GetForcedList
'    Call Force
'    IMPORT_PATH = vbNullString
End Sub

Sub Force()
    Dim i As Integer
    Dim j As Integer
    For i = LBound(FORCED_XL, 2) To UBound(FORCED_XL, 2)
        For j = LBound(FORCED_XL) To UBound(FORCED_XL)
            Workbook(FORCED_XL(j, i)).VBProject.vbcomponents.IMPORT IMPORT_PATH & FORCED_MODULE(i)
            'ThisWorkbook.VBProject.VBComponents.IMPORT IMPORT_PATH & "\" & IMPORT_LIST(i)
        Next
    Next
End Sub

Sub GetForcedList()
    Dim fso As Object
    Dim fs As Object
    Dim f As Object
    Dim buf As String
    Dim i As Integer
    Dim lns
    Dim line_cnt As Integer
    Dim file_cnt As Integer
    Dim x As Integer: x = 0
    Dim y As Integer: y = 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.GetFolder(FORCE_FILE_PATH).files
    
    file_cnt = fs.Count
    ReDim FORCED_XL(1000, 100)
    ReDim FORCED_MODULE(f.Count)
    
    file_cnt = 0
    line_cnt = 0
    
    For Each f In fs
        buf = fso.OpenTextFile(f.Name).ReadAll
        lns = Split(buf, vbCrLf)
        FORCED_MODULE(file_cnt) = Replace(f.Name, "__", "")
        For file_cnt = LBound(lns) To UBound(lns)
            FORCED_XL(line_cnt, file_cnt) = lns(i)
            line_cnt = line_cnt + 1
        Next
        file_cnt = file_cnt + 1
    Next
    
    ReDim Preserve FORCED_XL(line_cnt, file_cnt)
    
End Sub

Sub AddElementsOfExclusionList()
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
    Dim i As Integer
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

Sub Main()
    Dim i As Integer
    For i = LBound(IMPORT_LIST) To UBound(IMPORT_LIST)
        ThisWorkbook.VBProject.vbcomponents.IMPORT IMPORT_PATH & "\" & IMPORT_LIST(i)
    Next
End Sub

