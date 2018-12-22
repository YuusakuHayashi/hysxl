Attribute VB_Name = "NeoImporter"
Option Explicit

Dim IMPORT_PATH As String
Dim MODULE_PATH As String
Dim EXCLUSION_LIST As Variant

Sub Bundler()
    MODULE_PATH = ThisWorkbook.PATH & "\importer"
    IMPORT_PATH = MODULE_PATH & "\bas"
    EXCLUSION_LIST = Array("NeoImporter")
    
    Call Main
    
    IMPORT_PATH = vbNullString
End Sub

Sub Main()
    Dim fso As Object
    Dim fs As Object
    Dim f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.GetFolder(IMPORT_PATH).files
    Dim i As Integer
    Dim mn As String

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
