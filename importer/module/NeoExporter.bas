Attribute VB_Name = "NeoExporter"
Option Explicit

Dim EXPORT_PATH As String
Dim MODULE_PATH As String
Dim EXPORT_LIST
Dim EXCLUSION_LIST
Dim IGNORE_FILE As String

Sub Bundler()
    MODULE_PATH = ThisWorkbook.path & "\importer"
    IGNORE_FILE = MODULE_PATH & "\.exportignore"
    
    EXCLUSION_LIST = Array("NeoImporter")
    EXCLUSION_LIST = hysStr.AddElementsOfList(hysStr.GetListOfFileLines(IGNORE_FILE), EXCLUSION_LIST)

EXPORT_EXCLUSION:
    EXPORT_PATH = MODULE_PATH & "\exclude"
    Call ExportExclusionList
    
EXPORT_BAS:
    EXPORT_PATH = MODULE_PATH & "\module"
    EXPORT_LIST = hysVBComponenter.GetComponentsList(1)
    EXPORT_LIST = hysStr.ExcludeElementsOfList(EXCLUSION_LIST, EXPORT_LIST)
    Call Main
    Call RemoveComponents

EXPORT_CLS:
    EXPORT_LIST = hysVBComponenter.GetComponentsList(2)
    Call Main
    Call RemoveComponents
    
EXPORT_FRM:
    EXPORT_LIST = hysVBComponenter.GetComponentsList(3)
    Call Main
    Call RemoveComponents
    
End Sub

Sub ExportExclusionList()
    Dim i As Integer
    On Error Resume Next
    For i = LBound(EXCLUSION_LIST) To UBound(EXCLUSION_LIST)
        ThisWorkbook.VBProject.vbcomponents(EXCLUSION_LIST(i)).EXPORT EXPORT_PATH & "\" & EXCLUSION_LIST(i) & ".bas"
    Next
    On Error GoTo 0
End Sub

Sub Main()
    Dim i As Integer: i = 1
    With ThisWorkbook.VBProject
        For i = LBound(EXPORT_LIST) To UBound(EXPORT_LIST)
            Select Case .vbcomponents(EXPORT_LIST(i)).Type
                Case 1
                    .vbcomponents(EXPORT_LIST(i)).EXPORT EXPORT_PATH & "\" & EXPORT_LIST(i) & ".bas"
                Case 2
                    .vbcomponents(EXPORT_LIST(i)).EXPORT EXPORT_PATH & "\" & EXPORT_LIST(i) & ".cls"
                Case 3
                    .vbcomponents(EXPORT_LIST(i)).EXPORT EXPORT_PATH & "\" & EXPORT_LIST(i) & ".frm"
                Case 11
                    .vbcomponents(EXPORT_LIST(i)).EXPORT EXPORT_PATH & "\" & EXPORT_LIST(i) & ".axd"
                Case 100
                    .vbcomponents(EXPORT_LIST(i)).EXPORT EXPORT_PATH & "\" & EXPORT_LIST(i) & ".cls"
            End Select
        Next
    End With
End Sub

Sub RemoveComponents()
    Dim i As Integer: i = 1
    With ThisWorkbook.VBProject
        For i = LBound(EXPORT_LIST) To UBound(EXPORT_LIST)
            .vbcomponents.Remove ThisWorkbook.VBProject.vbcomponents(EXPORT_LIST(i))
        Next
    End With
End Sub
