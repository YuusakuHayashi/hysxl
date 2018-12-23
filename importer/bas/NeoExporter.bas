Attribute VB_Name = "NeoExporter"
Option Explicit

Dim EXPORT_PATH As String
Dim MODULE_PATH As String
Dim EXPORT_LIST() As String
Dim EXCLUSION_LIST As Variant
Dim IGNORE_FILE As String

Sub Bundler()
    MODULE_PATH = ThisWorkbook.PATH & "\importer"
    EXPORT_PATH = MODULE_PATH & "\exclude"
    IGNORE_FILE = MODULE_PATH & "\.exportignore"
    
    EXPORT_LIST = hysVBComponenter.GetComponentList(1)
    
    EXCLUSION_LIST = Array("NeoImporter")
    EXCLUSION_LIST = hysStr.AddElementsOfList(hysStr.GetListOfFileLines(IGNORE_FILE), EXCLUSION_LIST)
    EXPORT_LIST = hysStr.ExcludeElementsOfList(EXCLUSION_LIST, EXPORT_LIST)
    
    Call ExportReserve
    
    EXPORT_PATH = MODULE_PATH & "\bas"
    Call Main
End Sub

Sub ExportReserve()
    Dim i As Integer
    On Error Resume Next
    For i = LBound(EXCLUSION_LIST) To UBound(EXCLUSION_LIST)
        ThisWorkbook.VBProject.VBComponents(EXCLUSION_LIST(i)).EXPORT EXPORT_PATH & "\" & EXCLUSION_LIST(i) & ".bas"
    Next
    On Error GoTo 0
End Sub

Sub Main()
    Dim i As Integer: i = 1
    With ThisWorkbook.VBProject
        'For i = LBound(EXPORT_LIST) = 1 To UBound(EXPORT_LIST)
        For i = LBound(EXPORT_LIST) To UBound(EXPORT_LIST)    'ÉäÉeÉâÉãÇæÇ∆ãCï™Ç™ó«Ç≠Ç»Ç¢ÇÃÇ≈óvèCê≥
            .VBComponents(EXPORT_LIST(i)).EXPORT EXPORT_PATH & "\" & EXPORT_LIST(i) & ".bas"
            .VBComponents.Remove ThisWorkbook.VBProject.VBComponents(EXPORT_LIST(i))
        Next
    End With
End Sub
