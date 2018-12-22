Attribute VB_Name = "NeoExporter"
Option Explicit

Dim EXPORT_PATH As String
Dim MODULE_PATH As String
Dim MODULE_LIST() As String
Dim EXCLUSION_LIST As Variant

Sub Bundler()
    MODULE_PATH = ThisWorkbook.PATH & "\importer"
    MODULE_LIST = hysVBComponenter.GetComponentList(1)
    
    EXPORT_PATH = MODULE_PATH & "\exclude"
    EXCLUSION_LIST = Array("NeoImporter")
    Call ExportReserve
    
    EXPORT_PATH = MODULE_PATH & "\bas"
    MODULE_LIST = hysStr.RemoveElement(hysLister.GetIndexOfMember("NeoImporter", MODULE_LIST), MODULE_LIST)
    Call Main
End Sub

Sub ExportReserve()
    Dim i As Integer
    For i = LBound(EXCLUSION_LIST) To UBound(EXCLUSION_LIST)
        ThisWorkbook.VBProject.VBComponents(EXCLUSION_LIST(i)).EXPORT EXPORT_PATH & "\" & EXCLUSION_LIST(i) & ".bas"
    Next
End Sub

Sub Main()
    Dim i As Integer: i = 1
    With ThisWorkbook.VBProject
        'For i = LBound(MODULE_LIST) = 1 To UBound(MODULE_LIST)
        For i = 1 To UBound(MODULE_LIST)    'ÉäÉeÉâÉãÇæÇ∆ãCï™Ç™ó«Ç≠Ç»Ç¢ÇÃÇ≈óvèCê≥
            Debug.Print MODULE_LIST(i)
            .VBComponents(MODULE_LIST(i)).EXPORT EXPORT_PATH & "\" & MODULE_LIST(i) & ".bas"
            .VBComponents.Remove ThisWorkbook.VBProject.VBComponents(MODULE_LIST(i))
        Next
    End With
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
