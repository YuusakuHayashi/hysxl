Attribute VB_Name = "NeoExporter"
Option Explicit

Dim EXPORT_PATH As String
Dim MODULE_PATH As String
Dim MODULE_LIST() As String
Dim EXCLUSION_LIST As Variant

Sub Bundler()
    MODULE_PATH = ThisWorkbook.PATH & "\importer"
    EXPORT_PATH = MODULE_PATH & "\bas"
    MODULE_LIST = hysVBComponenter.GetComponentList(1)
    MODULE_LIST = _
        hysStr.RemoveElement( _
            hysLister.GetIndexOfMember("NeoImporter", MODULE_LIST), _
            MODULE_LIST _
        )
    Call Main
End Sub

Sub Main()
    Dim i As Integer: i = 1
    Dim max As Integer
    Dim mn As String: mn = "hoge"

    With ThisWorkbook.VBProject
        For i = 1 To UBound(MODULE_LIST) - 1
            Debug.Print MODULE_LIST(i)
            .VBComponents(MODULE_LIST(i)).EXPORT EXPORT_PATH & "\" & MODULE_LIST(i) & ".bas"
            .VBComponents.Remove ThisWorkbook.VBProject.VBComponents(MODULE_LIST(i))
        Next
    End With
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
