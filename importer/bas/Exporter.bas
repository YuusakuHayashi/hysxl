Attribute VB_Name = "Exporter"
'Option Explicit
'Dim TARGET_PATH As String
'Dim EXPORT_EXCLUSION_PATH As String
'Dim EXPORT_EXCLUSION_LIST As Variant
'
'Sub Bundler()
'    MODULE_PATH = ThisWorkbook.PATH & "\importer"
'    EXPORT_PATH = MODULE_PATH & "\bas"
'    EXCLUSION_LIST = Array("NeoImporter.bas")
'
'    Call Main
'
'    IMPORT_PATH = vbNullString
'End Sub
'
'Sub Main()
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim mn As String
'
'    With ThisWorkbook.VBProject
'        For i = 1 To .VBComponents.Count
'            If .VBComponents(i).Type = 1 Then
'                mn = .VBComponents(i).Name
'                If Not checkExclude(mn, ex) = False Then
'                    .VBComponents(i).EXPORT EXPORT_PATH & "\" & mn & ".bas"
'                    Call hysDebugger.Logger("export " & mn)
'                    exp_cnt = exp_cnt + 1
'                End If
'            End If
'        Next
'    End With
'    Call hysDebugger.Logger("export = ", exp_cnt)
'    Call hysDebugger.Logger("exporter end")
'End Sub
'
'Function checkExclude(module As String, list) As Boolean
'    Dim i As Integer
'    For i = LBound(list) To UBound(list)
'        If module = list(i) Then
'            checkExclude = True
'            Exit Function
'        End If
'    Next
'    checkExclude = False
'End Function
