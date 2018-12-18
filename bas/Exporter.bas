Attribute VB_Name = "Exporter"
Option Explicit
Dim TARGET_PATH As String
Dim EXPORT_EXCLUSION_PATH As String
Dim EXPORT_EXCLUSION_LIST As Variant

Sub Bundler()
    TARGET_PATH = ThisWorkbook.PATH & "\bas"
    EXPORT_EXCLUSION_PATH = ThisWorkbook.PATH & "\export_exclude"
    EXPORT_EXCLUSION_LIST = Array()
    Call Main(TARGET_PATH, EXPORT_EXCLUSION_LIST)
End Sub

Sub Main(folder, ex)
    
    Dim i As Integer
    Dim j As Integer
    Dim mn As String
    Dim exp_cnt As Integer
    
    Call hysDebugger.Logger("exporter start")
    Call hysDebugger.Logger("input = ", ThisWorkbook.VBProject.VBComponents.Count)
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = 1 Then
                mn = .VBComponents(i).Name
                If checkExclude(mn, ex) = False Then
                    .VBComponents(i).EXPORT folder & "\" & mn & ".bas"
                    Call hysDebugger.Logger("export " & mn)
                    exp_cnt = exp_cnt + 1
                End If
            End If
        Next
    End With
    Call hysDebugger.Logger("export = ", exp_cnt)
    Call hysDebugger.Logger("exporter end")
End Sub
