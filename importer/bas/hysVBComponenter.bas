Attribute VB_Name = "hysVBComponenter"
Option Explicit

Function GetComponentList(CType As Integer) As String()
    Dim i As Integer
    Dim j As Integer: j = 0
    Dim cnt As Integer: cnt = 0
    Dim var() As String
    
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = CType Then
                cnt = cnt + 1
            End If
        Next
        ReDim var(cnt - 1)
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = CType Then
                var(j) = .VBComponents(i).Name
                j = j + 1
            End If
        Next
    End With
    
    GetComponentList = var
    
End Function
