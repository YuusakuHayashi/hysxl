Attribute VB_Name = "hysVBComponenter"
Option Explicit

Function GetComponentList(CType As Integer) As String()
    Dim i As Integer
    Dim j As Integer: j = 1
    Dim cnt As Integer
    Dim var() As String
    
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = CType Then
                cnt = cnt + 1
            End If
        Next
        ReDim var(cnt)
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = CType Then
                var(j) = .VBComponents(i).Name
                j = j + 1
            End If
        Next
    End With
    
    GetComponentList = var
    
End Function
