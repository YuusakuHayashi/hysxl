Attribute VB_Name = "hysLister"
Option Explicit

Function GetIndexOfMember(el, list) As Integer
    Dim i As Integer
    For i = LBound(list) To UBound(list)
        If el = list(i) Then
            GetIndexOfMember = i
            Exit Function
        End If
    Next
End Function
