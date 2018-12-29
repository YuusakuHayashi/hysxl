Attribute VB_Name = "hysRanger"
Option Explicit

Sub inputToVanila(rgn, str As String)
    Dim i As Integer
    Dim j As Integer
    For i = 3 To 141
        For j = 6 To 12
            'Cells(i, j).Select
            If Cells(i, j).Text = "" Then
                Cells(i, j) = str
            End If
        Next
    Next
End Sub

Sub test()
    Call inputToVanila(Range(Cells(3, 6), Cells(141, 12)), "Å~")
End Sub
