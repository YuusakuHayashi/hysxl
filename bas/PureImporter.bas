Attribute VB_Name = "PureImporter"
Option Explicit
Dim IMPORT_RESERVATION_PATH

Sub Main()
    
    IMPORT_RESERVATION_PATH = ThisWorkbook.PATH & "\reserved"
    Dim fso As Object
    Dim fs As Object
    Dim f As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.GetFolder(IMPORT_RESERVATION_PATH).Files
    
    Dim i As Integer
    Dim mn As String
    Dim mbn As String
    
    Dim imp_cnt As Integer
    Dim rem_cnt As Integer
    
    For Each f In fs
        mn = f.Name
        mbn = fso.getBaseName(mn)
        imp_cnt = imp_cnt + 1
        ThisWorkbook.VBProject.VBComponents.IMPORT IMPORT_RESERVATION_PATH & "\" & mn
        imp_cnt = imp_cnt + 1
    Next
    
    Set fso = Nothing
    Set fs = Nothing
    Set f = Nothing
End Sub
