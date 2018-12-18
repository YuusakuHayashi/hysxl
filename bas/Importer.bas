Attribute VB_Name = "Importer"
Option Explicit
Dim TARGET_PATH As String
Dim IMPORT_EXCLUSION_PATH As String
Dim DUSTBOX_PATH As String
Dim IMPORT_EXCLUSION_LIST As Variant
Dim IMPORT_RESERVATION_PATH As String
Dim MODULE_PATH As String

'覚書
'一つのサブルーチンでしか使用しないAPIをわざわざ関数化しない
'必要ないIF文はなるべく避け、なるべく関数などを分割する
'汎用のモジュール内の各関数は、他モジュールの関数を呼び出さない

Sub Bundler()
    MODULE_PATH = ThisWorkbook.PATH
    TARGET_PATH = MODULE_PATH & "\bas"
    IMPORT_EXCLUSION_PATH = MODULE_PATH & "\import_exclude"
    DUSTBOX_PATH = MODULE_PATH & "\dustbox"
    IMPORT_EXCLUSION_LIST = Array("Importer.bas", "PureImporter.bas")
    IMPORT_RESERVATION_PATH = MODULE_PATH & "\reserved"
    
    Call hysFolderer.Migrater(IMPORT_RESERVATION_PATH, DUSTBOX_PATH)
    Call hysFolderer.Migrater(TARGET_PATH, IMPORT_EXCLUSION_PATH, IMPORT_EXCLUSION_LIST)
    Call Main(TARGET_PATH)
    
    TARGET_PATH = vbNullString
End Sub

Sub Main(folder)

    Dim fso As Object
    Dim fs As Object
    Dim f As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.GetFolder(folder).Files
    
    Dim i As Integer
    Dim mn As String
    Dim mbn As String
    
    Dim imp_cnt As Integer
    Dim rem_cnt As Integer
    
    Call hysDebugger.Logger("importer start")
    Call hysDebugger.Logger("input = ", fs.Count)
    
    For Each f In fs
        mn = f.Name
        mbn = fso.getBaseName(mn)
        If checkBasExist(mbn) Then
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(mbn)
            rem_cnt = rem_cnt + 1
            If checkBasExist(mbn) Then
                Call hysFolderer.Migrater(TARGET_PATH, IMPORT_RESERVATION_PATH, Array(mn))
            Else
                ThisWorkbook.VBProject.VBComponents.IMPORT TARGET_PATH & "\" & mn
                imp_cnt = imp_cnt + 1
            End If
        Else
            ThisWorkbook.VBProject.VBComponents.IMPORT TARGET_PATH & "\" & mn
            imp_cnt = imp_cnt + 1
        End If
    Next
    
    Call hysDebugger.Logger("import = ", imp_cnt)
    Call hysDebugger.Logger("remove = ", rem_cnt)
    Call hysDebugger.Logger("importer end")
    
    Set fso = Nothing
    Set fs = Nothing
    Set f = Nothing
    
End Sub


'Function checkBasExist(bas As String) As Boolean
Function checkBasExist(ByVal bas As String) As Boolean
    Dim i As Integer
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = 1 Then
                If .VBComponents(i).Name = bas Then
                    checkBasExist = True
                    Call hysDebugger.Logger(bas & " -> checkBasExist hit")
                    Exit Function
                End If
            End If
        Next
        checkBasExist = False
        Call hysDebugger.Logger(bas & " -> checkBasExist no hit")
    End With
End Function
