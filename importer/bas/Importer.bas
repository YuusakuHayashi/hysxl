Attribute VB_Name = "Importer"
'Option Explicit
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Dim TARGET_PATH As String
'Dim IMPORT_EXCLUSION_PATH As String
'Dim DUSTBOX_PATH As String
'Dim IMPORT_EXCLUSION_LIST As Variant
'Dim IMPORT_RESERVATION_PATH As String
'Dim MODULE_PATH As String
'
''要件
''VBAモジュールを外部ファイルで保存して、
''excel起動時にそれを自動的にインポート出来ることが最もしたいこと。
''ただ、excelを開いて、内蔵VBEのデバッグ機能とかも使って直接編集もしたいので、
''excel終了時に、VBAモジュールをエクスポートできるとなおよし。
'
''HISTORY
''- 2018/12/08 一区切り。ひな形にしてもいいかな？問題が山積み・・・
''- 2018/12/15 多分、今後はHISTORY等も、コード中に書いていくと思います。
''- 2018/12/15 コードを若干整理。本質的な問題にはまだ触れられていない
''  - 汎用的な関数は、別モジュールに分けて、それを`Call`していたが、
''               そうすると実行中はそれらモジュールを`remove`出来ず、実行後に適用される事が分かった。
''               `remove`後、`import`も行うので、hoge(1)のような名前でインポートしてしまう
''- 2018/12/15 主機能をImporterモジュール、PureImporterモジュール、Exporterモジュールに分割
''- 2018/12/20 無理が多くなってきたので、大幅に見直し予定。
''                               このモジュールは残して、別の名前モジュールを作成する予定
'
'Sub Bundler()
'    MODULE_PATH = ThisWorkbook.PATH
'    TARGET_PATH = MODULE_PATH & "\bas"
'    IMPORT_EXCLUSION_PATH = MODULE_PATH & "\import_exclude"
'    DUSTBOX_PATH = MODULE_PATH & "\dustbox"
'    IMPORT_EXCLUSION_LIST = Array()
'    IMPORT_RESERVATION_PATH = MODULE_PATH & "\reserved"
'
'    Call hysFolderer.Migrater(IMPORT_RESERVATION_PATH, DUSTBOX_PATH)
'    Call hysFolderer.Migrater(TARGET_PATH, IMPORT_EXCLUSION_PATH, IMPORT_EXCLUSION_LIST)
'    Call Main(TARGET_PATH)
'
'    TARGET_PATH = vbNullString
'End Sub
'
'Sub Main(folder)
'
'    Dim fso As Object
'    Dim fs As Object
'    Dim f As Object
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set fs = fso.GetFolder(folder).files
'
'    Dim i As Integer
'    Dim mn As String
'    Dim mbn As String
'
'    Dim imp_cnt As Integer
'    Dim rem_cnt As Integer
'
'    Call hysDebugger.Logger("importer start")
'    Call hysDebugger.Logger("input = ", fs.Count)
'
'    For Each f In fs
'        mn = f.Name
'        mbn = fso.getBaseName(mn)
'        If checkBasExist(mbn) Then
'            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(mbn)
'            Sleep 200
'            rem_cnt = rem_cnt + 1
'            If checkBasExist(mbn) Then
'                Call hysFolderer.Migrater(TARGET_PATH, IMPORT_RESERVATION_PATH, Array(mn))
'                Call hysDebugger.Logger("This Module is Reserved = " & mbn)
'            Else
'                ThisWorkbook.VBProject.VBComponents.IMPORT TARGET_PATH & "\" & mn
'                Sleep 200
'                Call hysDebugger.Logger("This Module is Removed & Imported = " & mbn)
'                imp_cnt = imp_cnt + 1
'            End If
'        Else
'            ThisWorkbook.VBProject.VBComponents.IMPORT TARGET_PATH & "\" & mn
'            Sleep 200
'            Call hysDebugger.Logger("This Module is Imported = " & mbn)
'            imp_cnt = imp_cnt + 1
'        End If
'    Next
'
'    Call hysDebugger.Logger("import = ", imp_cnt)
'    Call hysDebugger.Logger("remove = ", rem_cnt)
'    Call hysDebugger.Logger("importer end")
'
'    Set fso = Nothing
'    Set fs = Nothing
'    Set f = Nothing
'
'End Sub
'
'
''Function checkBasExist(bas As String) As Boolean
'Function checkBasExist(ByVal bas As String) As Boolean
'    Dim i As Integer
'    With ThisWorkbook.VBProject
'        For i = 1 To .VBComponents.Count
'            If .VBComponents(i).Type = 1 Then
'                If .VBComponents(i).Name = bas Then
'                    checkBasExist = True
'                    Call hysDebugger.Logger(bas & " -> checkBasExist hit")
'                    Exit Function
'                End If
'            End If
'        Next
'        checkBasExist = False
'        Call hysDebugger.Logger(bas & " -> checkBasExist no hit")
'    End With
'End Function
