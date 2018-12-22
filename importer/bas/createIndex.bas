Attribute VB_Name = "createIndex"
Option Explicit

Const CONFIGURATION = "\\Coral\個人情報-林祐\createIndex\config"

Function checkSheetName()

    If ActiveSheet.CodeName = "sh_createIndex" Then
        checkSheetName = True
    Else
        checkSheetName = False
    End If
    
End Function

Function getFilePath()  'ファイルパスを取得する関数
    Dim fp As String    'file path
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        getFilePath = .SelectedItems(1)
    End With
    
End Function

Function getConfig() As Variant()    '設定ファイルを読み込む関数
    Dim fso, fd, fs, f, ts As Object
    Dim l(4) As Variant
    
    'ChDir ThisWorkbook.Path
    'ChDrive ThisWorkbook.Path
    
    Set fso = CreateObject("Scripting.filesystemObject")
    
    Set f = fso.GetFile(CONFIGURATION)
    Set ts = f.OpenAsTextStream(1, -2)
    
    With ts
        l(0) = CInt(.ReadLine)      'インデックス作成フォルダの表示行
        l(1) = CInt(.ReadLine)      'インデックス作成フォルダの表示列
        l(2) = CInt(.ReadLine)      'インデックス作成の起点の行
        l(3) = CInt(.ReadLine)      'インデックス作成の起点の列
    End With
    
    getConfig = l()
    
End Function

Sub writefilepath() 'セルにファイルパスを書き込む関数

    Dim fp As String
    fp = getFilePath
    
    Dim cfg As Variant
    cfg = getConfig
    
    Cells(CInt(cfg(0)), CInt(cfg(1))) = fp
    
End Sub

Sub Main()

read_config:            '設定ファイルの中身を変数に代入
    Dim cfg() As Variant
    cfg = getConfig
    Dim i As Integer: i = cfg(2)
    Dim j As Integer: j = cfg(3)
    
initializing:           'セルクリア
    With Cells(i, j).Offset(-1)
        .CurrentRegion.Clear
        .Value = "▼"
    End With
       
read_input:             '入力ファイル読み込み
    Dim fp As String
    fp = Cells(CInt(cfg(0)), CInt(cfg(1))).Text
    
    Dim fso, fd, fs, f As Object
    Set fso = CreateObject("Scripting.filesystemObject")
    Set fd = fso.GetFolder(fp)
    Set fs = fd.files
       
create_index:           'インデックス作成
    Dim hl As Hyperlink
    
    For Each f In fs
        If CStr(fso.GetExtensionName(f)) = "db" Then
            '.dbはインデックスを作成しない
        Else
            Set hl = ActiveSheet.Hyperlinks.Add( _
                Anchor:=Cells(i, j), _
                Address:=fp + "\" + f.Name, _
                TextToDisplay:=f.Name _
            )
            i = i + 1
        End If
    Next
    
End Sub
