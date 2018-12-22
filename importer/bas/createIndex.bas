Attribute VB_Name = "createIndex"
Option Explicit

Const CONFIGURATION = "\\Coral\�l���-�їS\createIndex\config"

Function checkSheetName()

    If ActiveSheet.CodeName = "sh_createIndex" Then
        checkSheetName = True
    Else
        checkSheetName = False
    End If
    
End Function

Function getFilePath()  '�t�@�C���p�X���擾����֐�
    Dim fp As String    'file path
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        getFilePath = .SelectedItems(1)
    End With
    
End Function

Function getConfig() As Variant()    '�ݒ�t�@�C����ǂݍ��ފ֐�
    Dim fso, fd, fs, f, ts As Object
    Dim l(4) As Variant
    
    'ChDir ThisWorkbook.Path
    'ChDrive ThisWorkbook.Path
    
    Set fso = CreateObject("Scripting.filesystemObject")
    
    Set f = fso.GetFile(CONFIGURATION)
    Set ts = f.OpenAsTextStream(1, -2)
    
    With ts
        l(0) = CInt(.ReadLine)      '�C���f�b�N�X�쐬�t�H���_�̕\���s
        l(1) = CInt(.ReadLine)      '�C���f�b�N�X�쐬�t�H���_�̕\����
        l(2) = CInt(.ReadLine)      '�C���f�b�N�X�쐬�̋N�_�̍s
        l(3) = CInt(.ReadLine)      '�C���f�b�N�X�쐬�̋N�_�̗�
    End With
    
    getConfig = l()
    
End Function

Sub writefilepath() '�Z���Ƀt�@�C���p�X���������ފ֐�

    Dim fp As String
    fp = getFilePath
    
    Dim cfg As Variant
    cfg = getConfig
    
    Cells(CInt(cfg(0)), CInt(cfg(1))) = fp
    
End Sub

Sub Main()

read_config:            '�ݒ�t�@�C���̒��g��ϐ��ɑ��
    Dim cfg() As Variant
    cfg = getConfig
    Dim i As Integer: i = cfg(2)
    Dim j As Integer: j = cfg(3)
    
initializing:           '�Z���N���A
    With Cells(i, j).Offset(-1)
        .CurrentRegion.Clear
        .Value = "��"
    End With
       
read_input:             '���̓t�@�C���ǂݍ���
    Dim fp As String
    fp = Cells(CInt(cfg(0)), CInt(cfg(1))).Text
    
    Dim fso, fd, fs, f As Object
    Set fso = CreateObject("Scripting.filesystemObject")
    Set fd = fso.GetFolder(fp)
    Set fs = fd.files
       
create_index:           '�C���f�b�N�X�쐬
    Dim hl As Hyperlink
    
    For Each f In fs
        If CStr(fso.GetExtensionName(f)) = "db" Then
            '.db�̓C���f�b�N�X���쐬���Ȃ�
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
