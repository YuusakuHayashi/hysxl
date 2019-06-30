Attribute VB_Name = "Module1"
Option Explicit

Sub test()
    Call get_index("C:\Users\yuusaku.hayashi\Pictures")
End Sub

Function get_file_index(ByRef fc As Object) As String()
    Dim o As Object
    Dim f As Object
    
    Dim s() As String
    ReDim s(fc.Count)
    Dim i As Integer: i = 0
    
    For Each o In fc
        Set f = o
        s(i) = f.Name
        i = i + 1
    Next
    
    get_file_index = s
End Function

Sub get_index_re(ByRef foc As Object, y As Integer, x As Integer)
    Dim o As Object
    Dim fo As Object
    
    Dim s() As String
    
    For Each o In foc
        Set fo = o
        x = x + 1
        With fo
            Cells(y, x + 1) = .Name
            
            s = get_file_index(.Files)
            For y = LBound(s) To UBound(s)
                Cells(y + 1, x + 1) = s(y)
            Next
            
            Call get_index_re(.SubFolders, x, y)
        End With
    Next
End Sub

Sub get_index(ByVal root As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'ルートフォルダ取得
    Dim fo As Object
    Set fo = fso.GetFolder(root)
    
    
    Dim row As Integer: row = 0
    Dim col As Integer: col = 0
    
    
    '直接イテレータの各childを参照する方法が分からなかったので
    '汎用オブジェクトに格納させて、それを参照する
    Dim o As Object
    Dim s() As String
    
    With fo
        s = get_file_index(.Files)
        For row = LBound(s) To UBound(s)
            Cells(row + 1, col + 1) = s(row)
        Next
        
        Call get_index_re(.SubFolders, row, col)
    End With
    
'    function (pn, kv) {
'      var spreadsheet = hys.get_spreadsheet('1EUIFk7bX_vEE_4ibgRfEN2aTUeuTU43z6WCt9Tzalx0');
'      var sheet = hys.get_sheet(spreadsheet, 'get_index_folder@'+pn);
'      var root = hys.get_root_folder(pn);
'      var data = [];
'
'      var root_folder = hys.get_rootname(root);
'      var directory_name = root_folder;
'
'      data.push([root.getId(), directory_name]);
'
'      var folders = root.getFolders();
'      for(var i = 0; folders.hasNext(); i++) {
'        var folder = folders.next();
'        var parent_folder = folder.getName();
'        directory_name = root_folder + "/" + parent_folder;
'        data.push([folder.getId(), directory_name]);
'
'        (function get_index_folder_recursive(_f, _r) {
'          var folders = _f.getFolders();
'          if (folders.hasNext()) {
'            for(i=0; folders.hasNext(); i++) {
'              var folder = folders.next();
'              directory_name += ("/" + folder.getName());
'              data.push([folder.getId(), directory_name]);
'              get_index_folder_recursive(folder, directory_name);
'              directory_name = _r;
'            }
'          }
'        })(folder, directory_name);
'
'      }
'      sheet.clear()
'      sheet.getRange(1,1,data.length,data[0].length).setValues(data);
'    }, 'get_index_init'
'  )
End Sub
