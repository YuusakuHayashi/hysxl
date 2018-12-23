Attribute VB_Name = "hysFiler"
Option Explicit

Function Encode(FileName, EncodingCharset)
    With CreateObject("ADODB.Stream")
        .Charset = EncodingCharset
        .Open
        .LoadFromFile FileName
        Encode = .ReadText
    End With
End Function
