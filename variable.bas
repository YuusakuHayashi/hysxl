Attribute VB_Name = "variable"

Type struct_data
    b As Boolean
    m As String
End Type

Private Function return_sd() As struct_data
    Dim sd As struct_data
    sd.b = False
    sd.m = "message"
    return_sd = sd
End Function

Function test_f() As String
	test_f = return_sd.m
End Function
