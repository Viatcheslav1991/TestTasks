
'Sub test()
'   Debug.Print ������������(27)
'End Sub
 
Function ������������(ByVal col As Long) As String
   On Error Resume Next
   ������������ = Application.ConvertFormula("r1c" & col, xlR1C1, xlA1)
   ������������ = Replace(Replace(Mid(������������, 2), "$", ""), "1", "")
End Function
