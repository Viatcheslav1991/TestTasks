
'Sub test()
'   Debug.Print БукваСтолбца(27)
'End Sub
 
Function БукваСтолбца(ByVal col As Long) As String
   On Error Resume Next
   БукваСтолбца = Application.ConvertFormula("r1c" & col, xlR1C1, xlA1)
   БукваСтолбца = Replace(Replace(Mid(БукваСтолбца, 2), "$", ""), "1", "")
End Function
