Dim objXL
On Error Resume Next
Set objXL = GetObject("AOI.Debug")
If IsObject(objXL) Then
    MsgBox objXL.name, 0, "Success"
Else
    MsgBox "Error: Object could not be found.", 48, "Failure"
End If
Set objXL = Nothing
