Dim objXL
On Error Resume Next
Set objXL = GetObject("AutoIt3.COM.Test")
If IsObject(objXL) Then
    MsgBox objXL.name, 0, "Success"
Else
    MsgBox "Error: Object could not be found.", 48, "Failure"
End If
Set objXL = Nothing
