Dim objXL
REM Set objXL = CreateObject("AutoIt.COMDemo")
REM Set objXL = GetObject(,"AutoIt.COMDemo")
On Error Resume Next
Set objXL = GetObject("AutoIt.COMDemo")
If IsObject(objXL) Then
    MsgBox(objXL.name)
Else
    MsgBox("Error: Object could not be found.")
End If
Set objXL = Nothing
