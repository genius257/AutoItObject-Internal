Dim objXL
REM Set objXL = CreateObject("AutoIt.COMDemo")
REM Set objXL = GetObject(,"AutoIt.COMDemo")
Set objXL = GetObject("AutoIt.COMDemo")
MsgBox(objXL.name)
Set objXL = Nothing