#include "..\AutoItObject_Internal.au3"
#include "..\AutoItObject_Internal_ROT.au3"

Global $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")
If @error <> 0 Then
    ConsoleWriteError("Could not set AuotIt3 general COM error handler")
    Exit 1
EndIf

Func _ErrFunc($oError)
    ConsoleWrite(@ScriptName & " (" & $oError.scriptline & ") : ==> COM Error intercepted !" & @CRLF & _
        @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
        @TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
        @TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
        @TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
        @TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
        @TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
        @TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
        @TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
        @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>_ErrFunc

$oIDispatch = IDispatch()
$oIDispatch.name = "Shared IDispatch Object"
Global Const $dwID = _AOI_ROT_register($oIDispatch, "AutoIt3.COM.Test", True)
If @error <> 0 Then
    ConsoleWriteError("Could not register in the ROT"&@CRLF)
    Exit
EndIf

Opt("GUIOnEventMode", 1)
$hWnd = GUICreate("", 700, 320)
GUISetOnEvent(-3, "_MyExit")
GUICtrlCreateButton("revoke", 10, 10, 100, 30)
GUICtrlSetOnEvent(-1, "_revoke")
GUICtrlCreateButton("test", 10, 50, 100, 30)
GUICtrlSetOnEvent(-1, "_test")
GUISetState(@SW_SHOW, $HWnd)

While 1
    Sleep(10)
WEnd

Func _revoke()
    ConsoleWrite("_AOI_ROT_revoke: "&_AOI_ROT_revoke($dwID)&@CRLF)
EndFunc

Func _test()
    ShellExecute(StringRegExpReplace(@ScriptFullPath, ".\w+$", ".vbs"), "", @ScriptDir)
EndFunc

Func _MyExit()
    $oIDispatch = 0
    Exit
EndFunc
