#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         genius257

#ce ----------------------------------------------------------------------------
#include "..\AutoItObject_Internal.au3"

$IDispatch = IDispatch()
$IDispatch.__defineGetter("a", _MsgBox)

$oSC = ObjCreate("ScriptControl")
$oSC.language = "JScript"

#AutoIt3Wrapper_Run_Au3Check=N;need this line for now
($oSC.Eval("Function('e','return e.a;')"))($IDispatch);AutoIt calls JavaScript Function that invokes property a's getter

Func _MsgBox()
    Return MsgBox(0, "_MsgBox", "test")
EndFunc