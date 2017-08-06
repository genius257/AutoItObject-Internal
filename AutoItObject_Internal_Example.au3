#include <WINAPI.au3>
#include "AutoItObject_Internal.au3"

;~ #AutoIt3Wrapper_UseX64=Y
;~ #AutoIt3Wrapper_Run_Au3Check=N

$AutoItError = ObjEvent("AutoIt.Error", "ErrFunc") ; Install a custom error handler

Func ErrFunc($oError)
	ConsoleWrite("!>COM Error !"&@CRLF&"!>"&@TAB&"Number: "&Hex($oError.Number,8)&@CRLF&"!>"&@TAB&"Windescription: "&StringRegExpReplace($oError.windescription,"\R$","")&@CRLF&"!>"&@TAB&"Source: "&$oError.source&@CRLF&"!>"&@TAB&"Description: "&$oError.description&@CRLF&"!>"&@TAB&"Helpfile: "&$oError.helpfile&@CRLF&"!>"&@TAB&"Helpcontext: "&$oError.helpcontext&@CRLF&"!>"&@TAB&"Lastdllerror: "&$oError.lastdllerror&@CRLF&"!>"&@TAB&"Scriptline: "&$oError.scriptline&@CRLF)
EndFunc   ;==>ErrFunc

$oItem = IDispatch()

$oItem.a = "a"
$oItem.b = "b"
$oItem.c = "c"
$oItem.z = "z"

$oItem.__unset("z")

$oItem.__defineGetter('a',Getter)
$oItem.__defineSetter('a',Setter)

$oItem.ab = $oItem.a&$oItem.b
$oItem.abc = $oItem.a&$oItem.b&$oItem.c
$oItem.abc = $oItem.abc&$oItem.abc

$oItem.f = Test
Execute("($oItem.f)('a')");we use execute, due to Au3Check bug. For more, see: https://www.autoitscript.com/trac/autoit/ticket/3560
MsgBox(0, "", "a: "&$oItem.a&@CRLF&"b: "&$oItem.b&@CRLF&"c: "&$oItem.c&@CRLF&"ab: "&$oItem.ab&@CRLF&"abc: "&$oItem.abc)
$oItem.f = MsgBox
Call($oItem.f,0,"","test")

$oItem.__lock();locking the object

$oItem.z = "z"; this will fail because of lock, see console. Also @error is set to non zero

Func Test($string)
	MsgBox(0, "UserFunction: Test", $string)
EndFunc

Func Getter($oThis)
	Return $oThis.val&$oThis.parent.b
EndFunc

Func Setter($oThis)
	$oThis.val=$oThis.ret
EndFunc