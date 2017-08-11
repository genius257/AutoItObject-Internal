#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         genius257

#ce ----------------------------------------------------------------------------
#include "..\AutoItObject_Internal.au3"

$AutoItError = ObjEvent("AutoIt.Error", "ErrFunc") ; Install a custom error handler
Func ErrFunc($oError)
	ConsoleWrite("!>COM Error !"&@CRLF&"!>"&@TAB&"Number: "&Hex($oError.Number,8)&@CRLF&"!>"&@TAB&"Windescription: "&StringRegExpReplace($oError.windescription,"\R$","")&@CRLF&"!>"&@TAB&"Source: "&$oError.source&@CRLF&"!>"&@TAB&"Description: "&$oError.description&@CRLF&"!>"&@TAB&"Helpfile: "&$oError.helpfile&@CRLF&"!>"&@TAB&"Helpcontext: "&$oError.helpcontext&@CRLF&"!>"&@TAB&"Lastdllerror: "&$oError.lastdllerror&@CRLF&"!>"&@TAB&"Scriptline: "&$oError.scriptline&@CRLF)
EndFunc ;==>ErrFunc
;we need the error handler to prevent crash when IDispatch raises exceptions and the like

$oIDispatch = IDispatch()

$oIDispatch.string = "123";optional to define property first

$oIDispatch.__defineGetter("string", Getter1)
$oIDispatch.__defineSetter("string", Setter1)

MsgBox(0, "", $oIDispatch.string); invoke getter
$oIDispatch.string = "321"; invoke setter (this will fail)
If @error<>0 Then MsgBox(0, "", "setter set @error<>0"&@CRLF&"see console for more information")

;Replace accessors
$oIDispatch.__defineGetter("string", Getter2)
$oIDispatch.__defineSetter("string", Setter2)

MsgBox(0, "", $oIDispatch.string); invoke getter as before
MsgBox(0, "", $oIDispatch.string('"'));invoke getter with parameters. Try and replace the string if you want
$oIDispatch.string = "321"; invoke setter
MsgBox(0, "", $oIDispatch.string);get the new value after it has been manipulated by the getter

Func Getter1()
	Return "static string return"
EndFunc

Func Getter2($AccessorObject)
	If $AccessorObject.arguments.length=0 Then Return "Simon says: " & $AccessorObject.val
	If $AccessorObject.arguments.length>1 Then Return SetError(1, 1, 0)
	Return $AccessorObject.arguments.values[0]&$AccessorObject.val&$AccessorObject.arguments.values[0]
EndFunc

Func Setter1()
	Return SetError(1, 1, 0);set @error<>0 to invoke exception
EndFunc

Func Setter2($AccessorObject)
	If Not IsString($AccessorObject.val) Then Return SetError(1, 1, 0);in this example we only allow the value to be a string
	$AccessorObject.val = $AccessorObject.ret
EndFunc