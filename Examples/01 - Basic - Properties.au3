#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         genius257

#ce ----------------------------------------------------------------------------

#include "..\AutoItObject_Internal.au3"

$oIDispatch = IDispatch()

$oIDispatch.a = "a"
$oIDispatch.A = "A"
MsgBox(0, @ScriptName, "Properties are case sensetive"&@CRLF&"$oIDispatch.a: "&$oIDispatch.a&@CRLF&"$oIDispatch.A: "&$oIDispatch.A)
$oIDispatch.__unset("a")
$oIDispatch.__unset("A")

Global $aArray = [1,2,3]
$oIDispatch.array = $aArray

$oIDispatch.binary = Binary(123)

$oIDispatch.bool = True

$oIDispatch.dllStruct = DllStructCreate("BYTE")

$oIDispatch.function = MsgBox

$oIDispatch.float = 12.3

$oIDispatch.int = 123

$oIDispatch.keyword = Default

$oIDispatch.object = ObjCreate("ScriptControl")

$oIDispatch.pointer = ptr(123);ptr is converted to int. Not much i can do about that.

$oIDispatch.string = "123"

For $key In $oIDispatch.__keys
	ConsoleWrite("$oIDispatch."&$key&" : "&VarGetType(Execute("$oIDispatch."&$key))&@CRLF)
Next