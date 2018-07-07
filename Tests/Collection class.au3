#include "..\AutoItObject_Internal.au3"

Func Collection($items)
	Local $class = IDispatch()

	$class.__name = "Collection"
	$class.__defineSetter("__name", PrivateProperty)
	$class.__defineGetter("getArrayableItems", Collection_getArrayableItems)
	$class.items = $class.getArrayableItems($items)
	$class.__defineSetter("items", PrivateProperty)
	$class.__defineGetter("get", Collection_Get)

	Return $class
EndFunc

Func Collection_Get($this)
	Local $arguments = $this.arguments
	Local $length = $arguments.length
	If $length<1 Or $length>2 Then Return SetError(1)
	Local $items = $arguments.values
	Local $path = StringSplit($items[0], ".", 2)
	Local $return = $this.parent.items
	Local $value
	For $segment In $path
		$value = Execute("$return["&$segment&"]")
		If @error==0 Then
			$return = $value
			ContinueLoop
		EndIf
		$value = Execute("$return['"&$segment&"']")
		If @error==0 Then
			$return = $value
			ContinueLoop
		EndIf
		If IsObj($items) And Execute("$items.__name") == "Collection" Then
			$value = $return.items.get($segment)
			If @error==0 Then
				$return = $value
				ContinueLoop
			EndIf
		Else
			$value = Execute("$return."&$segment)
			If @error==0 Then
				$return = $value
				ContinueLoop
			EndIf
		EndIf
		Return $length==2?$items[1]:Null;SetError(2, 0, $length==2?$items[1]:Null)
	Next
	Return $return
EndFunc

Func Collection_getArrayableItems($this)
	Local $arguments = $this.arguments
	If $arguments.length <> 1 Then Return SetError(1)
	Local $items = $arguments.values[0]
	If IsArray($items) Then
		Return $items
	ElseIf IsObj($items) And Execute("$items.__name") == "Collection" Then
		Return $items.items
	Else
		Local $return = [$items]
		Return $return
	EndIf
EndFunc

#cs
$oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")
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
EndFunc
#ce

$a = "success"
$collection = Collection($a)
MsgBox(0, "", $collection.get("0.0", "failure"));FIXME: currently does not support multi-dimentional
Exit