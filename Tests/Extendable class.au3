#include "..\AutoItObject_Internal.au3"

Func Class($extension=Null)
	Local Static $extensions[0]

	If @NumParams = 0 Then
		Local $class = IDispatch()
		For $extension In $extensions
			Call($extension, $class)
		Next
		$class.__lock()
		Return $class
	Else
		Local $count = UBound($extensions)
		ReDim $extensions[$count+1]
		$extensions[$count] = $extension
	EndIf
EndFunc

Func ClassExtension01($class)
	$class.a = "test1"
	$class.b = "test2"
	$class.c = 200
EndFunc

Class(ClassExtension01)
$class = Class()
MsgBox(0, "", $class.a)