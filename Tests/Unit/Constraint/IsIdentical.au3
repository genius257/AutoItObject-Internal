#include-once
;https://github.com/sebastianbergmann/phpunit/blob/master/src/Framework/Constraint/IsIdentical.php

Func Au3UnitConstraintIsIdentical_Evaluate($other, $description = "", $returnResult = false, $line = Null, $exspected = Null)
	Local $success = $exspected == $other

	If $returnResult Then Return $success

	If Not $success Then
;~ 		Local $f = Null

		;FIXME: look into implementing some of the original functionallity, @see https://github.com/sebastianbergmann/phpunit/blob/master/src/Framework/Constraint/IsIdentical.php#L84-L102
		Call("Au3UnitConstraintIsIdentical_Fail", $other, $description, Null, $line, $exspected)
		If @error = 0xDEAD And @extended = 0xBEEF Then Call("Au3UnitConstraintConstraint_Fail", "IsIdentical", $other, $description, Null, $line)
		If @error = 0xDEAD And @extended = 0xBEEF Then Exit MsgBox(0, "Au3Unit", "Au3UnitConstraintConstraint_Fail function is missing"&@CRLF&"Exitting") + 1
		Return SetError(@error)
	EndIf
EndFunc

Func Au3UnitConstraintIsIdentical_Fail($other, $description, $comparisonFailure, $line, $exspected)
	Local $failureDescription = Call("Au3UnitConstraintIsIdentical_FailureDescription", $other, $exspected)
	If @error = 0xDEAD And @extended = 0xBEEF Then $failureDescription = Call("Au3UnitConstraintConstraint_FailureDescription", "IsIdentical", $other)
	$failureDescription = StringFormat("Failed asserting that %s.", $failureDescription)

	Local $additionalFailureDescription = Call("Au3UnitConstraintIsIdentical_AdditionalFailureDescription", $other)
	If @error = 0xDEAD And @extended = 0xBEEF Then $additionalFailureDescription = Call("Au3UnitConstraintConstraint_AdditionalFailureDescription", $other)

	If $additionalFailureDescription Then $failureDescription &= @CRLF & $additionalFailureDescription

	If Not ($description = "") Then $failureDescription = $description & @CRLF & $failureDescription

	ConsoleWriteError($failureDescription&@CRLF&@ScriptFullPath&":"&$line&@CRLF)
	Return SetError(1)
EndFunc

Func Au3UnitConstraintIsIdentical_FailureDescription($other, $exspected)
	Local $toString =  Call("Au3UnitConstraintIsIdentical_ToString", $exspected)
	If @error = 0xDEAD And @extended = 0xBEEF Then $toString = Call("Au3UnitConstraintConstraint_ToString", $other)
	Return Au3ExporterExporter_Export($other) & " " & $toString
EndFunc

Func Au3UnitConstraintIsIdentical_ToString($value)
	;If IsObj($value);TODO: https://github.com/sebastianbergmann/phpunit/blob/master/src/Framework/Constraint/IsIdentical.php#L115

	Return 'is identical to ' & Au3ExporterExporter_Export($value)
EndFunc