Global Const $Au3UnitConstraintCount = 1 ;https://github.com/sebastianbergmann/phpunit/blob/7.2/src/Framework/Constraint/Constraint.php#L72

Func Au3UnitConstraintConstraint_Evaluate($constraint, $other, $description = "", $returnResult = false, $line = Null, $passedToContraint = Null)
	Local $success = False

	Local $matches = Call("Au3UnitConstraint" & $constraint & "_Matches", $other, $passedToContraint)
	If @error = 0xDEAD And @extended = 0xBEEF Then $matches = Call("Au3UnitConstraint" & $constraint & "_Matches", $other)
	If @error = 0xDEAD And @extended = 0xBEEF Then $matches = Call("Au3UnitConstraintConstraint_Matches", $other)
	If $matches Then $success = True

	If $returnResult Then Return $success

	If Not $success Then
		Call("Au3UnitConstraint" & $constraint & "_Fail", $other, $description, Null, $line)
		If @error = 0xDEAD And @extended = 0xBEEF Then Call("Au3UnitConstraintConstraint_Fail", $constraint, $other, $description, Null, $line, $passedToContraint)
		If @error = 0xDEAD And @extended = 0xBEEF Then Exit MsgBox(0, "Au3Unit", "Au3UnitConstraintConstraint_Fail function is missing"&@CRLF&"Exitting") + 1
		Return SetError(@error)
	EndIf
EndFunc

Func Au3UnitConstraintConstraint_Fail($constraint, $other, $description, $comparisonFailure = Null, $line = Null, $passedToContraint = Null)
	Local $failureDescription = Call("Au3UnitConstraint" & $constraint & "_FailureDescription", $other, $passedToContraint)
	If @error = 0xDEAD And @extended = 0xBEEF Then $failureDescription = Call("Au3UnitConstraint" & $constraint & "_FailureDescription", $other)
	If @error = 0xDEAD And @extended = 0xBEEF Then $failureDescription = Call("Au3UnitConstraintConstraint_FailureDescription", $constraint, $other, $passedToContraint)
	$failureDescription = StringFormat("Failed asserting that %s.", $failureDescription)

	Local $additionalFailureDescription = Call("Au3UnitConstraint" & $constraint & "_AdditionalFailureDescription", $other)
	If @error = 0xDEAD And @extended = 0xBEEF Then $additionalFailureDescription = Call("Au3UnitConstraintConstraint_AdditionalFailureDescription", $other)

	If $additionalFailureDescription Then $failureDescription &= @CRLF & $additionalFailureDescription

	If Not ($description = "") Then $failureDescription = $description & @CRLF & $failureDescription

	ConsoleWriteError($failureDescription&@CRLF&@ScriptFullPath&":"&$line&@CRLF)
	Return SetError(1)
EndFunc

Func Au3UnitConstraintConstraint_Matches($other)
	Return False
EndFunc

#include "..\..\Exporter\Exporter.au3"
Func Au3UnitConstraintConstraint_FailureDescription($constraint, $other, $passedToContraint = Null)
	Local $toString =  Call("Au3UnitConstraint" & $constraint & "_ToString", $other)
	If @error = 0xDEAD And @extended = 0xBEEF Then $toString = Call("Au3UnitConstraintConstraint_ToString", $other)
	Return Au3ExporterExporter_Export(@NumParams=3?$passedToContraint:$other) & " " & $toString
EndFunc

Func Au3UnitConstraintConstraint_AdditionalFailureDescription($other)
	Return ""
EndFunc

Func Au3UnitConstraintConstraint_ToString()
	Return ""
EndFunc
