#include-once
#include <StringConstants.au3>
#include "..\..\helpers.au3"
;https://github.com/sebastianbergmann/phpunit/blob/master/src/Framework/Constraint/LogicalNot.php

Func Au3UnitConstraintLogicalNot_Negate($string)
	Local $matches, $negatedString, $nonInput

	Local $positives = [ _
		'contains ', _
		'exists', _
		'has ', _
		'is ', _
		'are ', _
		'matches ', _
		'starts with ', _
		'ends with ', _
		'reference ', _
		'not not ' _
	]

	Local $negatives = [ _
		'does not contain ', _
		'does not exist', _
		'does not have ', _
		'is not ', _
		'are not ', _
		'does not match ', _
		'starts not with ', _
		'ends not with ', _
		'don''''t reference ', _
		'not ' _
	]

	$matches = StringRegExp($string, '(?i)(\''''[\w\W]*'''')([\w\W]*)("[\w\W]*")', $STR_REGEXPARRAYFULLMATCH)

	If UBound($matches) > 0 Then
		$nonInput = $matches[2]

		$negatedString = PHP_str_replace($nonInput, PHP_str_replace($positives, $negatives, $nonInput), $string)
	Else
		$negatedString = PHP_str_replace($positives, $negatives, $string)
	EndIf

	Return $negatedString
EndFunc

Func Au3UnitConstraintLogicalNot_Evaluate($other, $description = "", $returnResult = false, $line = Null, $exspected = Null)
	Local $success = Not Call("Au3UnitConstraint" & $exspected[0] & "_Evaluate", $other, $description, True, $line, $exspected[1])
	If @error=0xDEAD And @extended=0xBEEF Then $success = Not Call("Au3UnitConstraint" & $exspected[0] & "_Evaluate", $other, $description, True, $line)
	If @error=0xDEAD And @extended=0xBEEF Then $success = Not Call("Au3UnitConstraintConstraint_Evaluate", $exspected[0], $other, $description, True, $line, $exspected[1])

	If $returnResult Then Return $success

	If Not $success Then
		Call("Au3UnitConstraintLogicalNot_Fail", $other, $description, Null, $line, $exspected)
		If @error = 0xDEAD And @extended = 0xBEEF Then Call("Au3UnitConstraintConstraint_Fail", "LogicalNot", $other, $description, Null, $line)
		If @error = 0xDEAD And @extended = 0xBEEF Then Exit MsgBox(0, "Au3Unit", "Au3UnitConstraintConstraint_Fail function is missing"&@CRLF&"Exitting") + 1
		Return SetError(@error)
	EndIf
EndFunc

Func Au3UnitConstraintLogicalNot_Fail($other, $description, $comparisonFailure, $line, $exspected)
	Local $failureDescription = Call("Au3UnitConstraintLogicalNot_FailureDescription", $other, $exspected)
	If @error = 0xDEAD And @extended = 0xBEEF Then $failureDescription = Call("Au3UnitConstraintConstraint_FailureDescription", "LogicalNot", $other)
	$failureDescription = StringFormat("Failed asserting that %s.", $failureDescription)

	Local $additionalFailureDescription = Call("Au3UnitConstraintLogicalNot_AdditionalFailureDescription", $other)
	If @error = 0xDEAD And @extended = 0xBEEF Then $additionalFailureDescription = Call("Au3UnitConstraintConstraint_AdditionalFailureDescription", $other)

	If $additionalFailureDescription Then $failureDescription &= @CRLF & $additionalFailureDescription

	If Not ($description = "") Then $failureDescription = $description & @CRLF & $failureDescription

	ConsoleWriteError($failureDescription&@CRLF&@ScriptFullPath&":"&$line&@CRLF)
	Return SetError(1)
EndFunc

Func Au3UnitConstraintLogicalNot_FailureDescription($other, $exspected)
	Switch $exspected
		Case "LogicalAnd", "LogicalNot", "LogicalOr"
			Local $iSize = UBound($exspected)-2
			Local $_exspected = Null
			If $iSize > 0 Then
				Local $_exspected[$iSize]
				Local $i
				For $i=0 To $iSize-1
					$_exspected[$i]=$exspected[2+$i]
				Next
			EndIf
			Local $failureDescription = Call("Au3UnitConstraint" & $exspected[0] & "_FailureDescription", $other, $exspected[0], $_exspected)
			If @error = 0xDEAD And @extended = 0xBEEF Then Call("Au3UnitConstraintConstraint_FailureDescription", $other, $exspected[0])
			Return 'not( ' & $failureDescription & ' )'
		Case Else
			Local $failureDescription = Call("Au3UnitConstraint" & $exspected[0] & "_FailureDescription", $exspected[0], $other, $exspected[1])
			If @error = 0xDEAD And @extended = 0xBEEF Then $failureDescription = Call("Au3UnitConstraintConstraint_FailureDescription", $exspected[0], $other)
			Return Au3UnitConstraintLogicalNot_Negate($failureDescription)
	EndSwitch
EndFunc

Func Au3UnitConstraintLogicalNot_ToString($value, $constraint = Null)
	Switch $constraint
		Case "LogicalAnd", "LogicalNot", "LogicalOr"
			Local $failureDescription = Call("Au3UnitConstraint" & $constraint & "_Tostring", $other)
			If @error = 0xDEAD And @extended = 0xBEEF Then Call("Au3UnitConstraintConstraint_ToString", $other)
			Return 'not( ' & $failureDescription & ' )'
		Case Else
			Local $failureDescription = Call("Au3UnitConstraint" & $constraint & "_ToString", $other)
			If @error = 0xDEAD And @extended = 0xBEEF Then Call("Au3UnitConstraintConstraint_ToString", $other)
			Return Au3UnitConstraintLogicalNot_Negate($failureDescription)
	EndSwitch
EndFunc
