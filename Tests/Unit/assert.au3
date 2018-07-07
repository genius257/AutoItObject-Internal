#include-once
; https://github.com/sebastianbergmann/phpunit/blob/7.2/src/Framework/Assert.php

Global $Au3UnitAssertCount = 0

#include "Constraint\Constraint.au3"
#include "Constraint\IsType.au3"

Func assertThat($value, $constraint, $message = "", $line = @ScriptLineNumber, $passedToContraint = Null)
	Local $constraintCount = "Au3UnitConstraint" & $constraint & "Count"
	$Au3UnitAssertCount += IsDeclared($constraintCount) ? Eval($constraintCount) : $Au3UnitConstraintCount
	Call("Au3UnitConstraint" & $constraint & "_Evaluate", $value, $message, False, $line, $passedToContraint)
	If @error==0xDEAD And @extended==0xBEEF Then Call("Au3UnitConstraintConstraint_Evaluate", $constraint, $value, $message, False, $line, $passedToContraint)
	Return SetError(@error)
EndFunc

; Func assertObjectHasAttribute($attributeName, $object, $message = "")
	;
; EndFunc

#include "Constraint\IsEqual.au3"
Func assertEquals($expected, $actual, $message = "", $line = @ScriptLineNumber)
	assertThat($actual, "IsEqual", $message, $line, $expected)
EndFunc

Func assertNotEquals($expected, $actual, $message = "", $line = @ScriptLineNumber)
	Local $passedToContraint = ["IsEqual", $expected]
	assertThat($actual, "LogicalNot", $message, $line, $passedToContraint)
EndFunc

#include "Constraint\IsFalse.au3"
Func assertFalse($condition, $message = "", $line = @ScriptLineNumber)
	assertThat($condition, "IsFalse", $message, $line, $condition)
EndFunc

Func assertNotFalse($condition, $message = "", $line = @ScriptLineNumber)
	Local $passedToContraint = ["IsFalse", $condition]
	assertThat($condition, "LogicalNot", $message, $line, $passedToContraint)
EndFunc

; Func assertGreaterThan($expected, $actual, $message = "")

; EndFunc

; Func assertGreaterThanOrEqual($expected, $actual, $message = "")

; EndFunc

; Func assertInfinite($variable, $message = "")

; EndFunc

; Func assertFinite($variable, $message = "")

; EndFunc

#include "Constraint\IsIdentical.au3"
Func assertInternalType($expected, $actual, $message = "", $line = @ScriptLineNumber)
	assertThat($actual, "IsType", $message, $line, $expected)
EndFunc

#include "Constraint\LogicalNot.au3"
Func assertNotInternalType($expected, $actual, $message = "", $line = @ScriptLineNumber)
	Local $passedToContraint = ["IsType", $expected]
	assertThat($actual, "LogicalNot", $message, $line, $passedToContraint)
EndFunc

; Func assertLessThan($expected, $actual, $message = "")

; EndFunc

; Func assertLessThanOrEqual($expected, $actual, $message = "")

; EndFunc

#include "Constraint\IsNull.au3"
Func assertNull($actual, $message = "", $line = @ScriptLineNumber)
	assertThat($actual, "IsNull", $message, $line)
EndFunc

#cs
Func assertNotNull()
	assertThat($actual, logicalNot("isNull"), $message)
EndFunc


Func assertStringMatchesFormat($format, $string, $message = "")

EndFunc
#ce

#include "Constraint\IsIdentical.au3"
Func assertSame($expected, $actual, $message = "", $line = @ScriptLineNumber)
	assertThat($actual, "IsIdentical", $message, $line, $expected)
EndFunc
#cs
Func assertNotSame($expected, $actual, $message = "")

EndFunc
#ce
#cs
Func assertStringEndsWith($suffix, $string, $message = "")

EndFunc

Func assertStringEndsNotWith($suffix, $string, $message = "")

EndFunc
#ce
#cs
Func assertStringStartsWith($prefix, $string, $message = "")

EndFunc

Func assertStringStartsNotWith($prefix, $string, $message = "")

EndFunc
#ce
#include "Constraint\IsTrue.au3"
Func assertTrue($condition, $message = "", $line = @ScriptLineNumber)
	assertThat($condition, "IsTrue", $message, $line, $condition)
EndFunc

Func assertNotTrue($condition, $message = "", $line = @ScriptLineNumber)
	Local $passedToContraint = ["IsTrue", $condition]
	assertThat($condition, "LogicalNot", $message, $line, $passedToContraint)
EndFunc
