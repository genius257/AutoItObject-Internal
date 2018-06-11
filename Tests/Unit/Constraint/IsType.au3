Func Au3UnitConstraintIsType_ToString($value)
	Return StringFormat('is of type "%s"', $value)
EndFunc

Func Au3UnitConstraintIsType_FailureDescription($constraint, $other, $exspected)
	Local $toString =  Call("Au3UnitConstraint" & $constraint & "_ToString", $other)
	If @error = 0xDEAD And @extended = 0xBEEF Then $toString = Call("Au3UnitConstraintConstraint_ToString", $other)
	Return Au3ExporterExporter_Export($exspected) & " " & $toString
EndFunc

Func Au3UnitConstraintIsType_Matches($other, $value)
	Return VarGetType($value) == $other
EndFunc