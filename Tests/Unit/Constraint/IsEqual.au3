#include-once
;https://github.com/sebastianbergmann/phpunit/blob/master/src/Framework/Constraint/IsEqual.php
;shallow implementation

Func Au3UnitConstraintIsEqual_Matches($other, $exspected)
	Return $exspected == $other
EndFunc

Func Au3UnitConstraintIsEqual_ToString($a)
	Return StringFormat("is equal to %s", $a)
EndFunc