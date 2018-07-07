# AutoItObject Internal
Provides easy object creation via IDispatch interface.

**Example**
```AutoIt
#include "AutoItObject_Internal.au3"
$myCar = IDispatch()
$myCar.make = 'Ford'
$myCar.model = 'Mustang'
$myCar.year = 1969
$myCar.__defineGetter('DisplayCar', DisplayCar)

Func DisplayCar($oThis)
	Return 'A Beautiful ' & $oThis.parent.year & ' ' & $oThis.parent.make & ' ' & $oThis.parent.model
EndFunc

MsgBox(0, "", $myCar.DisplayCar)
```
