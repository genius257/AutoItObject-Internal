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

> The way this library implements objects is to mimic javascript object behavior and functionality.
>
> This approach was chosen for it's ease of use and for developers to be able to "jump right in" without learning a whole new syntax.

## Comparing AutoItObject Internal to AutoItObject UDF

| Feature | AutoItObject-Internal | AutoItObject UDF |
| - | - | - |
| getters | ✔ | ❌ |
| setters | ✔ | ❌ |
| destructors | ✔⚠️¹ | ✔ |
| inheritance | ✔⚠️² | ✔ |
| supports all au3 var types | ✔ | ❌ |
| native windows dlls and AutoIt only | ✔ | ❌ |
| dynamic number of method parameters | ✔ | ❌ |
| extending/editing UDF object logic | AutoIt | c++ |

 ¹ destructors are not triggered on script shutdown. 
 
 ² this is done by using the __assign method see https://github.com/genius257/AutoItObject-Internal/issues/1
