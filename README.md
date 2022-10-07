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

## Comparing AutoItObject Internal to Other options

| Feature | AutoItObject-Internal | [AutoItObject UDF](https://www.autoitscript.com/forum/topic/110379-autoitobject-udf/) | [OOPEAu3](https://github.com/cosote/OOPEAu3) | [Map](https://www.autoitscript.com/autoit3/docs/intro/lang_variables.htm#ArrayMaps) |
| - | - | - | - | - |
| getters | :heavy_check_mark: | :x: | :x: | :x: |
| setters | :heavy_check_mark: | :x: | :x: | :x: |
| destructors | :heavy_check_mark::warning:¹ | :heavy_check_mark: | :question: | :x: |
| inheritance | :heavy_check_mark::warning:² | :heavy_check_mark: | :x: | :x: |
| supports all au3 var types | :heavy_check_mark: | :x: | :x: | :heavy_check_mark: |
| native windows dlls and AutoIt only | :heavy_check_mark: | :x: | :heavy_check_mark: | :heavy_check_mark: |
| dynamic number of method parameters | :heavy_check_mark: | :x: | :heavy_minus_sign:³ | :x: |
| registering objects in the ROT | :heavy_check_mark: | :heavy_check_mark: | :x: | :x: |
| extending/editing UDF object logic | AutoIt | c++ | AutoIt | :x: |

 ¹ destructors are not triggered on script shutdown. 
 
 ² this is done by using the __assign method see https://github.com/genius257/AutoItObject-Internal/issues/1

 ³ only for the constructor
