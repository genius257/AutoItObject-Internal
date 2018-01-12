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

## **TODO:**
* ~Garbage collection~ **_Added in v1.0.3 with the exception of interface methods_**
* ~Method support~ **_Added in v0.1.2_**
* ~Support for more/all AutoIt variable-types~ **_Added in v1.0.0_**
* ~Accessors~ **_Added in v0.1.0_**
* Inheritance
* ~equivalent of @error and @extended~ **_Added in v1.0.0_**


## **Methods:**

### __defineGetter

#### **Syntax**

__defineGetter("property", Function)

#### **Description:**

	set the get-accessor

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to set the get accessor
Function | Function | Function to be called when getting the property value

### __defineSetter

#### **Syntax**

__defineSetter("property", Function)

#### **Description:**

	set the set-accessor

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to set the get accessor
Function | Function | Function to be called when setting the property value

### __keys

#### **Syntax**

__keys()

#### **Description:**

	get all property names defined in object

#### **Parameters:**

Type | Name | description
--- | --- | ---

### __unset

#### **Syntax**

__unset("property")

#### **Description:**

	delete a property from the object

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to delete

### __lock

#### **Syntax**

__lock()

#### **Description:**

	prevents creation/changing of any properties, except values in already defined properties. To prevent value change of a property, use `__defineSetter`

#### **Parameters:**

Type | Name | description
--- | --- | ---

### __destructor

#### **Syntax**

__destructor(DestructorFunction)

#### **Description:**

	DestructorFunction is called when the lifetime of the object ends

Type | Name | description
--- | --- | ---
Function | DestructorFunction | Function to be called when lifetime of the object ends

## **Other**

### Function

#### **Syntax**

Function([AccessorObject])

#### **Description:**

	The callback function used with accessors
	Function is a placeholder for the application-defined function name
	
#### **Parameters:**

Type | Name | description
--- | --- | ---
IDispatch | AccessorObject | A IDispatch object used to access passed data and self

### AccessorObject

#### **Description:**

	A locked IDispatch object, used with Accessor callback Function
	
#### **Properties:**

Type | Name | description | Access
--- | --- | --- | ---
IDispatch | parent | The IDispatch object containing the accessor.<br/>**WARNING:** accessing other properties via this object will still trigger accessors. Be careful | Read only
IDispatch | arguments | A locked IDispatch object.<br />For more info, see ArgumentsObject
*Any* | ret | The value passed in the set-accessor. This is not used by the get-accessor | Read and Write
*Any* | val | The value of the property the accessor is bound to | Read and Write

### ArgumentsObject

#### **Description:**

	A locked IDispatch object, used with the AccessorObject

#### **Properties:**

Type | Name | description | Access
--- | --- | --- | ---
int32 | length | number of arguments passed with the call | Read and Write
Array | values | the arguments passed with the call | Read and Write

### DestructorFunction([self])

#### **Description:**

	the callback function used with destructors

#### **Parameters:**

Type | Name | description
--- | --- | ---
IDispatch | self | The IDispatch object containing the destructor
