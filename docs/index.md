## **Methods:**

### __get

#### **Syntax**

__get("property")

#### **Description:**

	get the value of property

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to get value of

### __set

#### **Syntax**

__set("property", value)

#### **Description:**

	set the value of property

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to get value of
Mixed | value | the value to set the property to

### __defineGetter

#### **Syntax**

__defineGetter("property", Function)

#### **Description:**

	set the get-accessor

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to set the get accessor
Function\|String | Function | Function to be called when getting the property value

### __lookupGetter

#### **Syntax**

__lookupGetter("property")

#### **Description:**

	return the get-accessor

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to set the get accessor

### __assign

#### **Syntax**

__assign(IDispatch[, ...])

#### **Description:**

	copy the values of all properties from one or more source objects to a the object

#### **Parameters:**

Type | Name | description
--- | --- | ---
IDispatch | IDispatch | Source IDispatch object. Note: normal IDispatch objects will result in a crash.

### __defineSetter

#### **Syntax**

__defineSetter("property", Function)

#### **Description:**

	set the set-accessor

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to set the get accessor
Function\|String | Function | Function to be called when setting the property value

### __lookupSetter

#### **Syntax**

__lookupSetter("property", Function)

#### **Description:**

	get the set-accessor

#### **Parameters:**

Type | Name | description
--- | --- | ---
String | "property" | The property to set the get accessor

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

### __preventExtensions

#### **Syntax**

__preventExtensions()

#### **Description:**

	prevents new properties from ever being added to the object

#### **Parameters:**

Type | Name | description
--- | --- | ---

### __isExtensible

#### **Syntax**

__isExtensible()

#### **Description:**

	determines if an object is extensible (whether it can have new properties added to it).

#### **Parameters:**

Type | Name | description
--- | --- | ---

### __seal

#### **Syntax**

__seal()

#### **Description:**

	seals the object, preventing new properties from being added to it and marking all existing properties as non-configurable. Values of present properties can still be changed

#### **Parameters:**

Type | Name | description
--- | --- | ---

### __isSealed

#### **Syntax**

__isSealed()

#### **Description:**

	determines if an object is sealed

#### **Parameters:**

Type | Name | description
--- | --- | ---

### __freeze

#### **Syntax**

__freeze()

#### **Description:**

	freezes an object: that is, prevents new properties from being added to it; prevents existing properties from being removed; and prevents existing properties, or their configurability, from being changed

#### **Parameters:**

Type | Name | description
--- | --- | ---

### __isFrozen

#### **Syntax**

__isFrozen()

#### **Description:**

	determines if an object is frozen.

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
Function\|String | DestructorFunction | Function to be called when lifetime of the object ends

## **Properties:**

### __case

#### **Type**

Bool

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
