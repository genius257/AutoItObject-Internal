# AutoItObject Internal
Provides object creation via IDispatch interface in an attempt to add more OOP to AutoIt.

**Things to be added:**
* Garbage collection
* ~Method support~ **_Added in v0.1.2_**
* ~Support for more/all AutoIt variable-types~ **_Added in v1.0.0_**
* ~Accessors~ **_Added in v0.1.0_**
* Inheritance
* ~equivalent of @error and @extended~ **_Added in v1.0.0_**

`__defineGetter("property", Function)`

set the get accessor


`__defineSetter("property", Function)`

set the set accessor


~~`__defineMethod("Property", "Function")`~~ deprecated in v1.0.0

~~Set a function as a property value~~


`__unset("property")`

delete a property from the object


`__lock()`

prevents creation/changing of any properties, except values in already defined properties. To prevent value change of a property, use `__defineSetter`