# Changelog
All notable changes to this project will be documented in this file

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
### Added

### Changed

### Removed

### Fixed
- Crash when creating too many objects (f86a0c40b45915a7b215b86a89aa5c2d92cdae23)
- Could not modify the default object method property `$oIDispatch()` (58fa551319727450e39ee171ae574438b0815372)

## [4.0.0] - 2022-07-08
### Added
- new global varaible `$__AOI_tagObject`
- new global enums `$__AOI_ConstantProperty_*`
- new test for `__unset` method
- internal use function `__AOI_Properties_Add`
- internal use function `__AOI_VariantReplace`
- internal use function `__AOI_Properties_Resize`
- internal use function `__AOI_Properties_Remove`
- internal use function `__AOI_Invoke_isExtensible`
- internal use function `__AOI_Invoke_case`
- internal use function `__AOI_Invoke_lookupSetter`
- internal use function `__AOI_Invoke_lookupGetter`
- internal use function `__AOI_Invoke_assign`
- internal use function `__AOI_Invoke_isSealed`
- internal use function `__AOI_Invoke_isFrozen`
- internal use function `__AOI_Invoke_get`
- internal use function `__AOI_Invoke_set`
- internal use function `__AOI_Invoke_exists`
- internal use function `__AOI_Invoke_destructor`
- internal use function `__AOI_Invoke_freeze`
- internal use function `__AOI_Invoke_seal`
- internal use function `__AOI_Invoke_preventExtensions`
- internal use function `__AOI_Invoke_unset`
- internal use function `__AOI_Invoke_keys`
- internal use function `__AOI_Invoke_defineGetter`
- internal use function `__AOI_Invoke_defineSetter`

### Changed
- use of new global variable `$__AOI_tagObject` for object struct reference.
- use of new global enums `$__AOI_ConstantProperty_*` for built in IDispatch object methods (negative id range).
- use of new global varaibles `$__AOI_Object_Element_*` for object pointer element offsets.
- IDispatch properties are now stored as an array instead of a singly linked list. (63836f6a2832c4e4123a00961a8d2184f58e0266)
- function rename `QueryInterface` => `__AOI_QueryInterface`
- function rename `AddRef` => `__AOI_AddRef`
- function rename `Release` => `__AOI_Release`
- function rename `GetIDsOfNames` => `__AOI_GetIDsOfNames`
- function rename `GetTypeInfo` => `__AOI_GetTypeInfo`
- function rename `GetTypeInfoCount` => `__AOI_GetTypeInfoCount`
- function rename `Invoke` => `__AOI_Invoke`
- function rename `MemCloneGlob` => `__AOI_MemCloneGlob`
- function rename `GlobalHandle` => `__AOI_GlobalHandle`
- function rename `VariantInit` => `__AOI_VariantInit`
- function rename `VariantCopy` => `__AOI_VariantCopy`
- function rename `VariantClear` => `__AOI_VariantClear`
- function rename `PrivateProperty` => `_AOI_PrivateProperty`
- variable rename `$IID_IUnknown` => `$__AOI_IID_IUnknown`
- variable rename `$IID_IDispatch` => `$__AOI_IID_IDispatch`
- variable rename `$IID_IConnectionPointContainer` => `$__AOI_IID_IConnectionPointContainer`
- variable rename `$DISPATCH_METHOD` => `$__AOI_DISPATCH_METHOD`
- variable rename `$DISPATCH_PROPERTYGET` => `$__AOI_DISPATCH_PROPERTYGET`
- variable rename `$DISPATCH_PROPERTYPUT` => `$__AOI_DISPATCH_PROPERTYPUT`
- variable rename `$DISPATCH_PROPERTYPUTREF` => `$__AOI_DISPATCH_PROPERTYPUTREF`
- variable rename `$S_OK` => `$__AOI_S_OK`
- variable rename `$E_NOTIMPL` => `$__AOI_E_NOTIMPL`
- variable rename `$E_NOINTERFACE` => `$__AOI_E_NOINTERFACE`
- variable rename `$E_POINTER` => `$__AOI_E_POINTER`
- variable rename `$E_ABORT` => `$__AOI_E_ABORT`
- variable rename `$E_FAIL` => `$__AOI_E_FAIL`
- variable rename `$E_ACCESSDENIED` => `$__AOI_E_ACCESSDENIED`
- variable rename `$E_HANDLE` => `$__AOI_E_HANDLE`
- variable rename `$E_OUTOFMEMORY` => `$__AOI_E_OUTOFMEMORY`
- variable rename `$E_INVALIDARG` => `$__AOI_E_INVALIDARG`
- variable rename `$E_UNEXPECTED` => `$__AOI_E_UNEXPECTED`
- variable rename `$DISP_E_UNKNOWNINTERFACE` => `$__AOI_DISP_E_UNKNOWNINTERFACE`
- variable rename `$DISP_E_MEMBERNOTFOUND` => `$__AOI_DISP_E_MEMBERNOTFOUND`
- variable rename `$DISP_E_PARAMNOTFOUND` => `$__AOI_DISP_E_PARAMNOTFOUND`
- variable rename `$DISP_E_TYPEMISMATCH` => `$__AOI_DISP_E_TYPEMISMATCH`
- variable rename `$DISP_E_UNKNOWNNAME` => `$__AOI_DISP_E_UNKNOWNNAME`
- variable rename `$DISP_E_NONAMEDARGS` => `$__AOI_DISP_E_NONAMEDARGS`
- variable rename `$DISP_E_BADVARTYPE` => `$__AOI_DISP_E_BADVARTYPE`
- variable rename `$DISP_E_EXCEPTION` => `$__AOI_DISP_E_EXCEPTION`
- variable rename `$DISP_E_OVERFLOW` => `$__AOI_DISP_E_OVERFLOW`
- variable rename `$DISP_E_BADINDEX` => `$__AOI_DISP_E_BADINDEX`
- variable rename `$DISP_E_UNKNOWNLCID` => `$__AOI_DISP_E_UNKNOWNLCID`
- variable rename `$DISP_E_ARRAYISLOCKED` => `$__AOI_DISP_E_ARRAYISLOCKED`
- variable rename `$DISP_E_BADPARAMCOUNT` => `$__AOI_DISP_E_BADPARAMCOUNT`
- variable rename `$DISP_E_PARAMNOTOPTIONAL` => `$__AOI_DISP_E_PARAMNOTOPTIONAL`
- variable rename `$DISP_E_BADCALLEE` => `$__AOI_DISP_E_BADCALLEE`
- variable rename `$DISP_E_NOTACOLLECTION` => `$__AOI_DISP_E_NOTACOLLECTION`
- variable rename `$tagVARIANT` => `$__AOI_tagVARIANT`
- variable rename `$tagDISPPARAMS` => `$__AOI_tagDISPPARAMS`
- variable rename `$VT_*` => `$__AOI_VT_*`
- variable rename `$tagProperty` => `$__AOI_tagProperty`

### Removed
- function `__AOI_PropertyCreate`
- function `LocalSize`
- function `LocalHandle`
- function `HeapSize`
- function `GetProcessHeap`
- function `VariantChangeType`
- function `VariantChangeTypeEx`
- function `__AOI_GetPtrValue`

## [3.0.0] - 2021-02-18
### Added
- __exists
- _AOI_ROT_register
- _AOI_ROT_revoke
- Internal use function __AOI_ROT_GetRunningObjectTable
- Internal use function __AOI_ROT_CreateFileMoniker

### Changed
- output of __keys() will now always return array
- value comparison operators for non string data now wont be type juggling values to strings

### Fixed
- QueryInterface did not call AddRef when returning valid pointer
- __unset did not handle as case insensitive if case insensitive mode was on
- Calling IDispatch object as function, would result in first element. Now property with key "" will be searched for. If not found in properties, object will throw exception for this special case!

## [2.0.0] - 2018-07-07
### Added
- This CHANGELOG.md file
- CONTRIBUTING.md file
- Tests
- __get
- __set
- __assign
- __freeze
- __isFrozen
- __isSealed
- __preventExtensions
- __isExtensible
- __lookupGetter
- __lookupSetter
- __seal
- __case
- internal use function __AOI_PropertyGetFromName
- internal use function __AOI_PropertyGetFromId
- internal use function __AOI_PropertyCreate
- internal use function __AOI_GetPtrOffset
- internal use function __AOI_GetPtrValue

### Changed
- Moved project to it's own repository
- __defineGetter and __defineSetter now also supports strings as second argument
- __destructor now also supports strings as first argument
- Split docs from readme into /docs/index.md
- documentation blocks are now following a more standardized docblock standard.

### Removed
- __lock

### Fixed
- desctructor pointer wrong offset after lock propterty was added to the object structure.

## [1.0.3] - 2017-08-11
### Added
- __desctructor
- Garbage collection, with the exception of interface methods

## [1.0.2] - 2017-08-09
### Added
- Parameters for assigning custom methods as IDispatch object internal functionality
- __keys
- PrivateProperty method for ease of use, when making private IDispatch object properties

## [1.0.1] - 2016-08-06
### Added
- Local scope for internal variables used by IDispatch object functionality

## [1.0.0] - 2016-08-06
### Added
- README.md file
- __unset
- __lock
- Accessors can set error code via setError
- Supprt for __all__ AutoIt variable types

## [0.1.2] - 2016-12-21
### Added
- __defineMethod
- Method support for objects

## [0.1.1] - 2016-11-25
### Changed
- Fix script breaking typo in Idispatch function, reported by [Chimp on the AutoIt forum thread](https://www.autoitscript.com/forum/topic/185720-autoitobject-pure-autoit/#elComment_1333941)

## [0.1.0] - 2016-11-24
### Added
- First iteration of the AutoItObject_Internal library
- Accessors via __defineGetter and __defineSetter
