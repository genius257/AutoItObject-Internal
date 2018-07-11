#include "..\AutoItObject_Internal.au3"
#include "Unit\assert.au3"

#AutoIt3Wrapper_Run_Au3Check=N

;TODO: add object struct property for retriving variant types through the Invoke function, name maybe Buffer or Variant
;TODO: implement helper IDispatch object with a function like: getValueFromVariant($pVariant):Mixed

$oError = ObjEvent("AutoIt.Error", "_ErrFunc")
; User's COM error function. Will be called if COM error occurs
Func _ErrFunc($oError)
    Return SetError($oError.number, 0, $oError.retcode)
EndFunc   ;==>_ErrFunc

testTypeSupport()
testAccessors()
testCase()
testDestructor()
testSeal()
testFreeze()
testPreventExtensions()
testAssign()
testMethods();tests the remaining non tested methods

Func testTypeSupport()
	Local $oIDispatch = IDispatch()

	$oIDispatch.a = $oIDispatch
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Object")
	$oIDispatch.a = Int(1, 1)
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Int32")
	$oIDispatch.a = Int(1, 2)
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Int64")
	$oIDispatch.a = 1.5
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Double")
	$oIDispatch.a = "string"
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "String")
	$oIDispatch.a = Binary(1)
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Binary")
	$oIDispatch.a = Ptr(1)
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Ptr", "Fail is expected, AutoIt related, not a fault of the library")
	$oIDispatch.a = HWnd(1)
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "HWnd", "Fail is expected, AutoIt related, not a fault of the library")
	$oIDispatch.a = True
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Bool")
	$oIDispatch.a = Null
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Keyword")
	$oIDispatch.a = Default
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Keyword")
	$oIDispatch.a = DllStructCreate("BYTE")
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "DLLStruct")
	$oIDispatch.a = MsgBox
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Function")
	$oIDispatch.a = IDispatch
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "UserFunction")
	Local $a = [1]
	$oIDispatch.a = $a
	assertEquals(@error, 0)
	assertInternalType($oIDispatch.a, "Array")
EndFunc

Func Accessor_Getter($oThis)
	Return $oThis.val
EndFunc

Func Accessor_Setter($oThis)
	$oThis.val = $oThis.val&$oThis.ret
	Return $oThis.val
EndFunc

Func testAccessors()
	Local $oIDispatch = IDispatch()
	$oIDispatch.a = "value"
	assertEquals(@error, 0)
	$oIDispatch.__defineGetter("a", "Accessor_Getter")
	assertEquals(@error, 0)
	$oIDispatch.__defineSetter("a", "Accessor_Setter")
	assertEquals(@error, 0)
	assertEquals($oIDispatch.a, "value")
	assertEquals($oIDispatch.__lookupGetter("a"), "Accessor_Getter")
	assertEquals($oIDispatch.__lookupSetter("a"), "Accessor_Setter")
	$oIDispatch.a = "value"
	assertEquals(@error, 0)
	assertEquals($oIDispatch.a, "valuevalue")

	$oIDispatch.__defineGetter("a", Accessor_Getter)
	assertEquals(@error, 0)
	$oIDispatch.__defineSetter("a", Accessor_Setter)
	assertEquals(@error, 0)
	assertEquals($oIDispatch.a, "valuevalue")
	assertEquals($oIDispatch.__lookupGetter("a"), Accessor_Getter)
	assertEquals($oIDispatch.__lookupSetter("a"), Accessor_Setter)
	$oIDispatch.a = "value"
	assertEquals(@error, 0)
	assertEquals($oIDispatch.a, "valuevaluevalue")

	$oIDispatch.__defineGetter("a", Null)
	assertNotEquals(@error, 0)
	$oIDispatch.__defineSetter("a", Null)
	assertNotEquals(@error, 0)
EndFunc

Func testCase()
	Local $oIDispatch = IDispatch()
	Local $value
	$value = $oIDispatch.__case
	assertEquals(@error, 0)
	assertEquals($value, True)
	$oIDispatch.a = "a"
	assertEquals(@error, 0)
	$oIDispatch.__seal()
	assertEquals(@error, 0)
	$value = $oIDispatch.a
	assertEquals(@error, 0)
	assertEquals($value, "a")
	$value = $oIDispatch.A
	assertNotEquals(@error, 0)
	assertNotEquals($value, "a")
	$oIDispatch.__case = False
	$value = $oIDispatch.a
	assertEquals(@error, 0)
	assertEquals($value, "a")
	$value = $oIDispatch.A
	assertEquals(@error, 0)
	assertEquals($value, "a")
EndFunc

Func __destructor()
	$__destructor+=1
EndFunc

Func testDestructor()
	$oIDispatch = IDispatch()
	Global $__destructor = 0
	$oIDispatch.__destructor(__destructor)
	assertEquals(@error, 0)
	$oIDispatch = Null
	assertEquals($__destructor, 1)
	$oIDispatch = IDispatch()
	$oIDispatch.__destructor("__destructor")
	assertEquals(@error, 0)
	$oIDispatch = Null
	assertEquals($__destructor, 2)
EndFunc

Func testSeal()
	Local $oIDispatch = IDispatch()
	$oIDispatch.a = "value"
	assertFalse($oIDispatch.__isSealed())
	$oIDispatch.__seal()
	assertEquals(@error, 0)
	assertTrue($oIDispatch.__isSealed())
	$oIDispatch.a = "otherValue"
	assertEquals(@error, 0)
	assertEquals($oIDispatch.a, "otherValue")
	$oIDispatch.b = "value";FIXME: will not result in @error flag. count keys to check
	assertNotEquals(@error, 0)
EndFunc

Func testFreeze()
	Local $oIDispatch = IDispatch()
	$oIDispatch.a = "value"
	assertFalse($oIDispatch.__isFrozen())
	$oIDispatch.__freeze()
	assertEquals(@error, 0)
	assertTrue($oIDispatch.__isFrozen())
	$oIDispatch.a = "otherValue"
	assertNotEquals(@error, 0)
	assertNotEquals($oIDispatch.a, "otherValue")
	$oIDispatch.b = "value"
	assertNotEquals(@error, 0)
EndFunc

Func testPreventExtensions()
	Local $oIDispatch = IDispatch()
	$oIDispatch.a = "value"
	assertFalse($oIDispatch.__isExtensible())
	$oIDispatch.__preventExtensions()
	assertEquals(@error, 0)
	assertTrue($oIDispatch.__isExtensible())
	$oIDispatch.a = "value"
	assertEquals(@error, 0)
	$oIDispatch.b
	assertNotEquals(@error, 0);FIXME: add the rest of the checks
EndFunc

Func testAssign()
	Local $oIDispatch1 = IDispatch()
	Local $oIDispatch2 = IDispatch()
	Local $oIDispatch3 = IDispatch()

	$oIDispatch3.a = 10
	$oIDispatch3.b = 20
	$oIDispatch3.c = 30

	$oIDispatch2.a = 1
	$oIDispatch2.b = 2
	$oIDispatch2.c = 3
	$oIDispatch2.d = 4

	$oIDispatch1.a = 3
	$oIDispatch1.b = 2
	$oIDispatch1.c = 1

	$oIDispatch1.__assign($oIDispatch2, $oIDispatch3)
	assertEquals(@error, 0)
	assertEquals($oIDispatch1.a, 10)
	assertEquals($oIDispatch1.b, 20)
	assertEquals($oIDispatch1.c, 30)
	assertEquals($oIDispatch1.d, 4)
EndFunc

Func testMethods()
	Local $oIDispatch = IDispatch()
	Local $value
	$oIDispatch.__set("a", "value")
	assertEquals(@error, 0)
	$value = $oIDispatch.__get("a")
	assertEquals(@error, 0)
	assertEquals($value, "value")
	$value = $oIDispatch.__keys
	assertEquals(@error, 0)
	assertEquals(UBound($value), 1)
EndFunc
