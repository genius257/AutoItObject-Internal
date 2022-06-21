#include-once
#include <Memory.au3>
#include <WinAPISys.au3>

If Not IsDeclared("IID_IUnknown") Then Global Const $IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
If Not IsDeclared("IID_IDispatch") Then Global Const $IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
If Not IsDeclared("IID_IConnectionPointContainer") Then Global Const $IID_IConnectionPointContainer = "{B196B284-BAB4-101A-B69C-00AA00341D07}"

If Not IsDeclared("DISPATCH_METHOD") Then Global Const $DISPATCH_METHOD =               1
If Not IsDeclared("DISPATCH_PROPERTYGET") Then Global Const $DISPATCH_PROPERTYGET =          2
If Not IsDeclared("DISPATCH_PROPERTYPUT") Then Global Const $DISPATCH_PROPERTYPUT =          4
If Not IsDeclared("DISPATCH_PROPERTYPUTREF") Then Global Const $DISPATCH_PROPERTYPUTREF =       8

If Not IsDeclared("S_OK") Then Global Const $S_OK = 0x00000000
If Not IsDeclared("E_NOTIMPL") Then Global Const $E_NOTIMPL = 0x80004001
If Not IsDeclared("E_NOINTERFACE") Then Global Const $E_NOINTERFACE = 0x80004002
If Not IsDeclared("E_POINTER") Then Global Const $E_POINTER = 0x80004003
If Not IsDeclared("E_ABORT") Then Global Const $E_ABORT = 0x80004004
If Not IsDeclared("E_FAIL") Then Global Const $E_FAIL = 0x80004005
If Not IsDeclared("E_ACCESSDENIED") Then Global Const $E_ACCESSDENIED = 0x80070005
If Not IsDeclared("E_HANDLE") Then Global Const $E_HANDLE = 0x80070006
If Not IsDeclared("E_OUTOFMEMORY") Then Global Const $E_OUTOFMEMORY = 0x8007000E
If Not IsDeclared("E_INVALIDARG") Then Global Const $E_INVALIDARG = 0x80070057
If Not IsDeclared("E_UNEXPECTED") Then Global Const $E_UNEXPECTED = 0x8000FFFF

If Not IsDeclared("DISP_E_UNKNOWNINTERFACE") Then Global Const $DISP_E_UNKNOWNINTERFACE = 0x80020001
If Not IsDeclared("DISP_E_MEMBERNOTFOUND") Then Global Const $DISP_E_MEMBERNOTFOUND = 0x80020003
If Not IsDeclared("DISP_E_PARAMNOTFOUND") Then Global Const $DISP_E_PARAMNOTFOUND = 0x80020004
If Not IsDeclared("DISP_E_TYPEMISMATCH") Then Global Const $DISP_E_TYPEMISMATCH = 0x80020005
If Not IsDeclared("DISP_E_UNKNOWNNAME") Then Global Const $DISP_E_UNKNOWNNAME = 0x80020006
If Not IsDeclared("DISP_E_NONAMEDARGS") Then Global Const $DISP_E_NONAMEDARGS = 0x80020007
If Not IsDeclared("DISP_E_BADVARTYPE") Then Global Const $DISP_E_BADVARTYPE = 0x80020008
If Not IsDeclared("DISP_E_EXCEPTION") Then Global Const $DISP_E_EXCEPTION = 0x80020009
If Not IsDeclared("DISP_E_OVERFLOW") Then Global Const $DISP_E_OVERFLOW = 0x8002000A
If Not IsDeclared("DISP_E_BADINDEX") Then Global Const $DISP_E_BADINDEX = 0x8002000B
If Not IsDeclared("DISP_E_UNKNOWNLCID") Then Global Const $DISP_E_UNKNOWNLCID = 0x8002000C
If Not IsDeclared("DISP_E_ARRAYISLOCKED") Then Global Const $DISP_E_ARRAYISLOCKED = 0x8002000D
If Not IsDeclared("DISP_E_BADPARAMCOUNT") Then Global Const $DISP_E_BADPARAMCOUNT = 0x8002000E
If Not IsDeclared("DISP_E_PARAMNOTOPTIONAL") Then Global Const $DISP_E_PARAMNOTOPTIONAL = 0x8002000F
If Not IsDeclared("DISP_E_BADCALLEE") Then Global Const $DISP_E_BADCALLEE = 0x80020010
If Not IsDeclared("DISP_E_NOTACOLLECTION") Then Global Const $DISP_E_NOTACOLLECTION = 0x80020011

Global Const $tagVARIANT = "ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2"
Global Const $tagDISPPARAMS = "ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;"

Global Const $__AOI_LOCK_CREATE = 1
Global Const $__AOI_LOCK_UPDATE = 2
Global Const $__AOI_LOCK_DELETE = 4
Global Const $__AOI_LOCK_CASE = 8

Global Enum $VT_EMPTY,$VT_NULL,$VT_I2,$VT_I4,$VT_R4,$VT_R8,$VT_CY,$VT_DATE,$VT_BSTR,$VT_DISPATCH, _
	$VT_ERROR,$VT_BOOL,$VT_VARIANT,$VT_UNKNOWN,$VT_DECIMAL,$VT_I1=16,$VT_UI1,$VT_UI2,$VT_UI4,$VT_I8, _
	$VT_UI8,$VT_INT,$VT_UINT,$VT_VOID,$VT_HRESULT,$VT_PTR,$VT_SAFEARRAY,$VT_CARRAY,$VT_USERDEFINED, _
	$VT_LPSTR,$VT_LPWSTR,$VT_RECORD=36,$VT_FILETIME=64,$VT_BLOB,$VT_STREAM,$VT_STORAGE,$VT_STREAMED_OBJECT, _
	$VT_STORED_OBJECT,$VT_BLOB_OBJECT,$VT_CF,$VT_CLSID,$VT_BSTR_BLOB=0xfff,$VT_VECTOR=0x1000, _
	$VT_ARRAY=0x2000,$VT_BYREF=0x4000,$VT_RESERVED=0x8000,$VT_ILLEGAL=0xffff,$VT_ILLEGALMASKED=0xfff, _
	$VT_TYPEMASK=0xfff

Global Const $__AOI_tagObject = "int RefCount;int Size;ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties;BYTE lock;PTR __destructor"
Global Const $tagProperty = "ptr Name;ptr Variant;ptr __getter;ptr __setter;ptr Next"

Global Enum Step -1 $__AOI_ConstantProperty_assign = -2, $__AOI_ConstantProperty_isExtensible, $__AOI_ConstantProperty_case, $__AOI_ConstantProperty_freeze, $__AOI_ConstantProperty_isFrozen, $__AOI_ConstantProperty_isSealed, $__AOI_ConstantProperty_keys, $__AOI_ConstantProperty_preventExtensions, $__AOI_ConstantProperty_defineGetter, $__AOI_ConstantProperty_defineSetter, $__AOI_ConstantProperty_lookupGetter, $__AOI_ConstantProperty_lookupSetter, $__AOI_ConstantProperty_seal, $__AOI_ConstantProperty_destructor = -16, $__AOI_ConstantProperty_unset, $__AOI_ConstantProperty_get, $__AOI_ConstantProperty_set, $__AOI_ConstantProperty_exists

Global Const $__AOI_Object_Element_RefCount = __AOI_GetPtrOffset("RefCount")
Global Const $__AOI_Object_Element_Size = __AOI_GetPtrOffset("Size")
Global Const $__AOI_Object_Element_Object = __AOI_GetPtrOffset("Object")
Global Const $__AOI_Object_Element_Methods = __AOI_GetPtrOffset("Methods")
Global Const $__AOI_Object_Element_Callbacks = __AOI_GetPtrOffset("Callbacks")
Global Const $__AOI_Object_Element_Properties = __AOI_GetPtrOffset("Properties")
Global Const $__AOI_Object_Element_lock = __AOI_GetPtrOffset("lock")
Global Const $__AOI_Object_Element___destructor = __AOI_GetPtrOffset("__destructor")

Func IDispatch($QueryInterface=QueryInterface, $AddRef=AddRef, $Release=Release, $GetTypeInfoCount=GetTypeInfoCount, $GetTypeInfo=GetTypeInfo, $GetIDsOfNames=GetIDsOfNames, $Invoke=Invoke)
	Local $tObject = DllStructCreate($__AOI_tagObject)

	$QueryInterface = DllCallbackRegister($QueryInterface, "LONG", "ptr;ptr;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($QueryInterface), 1)
	DllStructSetData($tObject, "Callbacks", $QueryInterface, 1)

	$AddRef = DllCallbackRegister($AddRef, "dword", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($AddRef), 2)
	DllStructSetData($tObject, "Callbacks", $AddRef, 2)

	$Release = DllCallbackRegister($Release, "dword", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Release), 3)
	DllStructSetData($tObject, "Callbacks", $Release, 3)

	$GetTypeInfoCount = DllCallbackRegister($GetTypeInfoCount, "long", "ptr;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfoCount), 4)
	DllStructSetData($tObject, "Callbacks", $GetTypeInfoCount, 4)

	$GetTypeInfo = DllCallbackRegister($GetTypeInfo, "long", "ptr;uint;int;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfo), 5)
	DllStructSetData($tObject, "Callbacks", $GetTypeInfo, 5)

	$GetIDsOfNames = DllCallbackRegister($GetIDsOfNames, "long", "ptr;ptr;ptr;uint;int;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetIDsOfNames), 6)
	DllStructSetData($tObject, "Callbacks", $GetIDsOfNames, 6)

	$Invoke = DllCallbackRegister($Invoke, "long", "ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Invoke), 7)
	DllStructSetData($tObject, "Callbacks", $Invoke, 7)

	DllStructSetData($tObject, "RefCount", 1) ; initial ref count is 1
	DllStructSetData($tObject, "Size", 7) ; number of interface methods

	Local $pData = MemCloneGlob($tObject)

	Local $tObject = DllStructCreate($__AOI_tagObject, $pData)

	DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
	Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IDispatch, Default, True) ; pointer that's wrapped into object
EndFunc

#cs
# @internal
#ce
Func QueryInterface($pSelf, $pRIID, $pObj)
	If $pObj=0 Then Return $E_POINTER
	Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
	If (Not ($sGUID=$IID_IDispatch)) And (Not ($sGUID=$IID_IUnknown)) Then Return $E_NOINTERFACE
	Local $tStruct = DllStructCreate("ptr", $pObj)
	DllStructSetData($tStruct, 1, $pSelf)
	AddRef($pSelf)
	Return $S_OK
EndFunc

#cs
# @internal
#ce
Func AddRef($pSelf)
   Local $tStruct = DllStructCreate("int Ref", $pSelf + $__AOI_Object_Element_RefCount)
   $tStruct.Ref += 1
   Return $tStruct.Ref
EndFunc

#cs
# @internal
#ce
Func Release($pSelf)
	Local $i
	Local $tObject = DllStructCreate($__AOI_tagObject, $pSelf + $__AOI_Object_Element_RefCount)
	$tObject.RefCount -= 1
	If $tObject.RefCount = 0 Then; initiate garbage collection
		Local $pDescructor = $tObject.__destructor
		If Not ($pDescructor=0) Then
			Local $tVARIANT = DllStructCreate($tagVARIANT, $pDescructor)
			$tObject.__destructor = 0
			Local $IDispatch = IDispatch()
			$IDispatch.a=0
			$tIDispatch = DllStructCreate($__AOI_tagObject, Ptr($IDispatch) + $__AOI_Object_Element_RefCount)
			Local $pProperty = $tIDispatch.Properties
			Local $pVARIANT = DllStructGetData(DllStructCreate($tagProperty, $pProperty),"Variant")
			VariantClear($pVARIANT)
			VariantCopy($pVARIANT, $tVARIANT)
			Local $f__destructor = $IDispatch.a ;FIXME: make a helper method for convertion between au3 varaible and variant (with optional copy variant flag)
			VariantClear($pVARIANT)
			$tObject.RefCount += 1
			Local $tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
			$tVARIANT.vt = $VT_DISPATCH
			$tVARIANT.data = $pSelf
			Call($f__destructor, $IDispatch.a)
			VariantClear($pVARIANT)
			$IDispatch=0
		EndIf
		$tObject.lock = 1;lock
		Local $pProperty = $tObject.Properties;get first property
		$tObject.Properties = 0;detatch properties from object
		While 1;releases all properties
			If $pProperty=0 Then ExitLoop
			Local $tProperty = DllStructCreate($tagProperty, $pProperty)
			Local $_pProperty = $pProperty
			$pProperty = $tProperty.Next
			If Not ($tProperty.__getter=0) Then
				VariantClear($tProperty.__getter)
				_MemGlobalFree(GlobalHandle($tProperty.__getter))
			EndIf
			If Not ($tProperty.__setter=0) Then
				VariantClear($tProperty.__setter)
				_MemGlobalFree(GlobalHandle($tProperty.__setter))
			EndIf
			VariantClear($tProperty.Variant)
			_MemGlobalFree(GlobalHandle($tProperty.Variant))
			_WinAPI_FreeMemory($tProperty.Name)
			$tProperty=0
			_MemGlobalFree(GlobalHandle($_pProperty))
		WEnd
		_MemGlobalFree(GlobalHandle(DllStructGetPtr($tObject)))
		Return 0
	EndIf
	Return $tObject.RefCount
EndFunc

#cs
# @internal
#ce
Func GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
	Local $tIds = DllStructCreate("long", $rgDispId); 2,147,483,647 properties available to define, per object. And 2,147,483,647 private properties to set in the negative space, per object.
	Local $tProperty = 0

	Local $pStr = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)
	Local $iSize = _WinAPI_StrLen($pStr, True)
	If $iSize = 0 Then $iSize = 1
	Local $t_rgszNames = DllStructCreate("WCHAR["&$iSize&"]", $pStr)
	Local $s_rgszName = DllStructGetData($t_rgszNames, 1)

	DllStructSetData($tIds, 1, -1)
	if StringLeft($s_rgszName, 2) = "__" Then
		__AOI_ConstantProperty_Lookup($s_rgszName, $tIds)
		If DllStructGetData($tIds, 1) <> -1 Then Return $S_OK
	EndIf

	Local $tObject = DllStructCreate($__AOI_tagObject, $pSelf + $__AOI_Object_Element_RefCount)
	Local $iLock = $tObject.lock
	Local $bCase = Not (BitAND($iLock, $__AOI_LOCK_CASE)>0)
	Local $pProperty = __AOI_PropertyGetFromName($tObject.Properties, $pStr, $bCase)
	Local $iID = @error<>0?-1:@extended
	Local $iIndex = @extended

	If ($iID=-1) And BitAND($iLock, $__AOI_LOCK_CREATE)=0 Then
		Local $pData = __AOI_PropertyCreate($s_rgszName)
		If $iIndex = -1 Then;first item in list
			$tObject.Properties = $pData
		Else
			$tProperty = DllStructCreate($tagProperty, $pProperty)
			$tProperty.next = $pData
		EndIf
		$iID = $iIndex+1
	EndIf

	If $iID=-1 Then Return $DISP_E_UNKNOWNNAME
	$iID += 1
	DllStructSetData($tIds, 1, $iID)
	Return $S_OK
EndFunc

#cs
# @internal
#ce
Func __AOI_ConstantProperty_Lookup($s_rgszName, $tIds)
	Switch $s_rgszName
		Case "__assign"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_assign)
		Case "__isExtensible"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_isExtensible)
		Case "__case"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_case)
		Case "__freeze"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_freeze)
		Case "__isFrozen"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_isFrozen)
		Case "__isSealed"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_isSealed)
		Case "__keys"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_keys)
		Case "__preventExtensions"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_preventExtensions)
		Case "__defineGetter"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_defineGetter)
		Case "__defineSetter"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_defineSetter)
		Case "__lookupGetter"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_lookupGetter)
		Case "__lookupSetter"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_lookupSetter)
		Case "__seal"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_seal)
		Case "__destructor"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_destructor)
		Case "__unset"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_unset)
		Case "__get"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_get)
		Case "__set"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_set)
		Case "__exists"
			DllStructSetData($tIds, 1, $__AOI_ConstantProperty_exists)
	EndSwitch
EndFunc

#cs
# @internal
#ce
Func GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
	If $iTInfo<>0 Then Return $DISP_E_BADINDEX
	If $ppTInfo=0 Then Return $E_INVALIDARG
	Return $S_OK
EndFunc

#cs
# @internal
#ce
Func GetTypeInfoCount($pSelf, $pctinfo)
	DllStructSetData(DllStructCreate("UINT",$pctinfo),1, 0)
	Return $S_OK
EndFunc

#cs
# @internal
#ce
Func Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
	Local $tObject = DllStructCreate($__AOI_tagObject, $pSelf + $__AOI_Object_Element_RefCount)
	If $dispIdMember = 0 Then
		Local $iLock = $tObject.lock
		Local $bCase = Not (BitAND($iLock, $__AOI_LOCK_CASE)>0)
		Local Static $tEmptyString = DllStructCreate("WCHAR")
		__AOI_PropertyGetFromName($tObject.Properties, DllStructGetPtr($tEmptyString), $bCase)
		$dispIdMember = @error<>0?-1:(@extended + 1)
	EndIf
	If $dispIdMember=-1 Then Return $DISP_E_MEMBERNOTFOUND
	Local $tVARIANT, $_tVARIANT, $tDISPPARAMS
	Local $t
	Local $i

	Local $pProperty = $tObject.Properties

	If $dispIdMember<-1 Then
		Switch $dispIdMember
			Case $__AOI_ConstantProperty_isExtensible
				Return __AOI_Invoke_isExtensible($tObject, $pVarResult)
			Case $__AOI_ConstantProperty_case
				Return __AOI_Invoke_case($tObject, $pDispParams, $pVarResult, $wFlags)
			Case $__AOI_ConstantProperty_lookupSetter
				Return __AOI_Invoke_lookupSetter($tObject, $riid, $lcid, $pDispParams, $pVarResult)
			Case $__AOI_ConstantProperty_lookupGetter
				Return __AOI_Invoke_lookupGetter($tObject, $riid, $lcid, $pDispParams, $pVarResult)
			Case $__AOI_ConstantProperty_assign
				Return __AOI_Invoke_assign($tObject, $pDispParams)
			Case $__AOI_ConstantProperty_isSealed
				Return __AOI_Invoke_isSealed($tObject, $pVarResult)
			Case $__AOI_ConstantProperty_isFrozen
				Return __AOI_Invoke_isFrozen($tObject, $pVarResult)
			Case $__AOI_ConstantProperty_get
				Return __AOI_Invoke_get($pSelf, $riid, $lcid, $pDispParams, $puArgErr, $pExcepInfo, $pVarResult)
			Case $__AOI_ConstantProperty_set
				Return __AOI_Invoke_set($pSelf, $riid, $lcid, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
			Case $__AOI_ConstantProperty_exists
				Return __AOI_Invoke_exists($tObject, $pDispParams, $pVarResult)
			Case $__AOI_ConstantProperty_destructor
				Return __AOI_Invoke_destructor($tObject, $pDispParams)
			Case $__AOI_ConstantProperty_freeze
				Return __AOI_Invoke_freeze($tObject, $pDispParams)
			Case $__AOI_ConstantProperty_seal
				Return __AOI_Invoke_seal($tObject, $pDispParams)
			Case $__AOI_ConstantProperty_preventExtensions
				Return __AOI_Invoke_preventExtensions($tObject, $pDispParams)
			Case $__AOI_ConstantProperty_unset
				Return __AOI_Invoke_unset($tObject, $pDispParams, $pProperty)
			Case $__AOI_ConstantProperty_keys
				Return __AOI_Invoke_keys($tObject, $pVarResult)
			Case $__AOI_ConstantProperty_defineGetter
				Return __AOI_Invoke_defineGetter($tObject, $pDispParams, $lcid)
			Case $__AOI_ConstantProperty_defineSetter
				Return __AOI_Invoke_defineSetter($tObject, $pDispParams, $lcid)
		EndSwitch
		Return $DISP_E_EXCEPTION
	EndIf

	Local $tProperty = __AOI_PropertyGetFromId($pProperty, $dispIdMember)

	$tVARIANT = DllStructCreate($tagVARIANT, $tProperty.Variant)

	If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)=0)) Then
		If Not($tProperty.__getter = 0) Then
			$_tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			Local $oIDispatch = IDispatch()
			$oIDispatch.val = 0
			$oIDispatch.ret = 0
			DllStructSetData(DllStructCreate("INT", $pSelf-4-4), 1, DllStructGetData(DllStructCreate("INT", $pSelf-4-4), 1)+1)
			$oIDispatch.parent = 0
			Local $tProperty02 = DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", ptr($oIDispatch) + $__AOI_Object_Element_Properties),1))
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tVARIANT = DllStructCreate($tagVARIANT, $tProperty02.Variant)
			$tVARIANT.vt = $VT_DISPATCH
			$tVARIANT.data = $pSelf
			$oIDispatch.arguments = IDispatch();
			$oIDispatch.arguments.length=$tDISPPARAMS.cArgs
			Local $aArguments[$tDISPPARAMS.cArgs], $iArguments=$tDISPPARAMS.cArgs-1
			Local $_pProperty = DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + $__AOI_Object_Element_Properties),1)
			Local $_tProperty = DllStructCreate($tagProperty, $_pProperty)
			For $i=0 To $iArguments
				VariantClear($_tProperty.Variant)
				VariantCopy($_tProperty.Variant, $tDISPPARAMS.rgvargs+(($iArguments-$i)*DllStructGetSize($_tVARIANT)))
				$aArguments[$i]=$oIDispatch.val
			Next
			$oIDispatch.arguments.values=$aArguments
			$oIDispatch.arguments.__seal()
			$oIDispatch.__defineSetter("parent", PrivateProperty)
			VariantClear($_tProperty.Variant)
			VariantCopy($_tProperty.Variant, $tProperty.__getter)
			Local $fGetter = $oIDispatch.val
			VariantClear($_tProperty.Variant)
			VariantCopy($_tProperty.Variant, $tProperty.Variant)
			$oIDispatch.__seal()
			Local $mRet = Call($fGetter, $oIDispatch)
			Local $iError = @error, $iExtended = @extended
			VariantClear($tProperty.Variant)
			VariantCopy($tProperty.Variant, $_tProperty.Variant)
			$oIDispatch.ret = $mRet
			$_tProperty = DllStructCreate($tagProperty, $_tProperty.Next)
			VariantCopy($pVarResult, $_tProperty.Variant)
			$oIDispatch=0
			Return ($iError<>0)?$DISP_E_EXCEPTION:$S_OK
		EndIf

		VariantCopy($pVarResult, $tVARIANT)
		Return $S_OK
	Else; ~ $DISPATCH_PROPERTYPUT
		$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
		If Not ($tProperty.__setter=0) Then
			Local $oIDispatch = IDispatch()
			$oIDispatch.val = 0
			$oIDispatch.ret = 0
			DllStructSetData(DllStructCreate("INT", $pSelf-4-4), 1, DllStructGetData(DllStructCreate("INT", $pSelf-4-4), 1)+1)
			$oIDispatch.parent = 0
			Local $tProperty02 = DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", ptr($oIDispatch) + $__AOI_Object_Element_Properties),1))
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tVARIANT = DllStructCreate($tagVARIANT, $tProperty02.Variant)
			$tVARIANT.vt = $VT_DISPATCH
			$tVARIANT.data = $pSelf
			Local $_pProperty = DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + $__AOI_Object_Element_Properties),1)
			Local $_tProperty = DllStructCreate($tagProperty, $_pProperty)
			Local $_tProperty2 = DllStructCreate($tagProperty, $_tProperty.Next)
			VariantClear($_tProperty.Variant)
			VariantCopy($_tProperty.Variant, $tProperty.__setter)
			VariantClear($_tProperty2.Variant)
			VariantCopy($_tProperty2.Variant, $tDISPPARAMS.rgvargs)
			Local $fSetter = $oIDispatch.val
			VariantClear($_tProperty.Variant)
			VariantCopy($_tProperty.Variant, $tProperty.Variant)
			$oIDispatch.__seal()
			Local $mRet = Call($fSetter, $oIDispatch)
			Local $iError = @error, $iExtended = @extended
			VariantClear($tProperty.Variant)
			VariantCopy($tProperty.Variant, $_tProperty.Variant)
			$oIDispatch.ret = $mRet
			$_tProperty = DllStructCreate($tagProperty, $_tProperty.Next)
			VariantCopy($pVarResult, $_tProperty.Variant)
			$oIDispatch=0
			Return ($iError<>0)?$DISP_E_EXCEPTION:$S_OK
		EndIf

		Local $iLock = $tObject.lock
		If BitAND($iLock, $__AOI_LOCK_UPDATE)>0 Then Return $DISP_E_EXCEPTION

		$_tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
		VariantClear($tVARIANT)
		VariantCopy($tVARIANT, $_tVARIANT)
	EndIf
	Return $S_OK
EndFunc

Func __AOI_Invoke_isExtensible($tObject, $pVarResult)
	Local $iLock = $tObject.lock
	Local $iExtensible = $__AOI_LOCK_CREATE
	$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
	$tVARIANT.vt = $VT_BOOL
	$tVARIANT.data = (BitAND($iLock, $iExtensible) = $iExtensible)?1:0
	Return $S_OK
EndFunc

Func __AOI_Invoke_case($tObject, $pDispParams, $pVarResult, $wFlags)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)=0)) Then
		If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
		$tVARIANT=DllStructCreate($tagVARIANT, $pVarResult)
		$tVARIANT.vt = $VT_BOOL
		Local $iLock = $tObject.lock
		$tVARIANT.data = (BitAND($iLock, $__AOI_LOCK_CASE)>0)?0:1
	Else; $DISPATCH_PROPERTYPUT
		If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
		$tVARIANT=DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
		If $tVARIANT.vt<>$VT_BOOL Then Return $DISP_E_BADVARTYPE
		Local $iLock = $tObject.lock
		If BitAND($iLock, $__AOI_LOCK_UPDATE)>0 Then Return $DISP_E_EXCEPTION
		$b = $tVARIANT.data
		$tObject.lock = (Not $b) ? BitOR($iLock, $__AOI_LOCK_CASE) : BitAND($iLock, BitNOT(BitShift(1 , 0-(Log($__AOI_LOCK_CASE)/log(2)))))
	EndIf
	Return $S_OK
EndFunc

Func __AOI_Invoke_lookupSetter($tObject, $riid, $lcid, $pDispParams, $pVarResult)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT

	$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
	DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
	DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs)
	$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
	$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
	If Not GetIDsOfNames($tObject.Object, $riid, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) = $S_OK Then Return $DISP_E_EXCEPTION

	$pProperty = $tObject.Properties
	$tProperty = __AOI_PropertyGetFromId($pProperty, $t.id)
	If Not $tProperty.__setter=0 Then
		VariantClear($pVarResult)
		VariantCopy($pVarResult, $tProperty.__setter)
	EndIf
	Return $S_OK
EndFunc

Func __AOI_Invoke_lookupGetter($tObject, $riid, $lcid, $pDispParams, $pVarResult)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT

	$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
	DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
	DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs)
	$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
	$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
	If Not GetIDsOfNames($tObject.Object, $riid, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) = $S_OK Then Return $DISP_E_EXCEPTION

	$pProperty = $tObject.Properties
	$tProperty = __AOI_PropertyGetFromId($pProperty, $t.id)
	If Not $tProperty.__getter=0 Then
		VariantClear($pVarResult)
		VariantCopy($pVarResult, $tProperty.__getter)
	EndIf
	Return $S_OK
EndFunc

Func __AOI_Invoke_assign($tObject, $pDispParams)
	Local $iLock = $tObject.lock
	If BitAND($iLock, $__AOI_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION


	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs=0 Then Return $DISP_E_BADPARAMCOUNT

	Local $tVARIANT = DllStructCreate($tagVARIANT)
	Local $iVARIANT = DllStructGetSize($tVARIANT)
	Local $pExternalProperty, $tExternalProperty
	Local $pProperty, $tProperty
	Local $iID, $iIndex, $pData
	For $i=$tDISPPARAMS.cArgs-1 To 0 Step -1
		$tVARIANT=DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+$iVARIANT*$i)
		If Not (DllStructGetData($tVARIANT, "vt")=$VT_DISPATCH) Then Return $DISP_E_BADVARTYPE
		$pExternalProperty = __AOI_GetPtrValue(DllStructGetData($tVARIANT, "data") + $__AOI_Object_Element_Properties, "ptr")
		While 1
			If $pExternalProperty = 0 Then ExitLoop
			$tExternalProperty = DllStructCreate($tagProperty, $pExternalProperty)

			$pProperty = __AOI_PropertyGetFromName($tObject.Properties, $tExternalProperty.Name, False)
			$iID = @error<>0?-1:@extended
			$iIndex = @extended

			If ($iID=-1) Then
				$pData = __AOI_PropertyCreate(_WinAPI_GetString($tExternalProperty.Name))
				$tProperty = DllStructCreate($tagProperty, $pData)
				VariantClear($tProperty.Variant)
				VariantCopy($tProperty.Variant, $tExternalProperty.Variant)

				If $iIndex=-1 Then;first item in list
					$tObject.Properties = $pData
				Else
					$tProperty = DllStructCreate($tagProperty, $pProperty)
					$tProperty.Next = $pData
				EndIf
			Else
				$tProperty = DllStructCreate($tagProperty, $pProperty)
				VariantClear($tProperty.Variant)
				VariantCopy($tProperty.Variant, $tExternalProperty.Variant)
			EndIf

			$pExternalProperty = $tExternalProperty.Next
		WEnd
	Next
	Return $S_OK
EndFunc

Func __AOI_Invoke_isSealed($tObject, $pVarResult)
	Local $iLock = $tObject.lock
	Local $iSeal = $__AOI_LOCK_CREATE + $__AOI_LOCK_DELETE
	$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
	$tVARIANT.vt = $VT_BOOL
	$tVARIANT.data = (BitAND($iLock, $iSeal) = $iSeal)?1:0
	Return $S_OK
EndFunc

Func __AOI_Invoke_isFrozen($tObject, $pVarResult)
	Local $iLock = $tObject.lock
	Local $iFreeze = $__AOI_LOCK_CREATE + $__AOI_LOCK_UPDATE + $__AOI_LOCK_DELETE
	$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
	$tVARIANT.vt = $VT_BOOL
	$tVARIANT.data = (BitAND($iLock, $iFreeze) = $iFreeze)?1:0
	Return $S_OK
EndFunc

Func __AOI_Invoke_get($pSelf, $riid, $lcid, $pDispParams, $puArgErr, $pExcepInfo, $pVarResult)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
	$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
	If $tVARIANT.vt<>$VT_BSTR Then Return $DISP_E_BADVARTYPE
	$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
	DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
	DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs)
	$t.str_ptr = DllStructGetData($tVARIANT, "data")
	$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
	If Not GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) = $S_OK Then Return $DISP_E_EXCEPTION
	Return Invoke($pSelf, $t.id, $riid, $lcid, $DISPATCH_PROPERTYGET, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
EndFunc

Func __AOI_Invoke_set($pSelf, $riid, $lcid, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
	$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
	If $tVARIANT.vt<>$VT_BSTR Then Return $DISP_E_BADVARTYPE
	$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
	DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
	$t.str_ptr = DllStructGetData($tVARIANT, "data")
	$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
	If Not GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) = $S_OK Then Return $DISP_E_EXCEPTION
	$tDISPPARAMS.cArgs=1
	Return Invoke($pSelf, $t.id, $riid, $lcid, $DISPATCH_PROPERTYPUT, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
EndFunc

Func __AOI_Invoke_exists($tObject, $pDispParams, $pVarResult)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
	$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
	If $tVARIANT.vt<>$VT_BSTR Then Return $DISP_E_BADVARTYPE
	Local $iLock = $tObject.lock
	Local $bCase = Not (BitAND($iLock, $__AOI_LOCK_CASE)>0)
	Local $pProperty = __AOI_PropertyGetFromName($tObject.Properties, $tVARIANT.data, $bCase)
	Local $error = @error
	$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
	$tVARIANT.vt = $VT_BOOL
	$tVARIANT.data = $error<>0?0:1
	Return $S_OK
EndFunc

Func __AOI_Invoke_destructor($tObject, $pDispParams)
	Local $iLock = $tObject.lock
	If BitAND($iLock, $__AOI_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
	If (Not (DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs),"vt")=$VT_RECORD)) And (Not (DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs),"vt")=$VT_BSTR)) Then Return $DISP_E_BADVARTYPE
	Local $tVARIANT = DllStructCreate($tagVARIANT)
	Local $pVARIANT = MemCloneGlob($tVARIANT)
	$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
	VariantInit($pVARIANT)
	VariantCopy($pVARIANT, $tDISPPARAMS.rgvargs)
	$tObject.__destructor = $pVARIANT
	Return $S_OK
EndFunc

Func __AOI_Invoke_freeze($tObject, $pDispParams)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
	$tObject.lock = BitOR($tObject.lock, $__AOI_LOCK_CREATE + $__AOI_LOCK_DELETE + $__AOI_LOCK_UPDATE)
	Return $S_OK
EndFunc

Func __AOI_Invoke_seal($tObject, $pDispParams)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
	$tObject.lock = BitOR($tObject.lock, $__AOI_LOCK_CREATE + $__AOI_LOCK_DELETE)
	Return $S_OK
EndFunc

Func __AOI_Invoke_preventExtensions($tObject, $pDispParams)
	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
	$tObject.lock = BitOR($tObject.lock, $__AOI_LOCK_CREATE)
	Return $S_OK
EndFunc

Func __AOI_Invoke_unset($tObject, $pDispParams, $pProperty)
	Local $iLock = $tObject.lock
	If BitAND($iLock, $__AOI_LOCK_DELETE)>0 Then Return $DISP_E_EXCEPTION

	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
	$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
	If Not($VT_BSTR=$tVARIANT.vt) Then Return $DISP_E_BADVARTYPE
	Local $sProperty = _WinAPI_GetString($tVARIANT.data);the string to search for
	Local $tProperty=0,$tProperty_Prev
	While 1
		If $pProperty=0 Then ExitLoop
		$tProperty_Prev = $tProperty
		$tProperty = DllStructCreate($tagProperty, $pProperty)
		Local $bCase = Not (BitAND($iLock, $__AOI_LOCK_CASE)>0)
		If ($bCase And _WinAPI_GetString($tProperty.Name)==$sProperty) Or ((Not $bCase) And _WinAPI_GetString($tProperty.Name)=$sProperty) Then
			If $tProperty_Prev=0 Then
				$tObject.Properties = $tProperty.Next
			Else
				$tProperty_Prev.Next = $tProperty.next
			EndIf
			VariantClear($tProperty.Variant)
			_MemGlobalFree(GlobalHandle($tProperty.Variant))
			$tProperty = 0
			_MemGlobalFree(GlobalHandle($pProperty))
			Return $S_OK
		EndIf
		$pProperty = $tProperty.Next
	WEnd
	Return $DISP_E_MEMBERNOTFOUND
EndFunc

Func __AOI_Invoke_keys($tObject, $pVarResult)
	Local $aKeys[0]
	Local $pProperty = $tObject.Properties
	While 1
		If $pProperty=0 Then ExitLoop
		Local $tProperty = DllStructCreate($tagProperty, $pProperty)
		ReDim $aKeys[UBound($aKeys,1)+1]
		$aKeys[UBound($aKeys,1)-1] = DllStructGetData(DllStructCreate("WCHAR["&_WinAPI_StrLen($tProperty.Name)&"]", $tProperty.Name), 1)
		If $tProperty.next=0 Then ExitLoop
		$pProperty = $tProperty.next
	WEnd
	Local $oIDispatch = IDispatch()
	$oIDispatch.a=$aKeys
	VariantClear($pVarResult)
	VariantCopy($pVarResult, DllStructGetData(DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + $__AOI_Object_Element_Properties),1)), "Variant"))
	$oIDispatch=0
	Return $S_OK
EndFunc

Func __AOI_Invoke_defineGetter($tObject, $pDispParams, $lcid)
	Local $iLock = $tObject.lock
	If BitAND($iLock, $__AOI_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION

	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
	$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
	If Not (($tVARIANT.vt=$VT_RECORD) Or ($tVARIANT.vt=$VT_BSTR)) Then Return $DISP_E_BADVARTYPE
	Local $tVARIANT2 = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize($tVARIANT))
	If Not ($tVARIANT2.vt=$VT_BSTR) Then Return $DISP_E_BADVARTYPE

	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
	DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
	DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
	$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
	$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
	GetIDsOfNames($tObject.Object, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"))

	$pProperty = $tObject.Properties
	$tProperty = __AOI_PropertyGetFromId($pProperty, $t.id)

	If ($tProperty.__getter=0) Then
		Local $tVARIANT_Getter = DllStructCreate($tagVARIANT)
		$pVARIANT_Getter = MemCloneGlob($tVARIANT_Getter)
		VariantInit($pVARIANT_Getter)
	Else
		Local $pVARIANT_Getter = $tProperty.__getter
		VariantClear($pVARIANT_Getter)
	EndIf
	$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
	VariantCopy($pVARIANT_Getter, $tVARIANT)
	$tProperty.__getter = $pVARIANT_Getter
	Return $S_OK
EndFunc

Func __AOI_Invoke_defineSetter($tObject, $pDispParams, $lcid)
	Local $iLock = $tObject.lock
	If BitAND($iLock, $__AOI_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION

	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
	$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
	If Not (($tVARIANT.vt=$VT_RECORD) Or ($tVARIANT.vt=$VT_BSTR)) Then Return $DISP_E_BADVARTYPE
	Local $tVARIANT2 = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize($tVARIANT))
	If Not ($tVARIANT2.vt=$VT_BSTR) Then Return $DISP_E_BADVARTYPE

	$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
	$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
	DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
	DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
	$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
	$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
	GetIDsOfNames($tObject.Object, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"))

	$tProperty = DllStructCreate($tagProperty, $pProperty)
	$pProperty = $tObject.Properties

	$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
	$tProperty = __AOI_PropertyGetFromId($pProperty, $t.id)
	If ($tProperty.__setter=0) Then
		Local $tVARIANT_Setter = DllStructCreate($tagVARIANT)
		Local $pVARIANT_Setter = MemCloneGlob($tVARIANT_Setter)
		VariantInit($pVARIANT_Setter)
	Else
		Local $pVARIANT_Setter = $tProperty.__setter
		VariantClear($pVARIANT_Setter)
	EndIf
	VariantCopy($pVARIANT_Setter, $tVARIANT)
	$tProperty.__setter = $pVARIANT_Setter
	Return $S_OK
EndFunc

Func MemCloneGlob($tObject);clones DllStruct to Global memory and return pointer to new allocated memory
   Local $iSize = DllStructGetSize($tObject)
   Local $hData = _MemGlobalAlloc($iSize, $GMEM_MOVEABLE)
   Local $pData = _MemGlobalLock($hData)
   _MemMoveMemory(DllStructGetPtr($tObject), $pData, $iSize)
   Return $pData
EndFunc

Func GlobalHandle($pMem)
   Local $aRet = DllCall("Kernel32.dll", "ptr", "GlobalHandle", "ptr", $pMem)
   If @error<>0 Then Return SetError(@error, @extended, 0)
   If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
   Return $aRet[0]
EndFunc

Func LocalSize($hMem)
	Local $aRet = DllCall("Kernel32.dll", "UINT", "LocalSize", "handle", $hMem)
	If @error<>0 Then Return SetError(@error, @extended, 0)
	If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc

Func LocalHandle($pMem)
	Local $aRet = DllCall("Kernel32.dll", "handle", "LocalHandle", "ptr", $pMem)
	If @error<>0 Then Return SetError(@error, @extended, 0)
	If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc

Func HeapSize($hHeap, $dwFlags, $lpMem)
	Local $aRet = DllCall("Kernel32.dll", "ULONG_PTR", "HeapSize", "handle", $hHeap, "dword", $dwFlags, "LONG_PTR", $lpMem)
	If @error<>0 Then Return SetError(@error, @extended, 0)
	If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc

Func GetProcessHeap()
	Local $aRet = DllCall("Kernel32.dll", "handle", "GetProcessHeap")
	If @error<>0 Then Return SetError(@error, @extended, 0)
	If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc

Func VariantInit($tVARIANT)
	Local $aRet=DllCall("OleAut32.dll","LONG","VariantInit",IsDllStruct($tVARIANT)?"struct*":"PTR",$tVARIANT)
	If @error<>0 Then Return SetError(-1, @error, 0)
	If $aRet[0]<>$S_OK Then SetError($aRet, 0, $tVARIANT)
	Return 1
EndFunc

Func VariantCopy($tVARIANT_Dest,$tVARIANT_Src)
	Local $aRet=DllCall("OleAut32.dll","LONG","VariantCopy",IsDllStruct($tVARIANT_Dest)?"struct*":"PTR",$tVARIANT_Dest, IsDllStruct($tVARIANT_Src)?"struct*":"PTR", $tVARIANT_Src)
	If @error<>0 Then Return SetError(-1, @error, 0)
	If $aRet[0]<>$S_OK Then SetError($aRet, 0, 0)
	Return 1
EndFunc

Func VariantClear($tVARIANT)
	Local $aRet=DllCall("OleAut32.dll","LONG","VariantClear",IsDllStruct($tVARIANT)?"struct*":"PTR",$tVARIANT)
	If @error<>0 Then Return SetError(-1, @error, 0)
	If $aRet[0]<>$S_OK Then SetError($aRet, 0, $tVARIANT)
	Return 1
EndFunc

Func VariantChangeType()
	;TODO
EndFunc

Func VariantChangeTypeEx()
	;TODO
EndFunc

Func PrivateProperty()
	Return SetError(1, 1, 0)
EndFunc

#cs
# @internal
#ce
Func __AOI_PropertyGetFromName($pProperty, $psName, $bCase = True)
	Local $iID = -1, $iIndex=-1, $tProperty
	While 1
		If $pProperty=0 Then ExitLoop
		$iIndex+=1
		$tProperty = DllStructCreate($tagProperty, $pProperty)
		If __AOI_StrCmp($tProperty.Name, $psName, $bCase) = 0 Then
			$iID = $iIndex
			ExitLoop
		EndIf
		If $tProperty.next=0 Then ExitLoop
		$pProperty = $tProperty.Next
	WEnd
	If $iID=-1 Then Return SetError(1, $iIndex, $pProperty)
	Return SetError(0, $iID, $pProperty)
EndFunc

#cs
# @internal
#ce
Func __AOI_PropertyGetFromId($pProperty, $iID)
	Local $tProperty = DllStructCreate($tagProperty, $pProperty)
	For $i=1 To $iID - 1
		$tProperty = DllStructCreate($tagProperty, $tProperty.Next)
	Next
	Return $tProperty
EndFunc

#cs
# Create new empty named property
# @internal
# @param String $sName new property name
# @return Pointer new empty property
#ce
Func __AOI_PropertyCreate($sName)
	Local $tProp = DllStructCreate($tagProperty)
		$tProp.Name = _WinAPI_CreateString($sName)
			$tVARIANT = DllStructCreate($tagVARIANT)
			$pVARIANT = MemCloneGlob($tVARIANT)
			$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
			$tVARIANT.vt = $VT_EMPTY
			VariantInit($tVARIANT)
		$tProp.Variant = $pVARIANT
	Return MemCloneGlob($tProp)
EndFunc

#cs
# Get IDispatch property pointer offset from Object property in bytes
# @internal
# @param String $sElement struct element name
# @return Integer offset
#ce
Func __AOI_GetPtrOffset($sElement)
	Local Static $tObject = DllStructCreate($__AOI_tagObject), $iObject = Int(DllStructGetPtr($tObject, "Object"), @AutoItX64?2:1)
	Return DllStructGetPtr($tObject, $sElement) - $iObject
EndFunc

#cs
# Gets single value from pointer
# @internal
# @param Pointer $pPointer the struct element pointer
# @param String $sElementType struct element specification
# @return Mixed the value of the pointer of type specified by $sElementType
#ce
Func __AOI_GetPtrValue($pPointer, $sElementType)
	Return DllStructGetData(DllStructCreate($sElementType, $pPointer), 1)
EndFunc

Func __AOI_StrCmp($pString1, $pString2, $bCase = False, $bUnicode = True)
	Local $aRet = DllCall('kernel32.dll', 'int', 'lstrcmp' & ($bCase ? '' : 'i') & ($bUnicode ? 'W' : ''), 'ptr', $pString1, 'ptr', $pString2)
	If @error Then Return SetError(@error, @extended, Null)

	Return $aRet[0]
EndFunc
