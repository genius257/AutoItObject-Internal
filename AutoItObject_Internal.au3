#cs ----------------------------------------------------------------------------
 AutoIt Version : 3.3.14.2
 Author.........: genius257
 Version........: 0.1.1
#ce ----------------------------------------------------------------------------

#include-once
#include <Memory.au3>
#include <WinAPISys.au3>

Global Const $IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Global Const $IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"

Global Const $DISPATCH_METHOD = 1
Global Const $DISPATCH_PROPERTYGET = 2
Global Const $DISPATCH_PROPERTYPUT = 4
Global Const $DISPATCH_PROPERTYPUTREF = 8

Global Const $DISP_E_MEMBERNOTFOUND = 0x80020003
Global Const $DISP_E_UNKNOWNNAME = 0x80020006
Global Const $DISP_E_BADVARTYPE = 0x80020008

Global Const $tagVARIANT = "ushort vt;ushort r1;ushort r2;ushort r3;uint64 data"
Global Const $tagDISPPARAMS = "ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;"

Global Enum $VT_EMPTY,$VT_NULL,$VT_I2,$VT_I4,$VT_R4,$VT_R8,$VT_CY,$VT_DATE,$VT_BSTR,$VT_DISPATCH, _
	$VT_ERROR,$VT_BOOL,$VT_VARIANT,$VT_UNKNOWN,$VT_DECIMAL,$VT_I1=16,$VT_UI1,$VT_UI2,$VT_UI4,$VT_I8, _
	$VT_UI8,$VT_INT,$VT_UINT,$VT_VOID,$VT_HRESULT,$VT_PTR,$VT_SAFEARRAY,$VT_CARRAY,$VT_USERDEFINED, _
	$VT_LPSTR,$VT_LPWSTR,$VT_RECORD=36,$VT_FILETIME=64,$VT_BLOB,$VT_STREAM,$VT_STORAGE,$VT_STREAMED_OBJECT, _
	$VT_STORED_OBJECT,$VT_BLOB_OBJECT,$VT_CF,$VT_CLSID,$VT_BSTR_BLOB=0xfff,$VT_VECTOR=0x1000, _
	$VT_ARRAY=0x2000,$VT_BYREF=0x4000,$VT_RESERVED=0x8000,$VT_ILLEGAL=0xffff,$VT_ILLEGALMASKED=0xfff, _
	$VT_TYPEMASK=0xfff

Global Const $tagProperty = "ptr Name;ptr Variant;ptr __getter;ptr __setter;ptr Next"

Func IUnknown()
   Local $tagObject = "int RefCount;int Size;ptr Object;ptr Methods[3];int_ptr Callbacks[3];ulong_ptr Slots[16]" ; 16 pointer sized elements more to create space for possible private props
   Local $tObject = DllStructCreate($tagObject)

   $proc = DllCallbackRegister(QueryInterface, "LONG", "ptr;ptr;ptr")
   DllStructSetData($tObject, "Methods", DllCallbackGetPtr($proc), 1)
   DllStructSetData($tObject, "Callbacks", $proc, 1)

   $proc = DllCallbackRegister(AddRef, "dword", "PTR")
   DllStructSetData($tObject, "Methods", DllCallbackGetPtr($proc), 2)
   DllStructSetData($tObject, "Callbacks", $proc, 2)

   $proc = DllCallbackRegister(Release, "dword", "PTR")
   DllStructSetData($tObject, "Methods", DllCallbackGetPtr($proc), 3)
   DllStructSetData($tObject, "Callbacks", $proc, 3)

   DllStructSetData($tObject, "RefCount", 2) ; initial ref count is 1
   DllStructSetData($tObject, "Size", 3) ; number of interface methods

   $pData = MemCloneGlob($tObject)

   $tObject = DllStructCreate($tagObject, $pData)

   DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
   Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IUnknown, Default, True) ; pointer that's wrapped into object
EndFunc

Func IDispatch()
	Local $tagObject = "int RefCount;int Size;ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties"
	Local $tObject = DllStructCreate($tagObject)
	Local $proc
	Local Static $QueryInterface = 0
	Local Static $AddRef = 0
	Local Static $Release = 0
	Local Static $GetTypeInfoCount = 0
	Local Static $GetTypeInfo = 0
	Local Static $GetIDsOfNames = 0
	Local Static $Invoke = 0

	If $QueryInterface==0 Then $QueryInterface = DllCallbackRegister(QueryInterface, "LONG", "ptr;ptr;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($QueryInterface), 1)
	DllStructSetData($tObject, "Callbacks", $QueryInterface, 1)

	If $AddRef==0 Then $AddRef = DllCallbackRegister(AddRef, "dword", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($AddRef), 2)
	DllStructSetData($tObject, "Callbacks", $AddRef, 2)

	If $Release==0 Then $Release = DllCallbackRegister(Release, "dword", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Release), 3)
	DllStructSetData($tObject, "Callbacks", $Release, 3)

	If $GetTypeInfoCount==0 Then $GetTypeInfoCount = DllCallbackRegister(GetTypeInfoCount, "long", "ptr;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfoCount), 4)
	DllStructSetData($tObject, "Callbacks", $GetTypeInfoCount, 4)

	If $GetTypeInfo==0 Then $GetTypeInfo = DllCallbackRegister(GetTypeInfo, "long", "ptr;uint;int;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfo), 5)
	DllStructSetData($tObject, "Callbacks", $GetTypeInfo, 5)

	If $GetIDsOfNames==0 Then $GetIDsOfNames = DllCallbackRegister(GetIDsOfNames, "long", "ptr;ptr;ptr;uint;int;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetIDsOfNames), 6)
	DllStructSetData($tObject, "Callbacks", $GetIDsOfNames, 6)

	If $Invoke==0 Then $Invoke = DllCallbackRegister(Invoke, "long", "ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Invoke), 7)
	DllStructSetData($tObject, "Callbacks", $Invoke, 7)

	DllStructSetData($tObject, "RefCount", 2) ; initial ref count is 1
	DllStructSetData($tObject, "Size", 7) ; number of interface methods

	Local $pData = MemCloneGlob($tObject)

	Local $tObject = DllStructCreate($tagObject, $pData)

	DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
	Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IDispatch, Default, True) ; pointer that's wrapped into object
EndFunc

Func __IDispatch();create IDispatch object and alter core functions to specialized debug alternets
	Local $oIDispatch = IDispatch()
	Local $tPtr = DllStructCreate("ptr")
	DllStructSetData($tPtr, 1, $oIDispatch)
	Local $pIDispatch = DllStructGetData($tPtr, 1)
	Local $_tIDispatch = DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]", $pIDispatch)

	Local Static $_GetIDsOfNames = 0
	Local Static $_Invoke = 0

	If $_GetIDsOfNames==0 Then $_GetIDsOfNames = DllCallbackRegister(__GetIDsOfNames, "long", "ptr;ptr;ptr;uint;int;ptr")
	DllStructSetData($_tIDispatch, "Methods", DllCallbackGetPtr($_GetIDsOfNames), 6)
	DllStructSetData($_tIDispatch, "Callbacks", $_GetIDsOfNames, 6)
	If $_Invoke==0 Then $_Invoke = DllCallbackRegister(__Invoke, "long", "ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr")
	DllStructSetData($_tIDispatch, "Methods", DllCallbackGetPtr($_Invoke), 7)
	DllStructSetData($_tIDispatch, "Callbacks", $_Invoke, 7)

	Return $oIDispatch
EndFunc

Func QueryInterface($pSelf, $pRIID, $pObj)
	Local $tStruct = DllStructCreate("ptr", $pObj)
	DllStructSetData($tStruct, 1, $pSelf)
	Return $S_OK
EndFunc

Func AddRef($pSelf)
   Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
   $tStruct.Ref += 1
   Return $tStruct.Ref
EndFunc

Func Release($pSelf)
	Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
	$tStruct.Ref -= 1
	If $tStruct.Ref == 0 Then
		_MemGlobalFree(GlobalHandle($pSelf-8))
		Return 0
	EndIf
   Return $tStruct.Ref
EndFunc

Func __GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
	Local $tIds = DllStructCreate("long", $rgDispId)
	Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))),1)
	If $pProperty=0 Then
		$tProp = DllStructCreate($tagProperty)
			$pData = MemCloneGlob($tProp)
			$tProp = DllStructCreate($tagProperty, $pData)
				$tProp.Name = 0;no reason to wase memory on a string in this object
					$tVARIANT = DllStructCreate($tagVARIANT)
					$pVARIANT = MemCloneGlob($tVARIANT)
					$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
						$tVARIANT.vt = $VT_EMPTY
				$tProp.Variant = $pVARIANT
		DllStructSetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))), 1, $pData)
	EndIf

	DllStructSetData($tIds, 1, 0)
EndFunc; __GetIDsOfNames

Func GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
	Local $tIds = DllStructCreate("long", $rgDispId)
	Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))),1)
	Local $tProperty = 0
	Local $iSize2

	Local $pStr = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)
	Local $iSize = _WinAPI_StrLen($pStr, True)
	Local $t_rgszNames = DllStructCreate("WCHAR["&$iSize&"]", $pStr)

	If DllStructGetData($t_rgszNames, 1)=="__defineGetter" Then
		DllStructSetData($tIds, 1, -2)
		Return $S_OK
	ElseIf DllStructGetData($t_rgszNames, 1)=="__defineSetter" Then
		DllStructSetData($tIds, 1, -3)
		Return $S_OK
	EndIf

	Local $iID = -1
	Local $iIndex=-1
	While 1
		If $pProperty=0 Then ExitLoop
		$iIndex+=1
		$tProperty = DllStructCreate($tagProperty, $pProperty)
		$iSize2 = _WinAPI_StrLen($tProperty.Name, True)
		If ($iSize2=$iSize) And (DllStructGetData(DllStructCreate("WCHAR["&$iSize2&"]", $tProperty.Name), 1)==DllStructGetData($t_rgszNames, 1)) Then
			$iID = $iIndex
			ExitLoop
		EndIf
		If $tProperty.next=0 Then ExitLoop
		$pProperty = $tProperty.next
	WEnd
	Local $tProp, $pData, $tVARIANT, $pVARIANT

	If $iID==-1 Then
		$tProp = DllStructCreate($tagProperty)
		$pData = MemCloneGlob($tProp)
		$tProp = DllStructCreate($tagProperty, $pData)
			$tProp.Name = _WinAPI_CreateString(DllStructGetData($t_rgszNames, 1))
				$tVARIANT = DllStructCreate($tagVARIANT)
				$pVARIANT = MemCloneGlob($tVARIANT)
				$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
					$tVARIANT.vt = $VT_EMPTY
			$tProp.Variant = $pVARIANT
		If $tProperty=0 Then;first item in list
			DllStructSetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))), 1, $pData)
		Else
			$tProperty.next = $pData
		EndIf
		$iID = $iIndex+1
	EndIf

	If $iID==-1 Then Return $DISP_E_UNKNOWNNAME
	DllStructSetData($tIds, 1, $iID)
	Return $S_OK
EndFunc

Func GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
   Return $E_NOTIMPL
EndFunc

Func GetTypeInfoCount($pSelf, $pctinfo)
   Return $E_NOTIMPL
EndFunc

Func __Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
	If $dispIdMember<>0 Then Return $DISP_E_MEMBERNOTFOUND
	Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))),1)
	Local $tProperty = DllStructCreate($tagProperty, $pProperty)

	Local $tVARIANT = DllStructCreate($tagVARIANT, $tProperty.Variant)

	If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)==0)) And (Not(BitAND($wFlags, $DISPATCH_METHOD)==0)) Then
		$_tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
		$_tVARIANT.vt = $tVARIANT.vt
		$_tVARIANT.data = $tVARIANT.data
	Else
		$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
		$_tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
		$tVARIANT.vt = $_tVARIANT.vt
		$tVARIANT.data = ($tVARIANT.vt = $VT_BSTR)?_WinAPI_CreateString(DllStructGetData(DllStructCreate("WCHAR["&_WinAPI_StrLen($_tVARIANT.data, True)&"]", $_tVARIANT.data), 1)):$_tVARIANT.data
	EndIf
EndFunc

Func Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
	If $dispIdMember=-1 Then Return $DISP_E_MEMBERNOTFOUND
	Local $tVARIANT, $_tVARIANT, $tDISPPARAMS
	Local $t

	Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))),1)
	Local $tProperty = DllStructCreate($tagProperty, $pProperty)

	If ($dispIdMember=-2) Or ($dispIdMember=-3) Then
		$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
		$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
		DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
		DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
		$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
		$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
		GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"))

		$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))),1)
		$tProperty = DllStructCreate($tagProperty, $pProperty)

		If $tDISPPARAMS.cArgs<>2 Then Return $S_OK
		$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
		For $i=1 To $t.id
			$pProperty = $tProperty.Next
			$tProperty = DllStructCreate($tagProperty, $pProperty)
		Next
		If Not($tVARIANT.vt = $VT_I4) Then Return $DISP_E_BADVARTYPE;$VT_I4 should be the type returned from DllCallbackRegister
		If ($dispIdMember=-2) Then $tProperty.__getter = $tVARIANT.data
		If ($dispIdMember=-3) Then $tProperty.__setter = $tVARIANT.data
		Return $S_OK
	EndIf

	For $i=1 To $dispIdMember
		$pProperty = $tProperty.Next
		$tProperty = DllStructCreate($tagProperty, $pProperty)
	Next

	Local $tVARIANT = DllStructCreate($tagVARIANT, $tProperty.Variant)

	If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)==0)) And (Not(BitAND($wFlags, $DISPATCH_METHOD)==0)) Then
		$_tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
		If Not($tProperty.__getter = 0) Then
			Local $oIDispatch = __IDispatch()
			Local $_oIDispatch = __IDispatch()
			$_oIDispatch.a = "";create data entry
			Local $tPtr = DllStructCreate("ptr ptr")
			$tPtr.ptr = $_oIDispatch
			Local $__tIDispatch = DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties", $tPtr.ptr)
			Local $__tProperty = DllStructCreate($tagProperty, $__tIDispatch.Properties)
			Local $__tVARIANT = DllStructCreate($tagVARIANT, $__tProperty.VARIANT)
			$__tVARIANT.vt = $tVARIANT.vt
			$__tVARIANT.data = $tVARIANT.data

			DllCallAddress("long", DllCallbackGetPtr($tProperty.__getter), "IDispatch", $oIDispatch, "IDispatch", $_oIDispatch)

			$tVARIANT.vt = $__tVARIANT.vt
			$tVARIANT.data = $__tVARIANT.data

			$tPtr.ptr = $oIDispatch
			$__tIDispatch = DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties", $tPtr.ptr)
			If Not($__tIDispatch.Properties=0) Then
				Local $__tProperty = DllStructCreate($tagProperty, $__tIDispatch.Properties)
				Local $__tVARIANT = DllStructCreate($tagVARIANT, $__tProperty.Variant)
				$_tVARIANT.vt = $__tVARIANT.vt
				$_tVARIANT.data = $__tVARIANT.data
			EndIf
			Return $S_OK
		EndIf

		$_tVARIANT.vt = $tVARIANT.vt
		$_tVARIANT.data = $tVARIANT.data
	Else
		$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
		$_tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)

		If Not($tProperty.__setter=0) Then
			Local $oIDispatch = __IDispatch()
			Local $_oIDispatch = __IDispatch()
			$_oIDispatch.a = "";create data entry
			Local $tPtr = DllStructCreate("ptr ptr")
			$tPtr.ptr = $_oIDispatch
			Local $__tIDispatch = DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties", $tPtr.ptr)
			Local $__tProperty = DllStructCreate($tagProperty, $__tIDispatch.Properties)
			Local $__tVARIANT = DllStructCreate($tagVARIANT, $__tProperty.VARIANT)
			$__tVARIANT.vt = $_tVARIANT.vt
			$__tVARIANT.data = $_tVARIANT.data

			$oIDispatch.a = "";create data entry
			$tPtr.ptr = $oIDispatch
			$__tIDispatch = DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties", $tPtr.ptr)
			$__tProperty = DllStructCreate($tagProperty, $__tIDispatch.Properties)
			$__tVARIANT = DllStructCreate($tagVARIANT, $__tProperty.VARIANT)
			$__tVARIANT.vt = $tVARIANT.vt
			$__tVARIANT.data = $tVARIANT.data

			DllCallAddress("long", DllCallbackGetPtr($tProperty.__setter), "IDispatch", $oIDispatch, "IDispatch", $_oIDispatch)
			$__tProperty = DllStructCreate($tagProperty, $__tIDispatch.Properties)
			$__tVARIANT = DllStructCreate($tagVARIANT, $__tProperty.Variant)
			$tVARIANT.vt = $__tVARIANT.vt
			$tVARIANT.data = $__tVARIANT.data
;~ 			EndIf
			Return $S_OK
		EndIf

		$tVARIANT.vt = $_tVARIANT.vt
		$tVARIANT.data = ($tVARIANT.vt = $VT_BSTR)?_WinAPI_CreateString(DllStructGetData(DllStructCreate("WCHAR["&_WinAPI_StrLen($_tVARIANT.data, True)&"]", $_tVARIANT.data), 1)):$_tVARIANT.data
	EndIf
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
   $aRet = DllCall("Kernel32.dll", "ptr", "GlobalHandle", "ptr", $pMem)
   If @error<>0 Then Return SetError(@error, @extended, 0)
   If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
   Return $aRet[0]
EndFunc

Func LocalSize($hMem)
	$aRet = DllCall("Kernel32.dll", "UINT", "LocalSize", "handle", $hMem)
	If @error<>0 Then Return SetError(@error, @extended, 0)
	If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc

Func LocalHandle($pMem)
	$aRet = DllCall("Kernel32.dll", "handle", "LocalHandle", "ptr", $pMem)
	If @error<>0 Then Return SetError(@error, @extended, 0)
	If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc

Func HeapSize($hHeap, $dwFlags, $lpMem)
	$aRet = DllCall("Kernel32.dll", "ULONG_PTR", "HeapSize", "handle", $hHeap, "dword", $dwFlags, "LONG_PTR", $lpMem)
	If @error<>0 Then Return SetError(@error, @extended, 0)
	If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc

Func GetProcessHeap()
	$aRet = DllCall("Kernel32.dll", "handle", "GetProcessHeap")
	If @error<>0 Then Return SetError(@error, @extended, 0)
	If $aRet[0]=0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc
