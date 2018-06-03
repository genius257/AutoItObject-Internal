#cs ----------------------------------------------------------------------------
 AutoIt Version : 3.3.14.2
 Author.........: genius257
 Version........: 1.0.3
#ce ----------------------------------------------------------------------------

#include-once
#include <Memory.au3>
#include <WinAPISys.au3>

Global Const $IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Global Const $IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
Global Const $IID_IConnectionPointContainer = "{B196B284-BAB4-101A-B69C-00AA00341D07}"

Global Const $DISPATCH_METHOD = 1
Global Const $DISPATCH_PROPERTYGET = 2
Global Const $DISPATCH_PROPERTYPUT = 4
Global Const $DISPATCH_PROPERTYPUTREF = 8

Global Const $DISP_E_UNKNOWNINTERFACE = 0x80020001
Global Const $DISP_E_MEMBERNOTFOUND = 0x80020003
Global Const $DISP_E_PARAMNOTFOUND = 0x80020004
Global Const $DISP_E_TYPEMISMATCH = 0x80020005
Global Const $DISP_E_UNKNOWNNAME = 0x80020006
Global Const $DISP_E_NONAMEDARGS = 0x80020007
Global Const $DISP_E_BADVARTYPE = 0x80020008
Global Const $DISP_E_EXCEPTION = 0x80020009
Global Const $DISP_E_OVERFLOW = 0x8002000A
Global Const $DISP_E_BADINDEX = 0x8002000B
Global Const $DISP_E_UNKNOWNLCID = 0x8002000C
Global Const $DISP_E_ARRAYISLOCKED = 0x8002000D
Global Const $DISP_E_BADPARAMCOUNT = 0x8002000E
Global Const $DISP_E_PARAMNOTOPTIONAL = 0x8002000F
Global Const $DISP_E_BADCALLEE = 0x80020010
Global Const $DISP_E_NOTACOLLECTION = 0x80020011

Global Const $tagVARIANT = "ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2"
Global Const $tagDISPPARAMS = "ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;"

Global Enum $VT_EMPTY,$VT_NULL,$VT_I2,$VT_I4,$VT_R4,$VT_R8,$VT_CY,$VT_DATE,$VT_BSTR,$VT_DISPATCH, _
	$VT_ERROR,$VT_BOOL,$VT_VARIANT,$VT_UNKNOWN,$VT_DECIMAL,$VT_I1=16,$VT_UI1,$VT_UI2,$VT_UI4,$VT_I8, _
	$VT_UI8,$VT_INT,$VT_UINT,$VT_VOID,$VT_HRESULT,$VT_PTR,$VT_SAFEARRAY,$VT_CARRAY,$VT_USERDEFINED, _
	$VT_LPSTR,$VT_LPWSTR,$VT_RECORD=36,$VT_FILETIME=64,$VT_BLOB,$VT_STREAM,$VT_STORAGE,$VT_STREAMED_OBJECT, _
	$VT_STORED_OBJECT,$VT_BLOB_OBJECT,$VT_CF,$VT_CLSID,$VT_BSTR_BLOB=0xfff,$VT_VECTOR=0x1000, _
	$VT_ARRAY=0x2000,$VT_BYREF=0x4000,$VT_RESERVED=0x8000,$VT_ILLEGAL=0xffff,$VT_ILLEGALMASKED=0xfff, _
	$VT_TYPEMASK=0xfff

Global Const $tagProperty = "ptr Name;ptr Variant;ptr __getter;ptr __setter;ptr Next"

Func IDispatch($QueryInterface=QueryInterface, $AddRef=AddRef, $Release=Release, $GetTypeInfoCount=GetTypeInfoCount, $GetTypeInfo=GetTypeInfo, $GetIDsOfNames=GetIDsOfNames, $Invoke=Invoke)
	Local $tagObject = "int RefCount;int Size;ptr Object;ptr Methods[7];int_ptr Callbacks[7];ptr Properties;BYTE lock;PTR __destructor"
	Local $tObject = DllStructCreate($tagObject)

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

	DllStructSetData($tObject, "RefCount", 2) ; initial ref count is 1
	DllStructSetData($tObject, "Size", 7) ; number of interface methods

	Local $pData = MemCloneGlob($tObject)

	Local $tObject = DllStructCreate($tagObject, $pData)

	DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
	Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IDispatch, Default, True) ; pointer that's wrapped into object
EndFunc

Func QueryInterface($pSelf, $pRIID, $pObj)
	If $pObj=0 Then Return $E_POINTER
	Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
	If (Not ($sGUID=$IID_IDispatch)) And (Not ($sGUID=$IID_IUnknown)) Then Return $E_NOINTERFACE
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
	Local $i
	Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
	$tStruct.Ref -= 1
	If $tStruct.Ref == 0 Then
		; initiate garbage collection
		Local $pDescructor = DllStructGetData(DllStructCreate("PTR", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4) + 1),1)
		Local $tVARIANT = DllStructCreate($tagVARIANT, $pDescructor)
		If Not ($pDescructor=0) Then
			DllStructSetData(DllStructCreate("PTR", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4) + 1),1,0)
			Local $IDispatch = IDispatch()
			$IDispatch.a=0
			Local $pProperty = DllStructGetData(DllStructCreate("PTR", Ptr($IDispatch) + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1)
			Local $pVARIANT = DllStructGetData(DllStructCreate($tagProperty, $pProperty),"Variant")
			VariantClear($pVARIANT)
			VariantCopy($pVARIANT, $tVARIANT)
			Local $f__destructor = $IDispatch.a
			VariantClear($pVARIANT)
			DllStructSetData(DllStructCreate("INT", $pSelf-4-4), 1, DllStructGetData(DllStructCreate("INT", $pSelf-4-4), 1)+1)
			Local $tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
			$tVARIANT.vt = $VT_DISPATCH
			$tVARIANT.data = $pSelf
			Call($f__destructor, $IDispatch.a)
			VariantClear($pVARIANT)
			$IDispatch=0
		EndIf
		DllStructSetData(DllStructCreate("BYTE", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2 + (@AutoItX64?8:4)),1,1);lock
		Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1);get first property
		DllStructSetData(DllStructCreate("ptr", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1,0);detatch properties from object
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
		Local $pMethods = $pSelf + (@AutoItX64?8:4)
		#cs
		For $i=0 To DllStructGetData(DllStructCreate("INT", $pSelf-4),1)-1
			DllStructSetData(DllStructCreate("PTR",$pMethods), 1, 0)
			$pMethods+=(@AutoItX64?8:4)
		Next
		Local $pCallbacks = $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7)
		#ce
		#cs
		;Cannot Free callback while in progress i guess... makes some sense
		For $i=0 To DllStructGetData(DllStructCreate("INT", $pSelf-4),1)-1
			DllCallbackFree(DllStructGetData(DllStructCreate("PTR",$pCallbacks),1))
			$pCallbacks+=(@AutoItX64?8:4)
		Next
		#ce
		_MemGlobalFree(GlobalHandle($pSelf-8))
		Return 0
	EndIf
	Return $tStruct.Ref
EndFunc

Func GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
	Local $tIds = DllStructCreate("long", $rgDispId)
	Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1)
	Local $tProperty = 0
	Local $iSize2

	Local $pStr = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)
	Local $iSize = _WinAPI_StrLen($pStr, True)
	Local $t_rgszNames = DllStructCreate("WCHAR["&$iSize&"]", $pStr)
	Local $s_rgszName = DllStructGetData($t_rgszNames, 1)

	If $s_rgszName=="__defineGetter" Then
		DllStructSetData($tIds, 1, -2)
		Return $S_OK
	ElseIf $s_rgszName=="__defineSetter" Then
		DllStructSetData($tIds, 1, -3)
		Return $S_OK
	ElseIf $s_rgszName=="__keys" Then
		DllStructSetData($tIds, 1, -4)
		Return $S_OK
	ElseIf $s_rgszName=="__unset" Then
		DllStructSetData($tIds, 1, -5)
		Return $S_OK
	ElseIf $s_rgszName=="__lock" Then
		DllStructSetData($tIds, 1, -6)
		Return $S_OK
	ElseIf $s_rgszName=="__destructor" Then
		DllStructSetData($tIds, 1, -7)
		Return $S_OK
	ElseIf $s_rgszName=="__register" Then
		DllStructSetData($tIds, 1, -8)
		Return $S_OK
	ElseIf $s_rgszName=="__unregister" Then
		DllStructSetData($tIds, 1, -9)
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

	Local $iLock = DllStructGetData(DllStructCreate("BYTE", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4)),1)
	If ($iID==-1) And $iLock=0 Then
		$tProp = DllStructCreate($tagProperty)
		$pData = MemCloneGlob($tProp)
		$tProp = DllStructCreate($tagProperty, $pData)
			$tProp.Name = _WinAPI_CreateString(DllStructGetData($t_rgszNames, 1))
				$tVARIANT = DllStructCreate($tagVARIANT)
				$pVARIANT = MemCloneGlob($tVARIANT)
				$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
;~ 					$tVARIANT.vt = $VT_EMPTY
				VariantInit($tVARIANT)
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
	If $iTInfo<>0 Then Return $DISP_E_BADINDEX
	If $ppTInfo=0 Then Return $E_INVALIDARG
	Return $E_NOTIMPL
	If DllStructGetData(DllStructCreate("UINT", $iTInfo),1)<>0 Then Return $DISP_E_BADINDEX
;~ 	Local $IID_ITypeInfo = "{00020401-0000-0000-C000-000000000046}"
;~ 	ConsoleWrite("$iTInfo: "&DllStructGetData(DllStructCreate("UINT", $iTInfo),1)&@CRLF)
;~ 	ConsoleWrite("$lcid: "&$lcid&@CRLF)
;~ 	Local $ITypeInfo=DllCall("OleAut32.dll","LONG","CreateDispTypeInfo","PTR",0,"UINT",$lcid,"PTR*",0)
;~ 	$ITypeInfo=ObjCreateInterface($ITypeInfo[3], $IID_IUnknown)
;~ 	$ITypeInfo=ObjCreateInterface($ITypeInfo[3], $IID_ITypeInfo, "GetTypeAttr;GetTypeComp;GetFuncDesc;GetVarDesc;GetNames;GetRefTypeOfImplType;GetImplTypeFlags;GetIDsOfNames;Invoke;GetDocumentation;GetDllEntry;GetRefTypeInfo;AddressOfMember;CreateInstance;GetMops;GetContainingTypeLib;ReleaseTypeAttr;ReleaseFuncDesc;ReleaseVarDesc")
;~ 	ConsoleWrite($ITypeInfo[3]&@CRLF)
;~ 	ConsoleWrite(DllStructSetData(DllStructCreate("PTR",$ppTInfo),1)&@CRLF)
;~ 	DllStructSetData(DllStructCreate("PTR",$ppTInfo),1,Ptr($ITypeInfo))
;~ 	$ITypeInfo=0
	Return $S_OK
EndFunc

Func GetTypeInfoCount($pSelf, $pctinfo)
;~ 	DllStructSetData(DllStructCreate("UINT",$pctinfo),1, 1)
	DllStructSetData(DllStructCreate("UINT",$pctinfo),1, 0)
	Return $S_OK
;~ 	Return $E_NOTIMPL
EndFunc

Func Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
	If $dispIdMember=-1 Then Return $DISP_E_MEMBERNOTFOUND
	Local $tVARIANT, $_tVARIANT, $tDISPPARAMS
	Local $t
	Local $i

	Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1)
	Local $tProperty = DllStructCreate($tagProperty, $pProperty)

	If $dispIdMember<-1 Then
		If $dispIdMember=-7 Then;__destructor
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
			If Not (DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs),"vt")==$VT_RECORD) Then Return $DISP_E_BADVARTYPE
			Local $tVARIANT = DllStructCreate($tagVARIANT)
			Local $pVARIANT = MemCloneGlob($tVARIANT)
			$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
			VariantInit($pVARIANT)
			VariantCopy($pVARIANT, $tDISPPARAMS.rgvargs)
			DllStructSetData(DllStructCreate("PTR", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4) + 1),1,$pVARIANT)
			Return $S_OK
		EndIf

		If $dispIdMember=-6 Then;__lock
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
			If DllStructGetData(DllStructCreate("BYTE", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4)),1)>0 Then Return $DISP_E_EXCEPTION
			DllStructSetData(DllStructCreate("BYTE", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4)),1, 1)
			Return $S_OK
		EndIf

		If $dispIdMember=-5 Then;__unset
			;TODO: remove getter/setter if !=0
			If DllStructGetData(DllStructCreate("BYTE", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4)),1)>0 Then
				Return $DISP_E_EXCEPTION
			EndIf

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
			$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
			If Not($VT_BSTR==$tVARIANT.vt) Then Return $DISP_E_BADVARTYPE
			Local $sProperty = _WinAPI_GetString($tVARIANT.data);the string to search for
			Local $tProperty=0,$tProperty_Prev
			While 1
				If $pProperty=0 Then ExitLoop
				$tProperty_Prev = $tProperty
				$tProperty = DllStructCreate($tagProperty, $pProperty)
				If _WinAPI_GetString($tProperty.Name)==$sProperty Then
					If $tProperty_Prev==0 Then
						DllStructSetData(DllStructCreate("ptr", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2), 1, $tProperty.Next)
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
		EndIf

		If ($dispIdMember=-4) Then;__keys
			Local $aKeys[1]
			Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1)
			While 1
				If $pProperty=0 Then ExitLoop
				Local $tProperty = DllStructCreate($tagProperty, $pProperty)
				$aKeys[UBound($aKeys,1)-1] = DllStructGetData(DllStructCreate("WCHAR["&_WinAPI_StrLen($tProperty.Name)&"]", $tProperty.Name), 1)
				If $tProperty.next=0 Then ExitLoop
				ReDim $aKeys[UBound($aKeys,1)+1]
				$pProperty = $tProperty.next
			WEnd
			If $pProperty=0 Then Return $S_OK
			Local $oIDispatch = IDispatch()
			$oIDispatch.a=$aKeys
			VariantClear($pVarResult)
			VariantCopy($pVarResult, DllStructGetData(DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1)), "Variant"))
			$oIDispatch=0
			Return $S_OK
		EndIf

		If ($dispIdMember=-2) Then;__defineGetter
			If DllStructGetData(DllStructCreate("BYTE", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4)),1)>0 Then
				Return $DISP_E_EXCEPTION
			EndIf

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
			$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
			If Not ($tVARIANT.vt==$VT_RECORD) Then Return $DISP_E_BADVARTYPE
			Local $tVARIANT2 = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize($tVARIANT))
			If Not ($tVARIANT2.vt==$VT_BSTR) Then Return $DISP_E_BADVARTYPE

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
			DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
			DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
			$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
			$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
			GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"));TODO

			$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))),1)
			$tProperty = DllStructCreate($tagProperty, $pProperty)

			$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
			For $i=1 To $t.id
				$pProperty = $tProperty.Next
				$tProperty = DllStructCreate($tagProperty, $pProperty)
			Next
			If ($tProperty.__getter=0) Then
				Local $tVARIANT_Getter = DllStructCreate($tagVARIANT)
				$pVARIANT_Getter = MemCloneGlob($tVARIANT_Getter)
				VariantInit($pVARIANT_Getter)
			Else
				Local $pVARIANT_Getter = $tProperty.__getter
				VariantClear($pVARIANT_Getter)
			EndIf
			VariantCopy($pVARIANT_Getter, $tVARIANT)
			$tProperty.__getter = $pVARIANT_Getter
			Return $S_OK
		ElseIf ($dispIdMember=-3) Then;defineSetter
			If DllStructGetData(DllStructCreate("BYTE", $pSelf + (@AutoItX64?8:4) + ((@AutoItX64?8:4)*7*2) + (@AutoItX64?8:4)),1)>0 Then
				Return $DISP_E_EXCEPTION
			EndIf

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
			$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
			If Not ($tVARIANT.vt==$VT_RECORD) Then Return $DISP_E_BADVARTYPE
			Local $tVARIANT2 = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize($tVARIANT))
			If Not ($tVARIANT2.vt==$VT_BSTR) Then Return $DISP_E_BADVARTYPE

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
			DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
			DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
			$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
			$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
			GetIDsOfNames($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"));TODO

			$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + DllStructGetSize(DllStructCreate("ptr Object;ptr Methods[7];int_ptr Callbacks[7]"))),1)
			$tProperty = DllStructCreate($tagProperty, $pProperty)

			$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
			For $i=1 To $t.id
				$pProperty = $tProperty.Next
				$tProperty = DllStructCreate($tagProperty, $pProperty)
			Next
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
		EndIf
		Return $DISP_E_EXCEPTION;TODO: The application needs to raise an exception. In this case, the structure passed in pexcepinfo should be filled in.
	EndIf

	For $i=1 To $dispIdMember
		$pProperty = $tProperty.Next
		$tProperty = DllStructCreate($tagProperty, $pProperty)
	Next

	$tVARIANT = DllStructCreate($tagVARIANT, $tProperty.Variant)

;~ 	If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)=0)) And (Not(BitAND($wFlags, $DISPATCH_METHOD)=0)) Then
	If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)=0)) Then
		$_tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
		If Not($tProperty.__getter = 0) Then
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			Local $oIDispatch = IDispatch()
			$oIDispatch.val = 0
			$oIDispatch.ret = 0
			DllStructSetData(DllStructCreate("INT", $pSelf-4-4), 1, DllStructGetData(DllStructCreate("INT", $pSelf-4-4), 1)+1)
			$oIDispatch.parent = 0
			Local $tProperty02 = DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", ptr($oIDispatch) + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1))
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tVARIANT = DllStructCreate($tagVARIANT, $tProperty02.Variant)
			$tVARIANT.vt = $VT_DISPATCH
			$tVARIANT.data = $pSelf
			$oIDispatch.arguments = IDispatch();
			$oIDispatch.arguments.length=$tDISPPARAMS.cArgs
			Local $aArguments[$tDISPPARAMS.cArgs], $iArguments=$tDISPPARAMS.cArgs-1
			Local $_pProperty = DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1)
			Local $_tProperty = DllStructCreate($tagProperty, $_pProperty)
			For $i=0 To $iArguments
				VariantClear($_tProperty.Variant)
				VariantCopy($_tProperty.Variant, $tDISPPARAMS.rgvargs+(($iArguments-$i)*DllStructGetSize($_tVARIANT)))
				$aArguments[$i]=$oIDispatch.val
			Next
			$oIDispatch.arguments.values=$aArguments
			$oIDispatch.arguments.__lock()
			$oIDispatch.__defineSetter("parent", PrivateProperty)
			VariantClear($_tProperty.Variant)
			VariantCopy($_tProperty.Variant, $tProperty.__getter)
			Local $fGetter = $oIDispatch.val
			VariantClear($_tProperty.Variant)
			VariantCopy($_tProperty.Variant, $tProperty.Variant)
			$oIDispatch.__lock()
			Local $mRet = Call($fGetter, $oIDispatch)
			Local $iError = @error, $iExtended = @extended
			VariantClear($tProperty.Variant)
			VariantCopy($tProperty.Variant, $_tProperty.Variant)
			$oIDispatch.ret = $mRet
			$_tProperty = DllStructCreate($tagProperty, $_tProperty.Next)
			VariantCopy($pVarResult, $_tProperty.Variant)
			$oIDispatch=0
			Return ($iError<>0)?$DISP_E_EXCEPTION:$S_OK
			Return $S_OK
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
			Local $tProperty02 = DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", ptr($oIDispatch) + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1))
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tVARIANT = DllStructCreate($tagVARIANT, $tProperty02.Variant)
			$tVARIANT.vt = $VT_DISPATCH
			$tVARIANT.data = $pSelf
			Local $_pProperty = DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + (@AutoItX64?8:4) + (@AutoItX64?8:4)*7*2),1)
			Local $_tProperty = DllStructCreate($tagProperty, $_pProperty)
			Local $_tProperty2 = DllStructCreate($tagProperty, $_tProperty.Next)
			VariantClear($_tProperty.Variant)
			VariantCopy($_tProperty.Variant, $tProperty.__setter)
			VariantClear($_tProperty2.Variant)
			VariantCopy($_tProperty2.Variant, $tDISPPARAMS.rgvargs)
			Local $fSetter = $oIDispatch.val
			VariantClear($_tProperty.Variant)
			VariantCopy($_tProperty.Variant, $tProperty.Variant)
			$oIDispatch.__lock()
			Local $mRet = Call($fSetter, $oIDispatch)
			Local $iError = @error, $iExtended = @extended
			VariantClear($tProperty.Variant)
			VariantCopy($tProperty.Variant, $_tProperty.Variant)
			$oIDispatch.ret = $mRet
			$_tProperty = DllStructCreate($tagProperty, $_tProperty.Next)
			VariantCopy($pVarResult, $_tProperty.Variant)
			$oIDispatch=0
			Return ($iError<>0)?$DISP_E_EXCEPTION:$S_OK
			Return $S_OK
		EndIf

		$_tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
		VariantClear($tVARIANT)
		VariantCopy($tVARIANT, $_tVARIANT)
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