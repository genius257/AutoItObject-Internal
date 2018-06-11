#include "..\AutoItObject_Internal.au3"
#include "Unit\assert.au3"

#AutoIt3Wrapper_Run_Au3Check=N

;~ $ADO = ObjCreate("ADODB.Connection")
;~ Exit MsgBox(0, "", @error)


$oIDispatch = IDispatch()

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
assertInternalType($oIDispatch.a, "Ptr");test will fail, AutoIt related, not a fault of the library
$oIDispatch.a = HWnd(1)
assertEquals(@error, 0)
assertInternalType($oIDispatch.a, "HWnd");test will fail, AutoIt related, not a fault of the library'
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
Dim $a = [1]
$oIDispatch.a = $a
assertInternalType($oIDispatch.a, "Array")



;~ assertSame("test", "test2", "message")
;~ assertInternalType(1, "String")
;~ assertNotInternalType("t", "String")

Exit

$oIDispatch = IDispatch(QueryInterface, AddRef, Release, GetTypeInfoCount, GetTypeInfo, GetIDsOfNames2, Invoke2)

Func GetIDsOfNames2($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
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

Func Invoke2($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
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

assertNull($a)

assertNull($a)

assertNull($a)

$oIDispatch.__assign()
$oIDispatch.__create()
$oIDispatch.__defineProperty()
$oIDispatch.__defineProperties()
$oIDispatch.__entries()
$oIDispatch.__freeze()
$oIDispatch.__getOwnPropertyDescriptor()
$oIDispatch.__getOwnPropertyDescriptors()
$oIDispatch.__getOwnPropertyNames()
$oIDispatch.__isExtensible()
$oIDispatch.__isFrozen()
$oIDispatch.__isSealed()
$oIDispatch.__keys()
$oIDispatch.__preventExtensions()
$oIDispatch.__seal()
$oIDispatch.__values()

;unrelated:

$oIDispatch.__clone(); Return IDispatch().assign($oThis.self); Maybe also set parameters in IDispatch from old object