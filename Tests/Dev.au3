Func GetIDsOfNames2($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
	Local $tIds = DllStructCreate("long", $rgDispId); 2,147,483,647 properties available to define, per object. And 2,147,483,647 private properties to set in the negative space, per object.
	Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
	Local $tProperty = 0
	Local $iSize2

	Local $pStr = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)
	Local $iSize = _WinAPI_StrLen($pStr, True)
	Local $t_rgszNames = DllStructCreate("WCHAR["&$iSize&"]", $pStr)
	Local $s_rgszName = DllStructGetData($t_rgszNames, 1)

	;TODO: implement for loop, to allow processing of each Name provided

;~ 	If $s_rgszName == "" Then
		;look into possible toString emulation
;~ 	ElseIf StringLeft($s_rgszName, 2) == "__" Then;to try and improve speed
		Switch $s_rgszName
			Case "__assign"
				DllStructSetData($tIds, 1, -2)
;~ 			Case "__clone"
;~ 				DllStructSetData($tIds, 1, -3)
			Case "__case"
				DllStructSetData($tIds, 1, -4)
			Case "__freeze"
				DllStructSetData($tIds, 1, -5)
			Case "__isFrozen"
				DllStructSetData($tIds, 1, -6)
			Case "__isSealed"
				DllStructSetData($tIds, 1, -7)
			Case "__keys"
				DllStructSetData($tIds, 1, -8)
			Case "__preventExtensions"
				DllStructSetData($tIds, 1, -9)
			Case "__defineGetter"
				DllStructSetData($tIds, 1, -10)
			Case "__defineSetter"
				DllStructSetData($tIds, 1, -11)
			Case "__lookupGetter"
				DllStructSetData($tIds, 1, -12)
			Case "__lookupSetter"
				DllStructSetData($tIds, 1, -13)
			Case "__seal"
				DllStructSetData($tIds, 1, -14)
;~ 			Case "__values"
;~ 				DllStructSetData($tIds, 1, -15)
			Case "__destructor"
				DllStructSetData($tIds, 1, -16)
			Case "__unset"
				DllStructSetData($tIds, 1, -17)
			Case "__get"
				DllStructSetData($tIds, 1, -18)
			Case "__set"
				DllStructSetData($tIds, 1, -19)
			Case Else
				DllStructSetData($tIds, 1, -1)
;~ 				Return $DISP_E_UNKNOWNNAME
		EndSwitch
		If DllStructGetData($tIds, 1) <> -1 Then Return $S_OK
;~ 	EndIf

	Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
	Local $bCase = Not (BitAND($iLock, $__AutoItObjectInternal_LOCK_CASE)>0)
	Local $pProperty = __AutoItObjectInternal_PropertyGetFromName(__AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("Properties"), "ptr"), DllStructGetData($t_rgszNames, 1), $bCase)
	Local $iID = @error<>0?-1:@extended
	Local $iIndex = @extended

	If ($iID==-1) And BitAND($iLock, $__AutoItObjectInternal_LOCK_CREATE)=0 Then
		Local $pData = __AutoItObjectInternal_PropertyCreate(DllStructGetData($t_rgszNames, 1))
		If $iIndex=-1 Then;first item in list
			DllStructSetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")), 1, $pData)
		Else
			$tProperty = DllStructCreate($tagProperty, $pProperty)
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

	Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
	Local $tProperty = DllStructCreate($tagProperty, $pProperty)

	If $dispIdMember<-1 Then
		If $dispIdMember=-4 Then;__case
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If (Not(BitAND($wFlags, $DISPATCH_PROPERTYGET)=0)) Then
				If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
				$tVARIANT=DllStructCreate($tagVARIANT, $pVarResult)
				$tVARIANT.vt = $VT_BOOL
				Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
				$tVARIANT.data = (BitAND($iLock, $__AutoItObjectInternal_LOCK_CASE)>0)?0:1
			Else; $DISPATCH_PROPERTYPUT
				If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
				$tVARIANT=DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
				If $tVARIANT.vt<>$VT_BOOL Then Return $DISP_E_BADVARTYPE
				Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
				If BitAND($iLock, $__AutoItObjectInternal_LOCK_UPDATE)>0 Then Return $DISP_E_EXCEPTION
				Local $tLock = DllStructCreate("BYTE", $pSelf + __AutoItObjectInternal_GetPtrOffset("lock"))
				$b = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
				DllStructSetData($tLock, 1, _
					(Not $b) ? BitOR(DllStructGetData($tLock, 1), $__AutoItObjectInternal_LOCK_CASE) : BitAND(DllStructGetData($tLock, 1), BitNOT(BitShift(1 , 0-(Log($__AutoItObjectInternal_LOCK_CASE)/log(2))))) _
				)

			EndIf
			Return $S_OK
		EndIf

		If $dispIdMember=-13 Then;__lookupSetter
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT

			$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
			DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
			DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs)
			$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
			$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
			If Not GetIDsOfNames2($pSelf, $riid, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) == $S_OK Then Return $DISP_E_EXCEPTION

			$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
			$tProperty = __AutoItObjectInternal_PropertyGetFromId($pProperty, $t.id)
			If Not $tProperty.__setter=0 Then
				VariantClear($pVarResult)
				VariantCopy($pVarResult, $tProperty.__setter)
			EndIf
			Return $S_OK
		EndIf

		If $dispIdMember=-12 Then;__lookupGetter
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT

			$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
			DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
			DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs)
			$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
			$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
			If Not GetIDsOfNames2($pSelf, $riid, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) == $S_OK Then Return $DISP_E_EXCEPTION

			$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
			$tProperty = __AutoItObjectInternal_PropertyGetFromId($pProperty, $t.id)
			If Not $tProperty.__getter=0 Then
				VariantClear($pVarResult)
				VariantCopy($pVarResult, $tProperty.__getter)
			EndIf
			Return $S_OK
		EndIf

		If $dispIdMember=-2 Then;__assign
			Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
			If BitAND($iLock, $__AutoItObjectInternal_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION


			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs=0 Then Return $DISP_E_BADPARAMCOUNT

			;FIXME
			Local $tVARIANT = DllStructCreate($tagVARIANT)
			Local $iVARIANT = DllStructGetSize($tVARIANT)
			Local $i
			Local $pExternalProperty, $tExternalProperty
			Local $pProperty, $tProperty
			Local $iID, $iIndex, $pData
			For $i=$tDISPPARAMS.cArgs-1 To 0 Step -1
				$tVARIANT=DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+$iVARIANT*$i)
				If Not (DllStructGetData($tVARIANT, "vt")==$VT_DISPATCH) Then Return $DISP_E_BADVARTYPE
				$pExternalProperty = __AutoItObjectInternal_GetPtrValue(DllStructGetData($tVARIANT, "data") + __AutoItObjectInternal_GetPtrOffset("Properties"), "ptr")
				While 1
					ConsoleWrite($pExternalProperty&@CRLF)
					If $pExternalProperty = 0 Then ExitLoop
					$tExternalProperty = DllStructCreate($tagProperty, $pExternalProperty)
					ConsoleWrite(_WinAPI_GetString($tExternalProperty.Name)&@CRLF)

					$pProperty = __AutoItObjectInternal_PropertyGetFromName(__AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("Properties"), "ptr"), _WinAPI_GetString($tExternalProperty.Name), False)
					$iID = @error<>0?-1:@extended
					$iIndex = @extended

					If ($iID==-1) Then
						$pData = __AutoItObjectInternal_PropertyCreate(_WinAPI_GetString($tExternalProperty.Name))
						$tProperty = DllStructCreate($tagProperty, $pData)
						VariantClear($tProperty.Variant)
						VariantCopy($tProperty.Variant, $tExternalProperty.Variant)

						If $iIndex=-1 Then;first item in list
							DllStructSetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")), 1, $pData)
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
		EndIf

		If $dispIdMember=-7 Then;__isSealed
			Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
			Local $iSeal = $__AutoItObjectInternal_LOCK_CREATE + $__AutoItObjectInternal_LOCK_DELETE
			$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
			$tVARIANT.vt = $VT_BOOL
			$tVARIANT.data = (BitAND($iLock, $iSeal) = $iSeal)?1:0
			Return $S_OK
		EndIf

		If $dispIdMember=-6 Then;__isFrozen
			Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
			Local $iFreeze = $__AutoItObjectInternal_LOCK_CREATE + $__AutoItObjectInternal_LOCK_UPDATE + $__AutoItObjectInternal_LOCK_DELETE
			$tVARIANT = DllStructCreate($tagVARIANT, $pVarResult)
			$tVARIANT.vt = $VT_BOOL
			$tVARIANT.data = (BitAND($iLock, $iFreeze) = $iFreeze)?1:0
			Return $S_OK
		EndIf

		If $dispIdMember=-18 Then;__get
			;TODO: check number of parameters == 1 and the parameter type is bstr
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
			DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
			DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
			$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs), "data")
			$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
			If Not GetIDsOfNames2($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) = $S_OK Then Return $DISP_E_EXCEPTION
			Return Invoke2($pSelf, $t.id, $riid, $lcid, $DISPATCH_PROPERTYGET, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
		EndIf

		If $dispIdMember=-19 Then;__set
			;TODO: check number of parameters == 2 and parameter 1 type is bstr
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
			DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
			DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
			$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
			$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
			If Not GetIDsOfNames2($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id")) = $S_OK Then Return $DISP_E_EXCEPTION
			Return Invoke2($pSelf, $t.id, $riid, $lcid, $DISPATCH_PROPERTYPUT, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
		EndIf

		If $dispIdMember=-16 Then;__destructor
			Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
			If BitAND($iLock, $__AutoItObjectInternal_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>1 Then Return $DISP_E_BADPARAMCOUNT
			If Not (DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs),"vt")==$VT_RECORD) Then Return $DISP_E_BADVARTYPE
			Local $tVARIANT = DllStructCreate($tagVARIANT)
			Local $pVARIANT = MemCloneGlob($tVARIANT)
			$tVARIANT = DllStructCreate($tagVARIANT, $pVARIANT)
			VariantInit($pVARIANT)
			VariantCopy($pVARIANT, $tDISPPARAMS.rgvargs)
			DllStructSetData(DllStructCreate("PTR", $pSelf + __AutoItObjectInternal_GetPtrOffset("__destructor")),1,$pVARIANT)
			Return $S_OK
		EndIf

		If $dispIdMember=-5 Then;__freeze
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
			Local $tLock = DllStructCreate("BYTE", $pSelf + __AutoItObjectInternal_GetPtrOffset("lock"))
			DllStructSetData($tLock, 1, BitOR(DllStructGetData($tLock, 1), $__AutoItObjectInternal_LOCK_CREATE + $__AutoItObjectInternal_LOCK_DELETE + $__AutoItObjectInternal_LOCK_UPDATE))
			Return $S_OK
		EndIf

		If $dispIdMember=-14 Then;__seal
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
			Local $tLock = DllStructCreate("BYTE", $pSelf + __AutoItObjectInternal_GetPtrOffset("lock"))
			DllStructSetData($tLock, 1, BitOR(DllStructGetData($tLock, 1), $__AutoItObjectInternal_LOCK_CREATE + $__AutoItObjectInternal_LOCK_DELETE))
			Return $S_OK
		EndIf

		If $dispIdMember=-9 Then;__preventExtensions
			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>0 Then Return $DISP_E_BADPARAMCOUNT
			Local $tLock = DllStructCreate("BYTE", $pSelf + __AutoItObjectInternal_GetPtrOffset("lock"))
			DllStructSetData($tLock, 1, BitOR(DllStructGetData($tLock, 1), $__AutoItObjectInternal_LOCK_CREATE))
			Return $S_OK
		EndIf

		If $dispIdMember=-17 Then;__unset
			;TODO: remove getter/setter if !=0
			Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
			If BitAND($iLock, $__AutoItObjectInternal_LOCK_DELETE)>0 Then Return $DISP_E_EXCEPTION

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
						DllStructSetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")), 1, $tProperty.Next)
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

		If ($dispIdMember=-8) Then;__keys
			Local $aKeys[1]
			Local $pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
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
			VariantCopy($pVarResult, DllStructGetData(DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + __AutoItObjectInternal_GetPtrOffset("Properties")),1)), "Variant"))
			$oIDispatch=0
			Return $S_OK
		EndIf

		If ($dispIdMember=-10) Then;__defineGetter
			Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
			If BitAND($iLock, $__AutoItObjectInternal_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
			$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
			If Not (($tVARIANT.vt==$VT_RECORD) Or ($tVARIANT.vt==$VT_BSTR)) Then Return $DISP_E_BADVARTYPE
			Local $tVARIANT2 = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize($tVARIANT))
			If Not ($tVARIANT2.vt==$VT_BSTR) Then Return $DISP_E_BADVARTYPE

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
			DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
			DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
			$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
			$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
			GetIDsOfNames2($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"));TODO

			$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
			$tProperty = __AutoItObjectInternal_PropertyGetFromId($pProperty, $t.id)

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
		ElseIf ($dispIdMember=-11) Then;defineSetter
			Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
			If BitAND($iLock, $__AutoItObjectInternal_LOCK_CREATE)>0 Then Return $DISP_E_EXCEPTION

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			If $tDISPPARAMS.cArgs<>2 Then Return $DISP_E_BADPARAMCOUNT
			$tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
			If Not (($tVARIANT.vt==$VT_RECORD) Or ($tVARIANT.vt==$VT_BSTR)) Then Return $DISP_E_BADVARTYPE
			Local $tVARIANT2 = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize($tVARIANT))
			If Not ($tVARIANT2.vt==$VT_BSTR) Then Return $DISP_E_BADVARTYPE

			$tDISPPARAMS = DllStructCreate($tagDISPPARAMS, $pDispParams)
			$t = DllStructCreate("ptr id_ptr;long id;ptr str_ptr_ptr;ptr str_ptr")
			DllStructSetData($t, "id_ptr", DllStructGetPtr($t, 2))
			DllStructSetData($t, "str_ptr", $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT)))
			$t.str_ptr = DllStructGetData(DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs+DllStructGetSize(DllStructCreate($tagVARIANT))), "data")
			$t.str_ptr_ptr = DllStructGetPtr($t, "str_ptr")
			GetIDsOfNames2($pSelf, 0, $t.str_ptr_ptr, 1, $lcid, DllStructGetPtr($t, "id"));TODO

			$pProperty = DllStructGetData(DllStructCreate("ptr", $pSelf + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
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
			Local $tProperty02 = DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", ptr($oIDispatch) + __AutoItObjectInternal_GetPtrOffset("Properties")),1))
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tVARIANT = DllStructCreate($tagVARIANT, $tProperty02.Variant)
			$tVARIANT.vt = $VT_DISPATCH
			$tVARIANT.data = $pSelf
			$oIDispatch.arguments = IDispatch();
			$oIDispatch.arguments.length=$tDISPPARAMS.cArgs
			Local $aArguments[$tDISPPARAMS.cArgs], $iArguments=$tDISPPARAMS.cArgs-1
			Local $_pProperty = DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
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
			Local $tProperty02 = DllStructCreate($tagProperty, DllStructGetData(DllStructCreate("ptr", ptr($oIDispatch) + __AutoItObjectInternal_GetPtrOffset("Properties")),1))
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tProperty02=DllStructCreate($tagProperty, $tProperty02.Next)
			$tVARIANT = DllStructCreate($tagVARIANT, $tProperty02.Variant)
			$tVARIANT.vt = $VT_DISPATCH
			$tVARIANT.data = $pSelf
			Local $_pProperty = DllStructGetData(DllStructCreate("ptr", Ptr($oIDispatch) + __AutoItObjectInternal_GetPtrOffset("Properties")),1)
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
			Return $S_OK
		EndIf

		Local $iLock = __AutoItObjectInternal_GetPtrValue($pSelf + __AutoItObjectInternal_GetPtrOffset("lock"), "BYTE")
		If BitAND($iLock, $__AutoItObjectInternal_LOCK_UPDATE)>0 Then Return $DISP_E_EXCEPTION

		$_tVARIANT = DllStructCreate($tagVARIANT, $tDISPPARAMS.rgvargs)
		VariantClear($tVARIANT)
		VariantCopy($tVARIANT, $_tVARIANT)
	EndIf
	Return $S_OK
EndFunc