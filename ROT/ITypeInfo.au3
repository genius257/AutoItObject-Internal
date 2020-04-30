#include-once
#include "ITypeLib.au3"

;INT after ULONG is unsure
Const $tagTYPEATTR = "DWORD;WORD;WORD;BYTE[8];INT;DWORD;INT;INT;PTR;ULONG;INT;WORD;WORD;WORD;WORD;WORD;WORD;WORD;WORD;PTR;PTR;DWORD;USHORT;ULONG;USHORT"; https://msdn.microsoft.com/en-us/library/windows/desktop/ms221003(v=vs.85).aspx

Func ITypeInfo(); https://msdn.microsoft.com/en-us/library/cc237759.aspx
	Local $tagObject = "int RefCount;int Size;ptr Object;ptr Methods[22];int_ptr Callbacks[22];"
	Local $tObject = DllStructCreate($tagObject)

	Local $QueryInterface = DllCallbackRegister(ITypeInfo_QueryInterface, "LONG", "ptr;ptr;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($QueryInterface), 1)
	DllStructSetData($tObject, "Callbacks", $QueryInterface, 1)

	Local $AddRef = DllCallbackRegister(ITypeInfo_AddRef, "dword", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($AddRef), 2)
	DllStructSetData($tObject, "Callbacks", $AddRef, 2)

	Local $Release = DllCallbackRegister(ITypeInfo_Release, "dword", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Release), 3)
	DllStructSetData($tObject, "Callbacks", $Release, 3)

	Local $GetTypeAttr = DllCallbackRegister(ITypeInfo_GetTypeAttr, "LONG", "PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeAttr), 4)
	DllStructSetData($tObject, "Callbacks", $GetTypeAttr, 4)

	Local $GetTypeComp = DllCallbackRegister(ITypeInfo_GetTypeComp, "LONG", "PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeComp), 5)
	DllStructSetData($tObject, "Callbacks", $GetTypeComp, 5)

	Local $GetFuncDesc = DllCallbackRegister(ITypeInfo_GetFuncDesc, "LONG", "PTR;UINT;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetFuncDesc), 6)
	DllStructSetData($tObject, "Callbacks", $GetFuncDesc, 6)

	Local $GetVarDesc = DllCallbackRegister(ITypeInfo_GetVarDesc, "LONG", "PTR;UINT;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetVarDesc), 7)
	DllStructSetData($tObject, "Callbacks", $GetVarDesc, 7)

	Local $GetNames = DllCallbackRegister(ITypeInfo_GetNames, "LONG", "PTR;LONG;PTR;UINT;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetNames), 8)
	DllStructSetData($tObject, "Callbacks", $GetNames, 8)

	Local $GetRefTypeOfImplType = DllCallbackRegister(ITypeInfo_GetRefTypeOfImplType, "LONG", "PTR;UINT;DWORD")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetRefTypeOfImplType), 9)
	DllStructSetData($tObject, "Callbacks", $GetRefTypeOfImplType, 9)

	Local $GetImplTypeFlags = DllCallbackRegister(ITypeInfo_GetImplTypeFlags, "LONG", "PTR;UINT;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetImplTypeFlags), 10)
	DllStructSetData($tObject, "Callbacks", $GetImplTypeFlags, 10)

	Local $Opnum10NotUsedOnWire = DllCallbackRegister(ITypeInfo_Opnum10NotUsedOnWire, "LONG", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Opnum10NotUsedOnWire), 11)
	DllStructSetData($tObject, "Callbacks", $Opnum10NotUsedOnWire, 11)

	Local $Opnum11NotUsedOnWire = DllCallbackRegister(ITypeInfo_Opnum11NotUsedOnWire, "LONG", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Opnum11NotUsedOnWire), 12)
	DllStructSetData($tObject, "Callbacks", $Opnum11NotUsedOnWire, 12)

	Local $GetDocumentation = DllCallbackRegister(ITypeInfo_GetDocumentation, "LONG", "PTR;LONG;DWORD;PTR;PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetDocumentation), 13)
	DllStructSetData($tObject, "Callbacks", $GetDocumentation, 13)

	Local $GetDllEntry = DllCallbackRegister(ITypeInfo_GetDllEntry, "LONG", "PTR;LONG;INT;DWORD;PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetDllEntry), 14)
	DllStructSetData($tObject, "Callbacks", $GetDllEntry, 14)

	Local $GetRefTypeInfo = DllCallbackRegister(ITypeInfo_GetRefTypeInfo, "LONG", "PTR;DWORD;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetRefTypeInfo), 15)
	DllStructSetData($tObject, "Callbacks", $GetRefTypeInfo, 15)

	Local $Opnum15NotUsedOnWire = DllCallbackRegister(ITypeInfo_Opnum15NotUsedOnWire, "LONG", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Opnum15NotUsedOnWire), 16)
	DllStructSetData($tObject, "Callbacks", $Opnum15NotUsedOnWire, 16)

	Local $CreateInstance = DllCallbackRegister(ITypeInfo_CreateInstance, "LONG", "PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($CreateInstance), 17)
	DllStructSetData($tObject, "Callbacks", $CreateInstance, 17)

	Local $GetMops = DllCallbackRegister(ITypeInfo_GetMops, "LONG", "PTR;LONG;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetMops), 18)
	DllStructSetData($tObject, "Callbacks", $GetMops, 18)

	Local $GetContainingTypeLib = DllCallbackRegister(ITypeInfo_GetContainingTypeLib, "LONG", "PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetContainingTypeLib), 19)
	DllStructSetData($tObject, "Callbacks", $GetContainingTypeLib, 19)

	Local $Opnum19NotUsedOnWire = DllCallbackRegister(ITypeInfo_Opnum19NotUsedOnWire, "LONG", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Opnum19NotUsedOnWire), 20)
	DllStructSetData($tObject, "Callbacks", $Opnum19NotUsedOnWire, 20)

	Local $Opnum20NotUsedOnWire = DllCallbackRegister(ITypeInfo_Opnum20NotUsedOnWire, "LONG", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Opnum20NotUsedOnWire), 21)
	DllStructSetData($tObject, "Callbacks", $Opnum20NotUsedOnWire, 21)

	Local $Opnum21NotUsedOnWire = DllCallbackRegister(ITypeInfo_Opnum21NotUsedOnWire, "LONG", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Opnum21NotUsedOnWire), 22)
	DllStructSetData($tObject, "Callbacks", $Opnum21NotUsedOnWire, 22)

	Local $pData = MemCloneGlob($tObject)
	Local $tObject = DllStructCreate($tagObject, $pData)

	DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers

	Return DllStructGetPtr($tObject, "Object")
EndFunc

Func ITypeInfo_QueryInterface($pSelf, $pRIID, $pObj)
	ConsoleWrite("ITypeInfo_QueryInterface"&@CRLF)
EndFunc

Func ITypeInfo_AddRef($pSelf)
	ConsoleWrite("ITypeInfo_AddRef"&@CRLF)
	Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
	$tStruct.Ref += 1
	Return $tStruct.Ref
EndFunc

Func ITypeInfo_Release($pSelf)
	ConsoleWrite("ITypeInfo_Release"&@CRLF)
	Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
	$tStruct.Ref -= 1
	Return $tStruct.Ref
EndFunc

Func ITypeInfo_GetTypeAttr($pSelf, $ppTypeAttr, $pReserved)
	ConsoleWrite("ITypeInfo_GetTypeAttr"&@CRLF);TODO
	$tTYPEATTR = DllStructCreate($tagTYPEATTR)

	Local $aRet = DllCall("Ole32.dll", "LONG", "IIDFromString", "STR", $IID_IConnectionPointContainer, "PTR", 0)
	MsgBox(0, "", @error)
	Local $tGUID = DllStructCreate("DWORD;WORD;WORD;BYTE[8]", $aRet[2])

	DllStructSetData($tTYPEATTR, 1, DllStructGetData($tGUID, 1))
	DllStructSetData($tTYPEATTR, 2, DllStructGetData($tGUID, 2))
	DllStructSetData($tTYPEATTR, 3, DllStructGetData($tGUID, 3))
	DllStructSetData($tTYPEATTR, 4, DllStructGetData($tGUID, 4, 1), 1)
	DllStructSetData($tTYPEATTR, 4, DllStructGetData($tGUID, 4, 2), 2)
	DllStructSetData($tTYPEATTR, 4, DllStructGetData($tGUID, 4, 3), 3)
	DllStructSetData($tTYPEATTR, 4, DllStructGetData($tGUID, 4, 4), 4)
	DllStructSetData($tTYPEATTR, 4, DllStructGetData($tGUID, 4, 5), 5)
	DllStructSetData($tTYPEATTR, 4, DllStructGetData($tGUID, 4, 6), 6)
	DllStructSetData($tTYPEATTR, 4, DllStructGetData($tGUID, 4, 7), 7)
	DllStructSetData($tTYPEATTR, 4, DllStructGetData($tGUID, 4, 8), 8)

	;; MEMBERID_NIL = DISPID_UNKNOWN = -1

	DllStructSetData($tTYPEATTR, 7, -1)
	DllStructSetData($tTYPEATTR, 8, -1)
;~ 	DllStructSetData($tTYPEATTR, 10, 10)

	DllStructSetData(DllStructCreate("PTR", $ppTypeAttr), 1, MemCloneGlob($tTYPEATTR))
	DllStructSetData(DllStructCreate("DWORD", $pReserved), 1, 0)
	Return $S_OK
;~ 	Return $E_NOTIMPL
EndFunc

Func ITypeInfo_GetTypeComp($pSelf, $ppTComp)
	ConsoleWrite("ITypeInfo_GetTypeComp"&@CRLF)
EndFunc

Func ITypeInfo_GetFuncDesc($pSelf, $index, $ppFuncDesc, $pReserved)
	ConsoleWrite("ITypeInfo_GetFuncDesc"&@CRLF)
EndFunc

Func ITypeInfo_GetVarDesc($pSelf, $index, $ppVarDesc, $pReserved)
	ConsoleWrite("ITypeInfo_GetVarDesc"&@CRLF)
EndFunc

Func ITypeInfo_GetNames($pSelf, $memid, $rgBstrNames, $cMaxNames, $pcNames)
	ConsoleWrite("ITypeInfo_GetNames"&@CRLF)
EndFunc

Func ITypeInfo_GetRefTypeOfImplType($pSelf, $index, $pRefType)
	ConsoleWrite("ITypeInfo_GetRefTypeOfImplType"&@CRLF)
EndFunc

Func ITypeInfo_GetImplTypeFlags($pSelf, $index, $pImplTypeFlags)
	ConsoleWrite("ITypeInfo_GetImplTypeFlags"&@CRLF)
EndFunc

Func ITypeInfo_Opnum10NotUsedOnWire($pSelf)
	ConsoleWrite("ITypeInfo_Opnum10NotUsedOnWire"&@CRLF)
EndFunc

Func ITypeInfo_Opnum11NotUsedOnWire($pSelf)
	ConsoleWrite("ITypeInfo_Opnum11NotUsedOnWire"&@CRLF)
EndFunc

Func ITypeInfo_GetDocumentation($pSelf, $memid, $refPtrFlags, $pBstrName, $pBstrDocString, $pdwHelpContext, $pBstrHelpFile)
	ConsoleWrite("ITypeInfo_GetDocumentation"&@CRLF)
EndFunc

Func ITypeInfo_GetDllEntry($pSelf, $memid, $invKind, $refPtrFlags, $pBstrDllName, $pBstrName, $pwOrdinal)
	ConsoleWrite("ITypeInfo_GetDllEntry"&@CRLF)
EndFunc

Func ITypeInfo_GetRefTypeInfo($pSelf, $hRefType, $ppTInfo)
	ConsoleWrite("ITypeInfo_GetRefTypeInfo"&@CRLF)
EndFunc

Func ITypeInfo_Opnum15NotUsedOnWire($pSelf)
	ConsoleWrite("ITypeInfo_Opnum15NotUsedOnWire"&@CRLF)
EndFunc

Func ITypeInfo_CreateInstance($pSelf, $riid, $ppvObj)
	ConsoleWrite("ITypeInfo_CreateInstance"&@CRLF)
EndFunc

Func ITypeInfo_GetMops($pSelf, $memid, $pBstrMops)
	ConsoleWrite("ITypeInfo_GetMops"&@CRLF)
EndFunc

Func ITypeInfo_GetContainingTypeLib($pSelf, $ppTLib, $pIndex)
	DllStructSetData(DllStructCreate("UINT", $ppTLib), 1, ITypeLib())
	DllStructSetData(DllStructCreate("UINT", $pIndex), 1, 0)
	ConsoleWrite("ITypeInfo_GetContainingTypeLib"&@CRLF)
	Return $S_OK
EndFunc

Func ITypeInfo_Opnum19NotUsedOnWire($pSelf)
	ConsoleWrite("ITypeInfo_Opnum19NotUsedOnWire"&@CRLF)
	Return $E_NOTIMPL
EndFunc

Func ITypeInfo_Opnum20NotUsedOnWire($pSelf)
	ConsoleWrite("ITypeInfo_Opnum20NotUsedOnWire"&@CRLF)
EndFunc

Func ITypeInfo_Opnum21NotUsedOnWire($pSelf)
	ConsoleWrite("ITypeInfo_Opnum21NotUsedOnWire"&@CRLF)
EndFunc