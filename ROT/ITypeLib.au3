#include-once
#include <WinAPIMisc.au3>

Func ITypeLib(); https://msdn.microsoft.com/en-us/library/cc237787.aspx
	Local $tagObject = "int RefCount;int Size;ptr Object;ptr Methods[22];int_ptr Callbacks[22];"
	Local $tObject = DllStructCreate($tagObject)

	Local $QueryInterface = DllCallbackRegister(ITypeLib_QueryInterface, "LONG", "ptr;ptr;ptr")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($QueryInterface), 1)
	DllStructSetData($tObject, "Callbacks", $QueryInterface, 1)

	Local $AddRef = DllCallbackRegister(ITypeLib_AddRef, "dword", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($AddRef), 2)
	DllStructSetData($tObject, "Callbacks", $AddRef, 2)

	Local $Release = DllCallbackRegister(ITypeLib_Release, "dword", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Release), 3)
	DllStructSetData($tObject, "Callbacks", $Release, 3)

	Local $GetTypeInfoCount = DllCallbackRegister(ITypeLib_GetTypeInfoCount, "LONG", "PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfoCount), 4)
	DllStructSetData($tObject, "Callbacks", $GetTypeInfoCount, 4)

	Local $GetTypeInfo = DllCallbackRegister(ITypeLib_GetTypeInfo, "LONG", "PTR;UINT;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfo), 5)
	DllStructSetData($tObject, "Callbacks", $GetTypeInfo, 5)

	Local $GetTypeInfoType = DllCallbackRegister(ITypeLib_GetTypeInfoType, "LONG", "PTR;UINT;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfoType), 6)
	DllStructSetData($tObject, "Callbacks", $GetTypeInfoType, 6)

	Local $GetTypeInfoOfGuid = DllCallbackRegister(ITypeLib_GetTypeInfoOfGuid, "LONG", "PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeInfoOfGuid), 7)
	DllStructSetData($tObject, "Callbacks", $GetTypeInfoOfGuid, 7)

	Local $GetLibAttr = DllCallbackRegister(ITypeLib_GetLibAttr, "LONG", "PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetLibAttr), 8)
	DllStructSetData($tObject, "Callbacks", $GetLibAttr, 8)

	Local $GetTypeComp = DllCallbackRegister(ITypeLib_GetTypeComp, "LONG", "PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetTypeComp), 9)
	DllStructSetData($tObject, "Callbacks", $GetTypeComp, 9)

	Local $GetDocumentation = DllCallbackRegister(ITypeLib_GetDocumentation, "LONG", "PTR;INT;DWORD;PTR;PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($GetDocumentation), 10)
	DllStructSetData($tObject, "Callbacks", $GetDocumentation, 10)

	Local $IsName = DllCallbackRegister(ITypeLib_IsName, "LONG", "PTR;PTR;ULONG;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($IsName), 11)
	DllStructSetData($tObject, "Callbacks", $IsName, 11)

	Local $FindName = DllCallbackRegister(ITypeLib_FindName, "LONG", "PTR;PTR;ULONG;PTR;PTR;PTR;PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($FindName), 12)
	DllStructSetData($tObject, "Callbacks", $FindName, 12)

	Local $Opnum12NotUsedOnWire = DllCallbackRegister(ITypeLib_Opnum12NotUsedOnWire, "LONG", "PTR")
	DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Opnum12NotUsedOnWire), 13)
	DllStructSetData($tObject, "Callbacks", $Opnum12NotUsedOnWire, 13)

	Local $pData = MemCloneGlob($tObject)
	Local $tObject = DllStructCreate($tagObject, $pData)

	DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers

	Return DllStructGetPtr($tObject, "Object")
EndFunc

Func ITypeLib_QueryInterface($pSelf, $pRIID, $pObj)
	ConsoleWrite("ITypeInfo_QueryInterface"&@CRLF)
EndFunc

Func ITypeLib_AddRef($pSelf)
	ConsoleWrite("ITypeInfo_AddRef"&@CRLF)
	Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
	$tStruct.Ref += 1
	Return $tStruct.Ref
EndFunc

Func ITypeLib_Release($pSelf)
	ConsoleWrite("ITypeInfo_Release"&@CRLF)
	Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
	$tStruct.Ref -= 1
	Return $tStruct.Ref
EndFunc

Func ITypeLib_GetTypeInfoCount($pSelf, $pcTInfo)
	ConsoleWrite("ITypeLib_GetTypeInfoCount"&@CRLF)
EndFunc

Func ITypeLib_GetTypeInfo($pSelf, $index, $ppTInfo)
	ConsoleWrite("ITypeLib_GetTypeInfo"&@CRLF)
EndFunc

Func ITypeLib_GetTypeInfoType($pSelf, $index, $pTKind)
	ConsoleWrite("ITypeLib_GetTypeInfoType"&@CRLF)
EndFunc

Func ITypeLib_GetTypeInfoOfGuid($pSelf, $guid, $ppTInfo)
	ConsoleWrite("ITypeLib_GetTypeInfoOfGuid"&@CRLF)
EndFunc

Func ITypeLib_GetLibAttr($pSelf, $ppTLibAttr, $pReserved)
	ConsoleWrite("ITypeLib_GetLibAttr"&@CRLF)
EndFunc

Func ITypeLib_GetTypeComp($pSelf, $ppTComp)
	ConsoleWrite("ITypeLib_GetTypeComp"&@CRLF)
EndFunc

Func ITypeLib_GetDocumentation($pSelf, $index, $refPtrFlags, $pBstrName, $pBstrDocString, $pdwHelpContext, $pBstrHelpFile)
	ConsoleWrite("ITypeLib_GetDocumentation"&@CRLF)
EndFunc

Func ITypeLib_IsName($pSelf, $szNameBuf, $lHashVal, $pfName, $pBstrNameInLibrary)
	ConsoleWrite("ITypeLib_IsName"&@CRLF)
EndFunc

Func ITypeLib_FindName($pSelf, $szNameBuf, $lHashVal, $ppTInfo, $rgMemId, $pcFound, $pBstrNameInLibrary)
	ConsoleWrite("ITypeLib_FindName"&@CRLF)
	ConsoleWrite("+> szNameBuf: "&_winapi_getstring($szNameBuf, True)&@CRLF)
	ConsoleWrite("+> lHashVal: "&$lHashVal&@CRLF)
	ConsoleWrite("+> pcFound: "&DllStructGetData(DllStructCreate("USHORT", $pcFound), 1)&@CRLF)
;~ 	Return $E_NOTIMPL
	DllStructSetData(DllStructCreate("PTR", $ppTInfo), 1, ITypeInfo())
	DllStructSetData(DllStructCreate("INT", $rgMemId), 1, 123)
	DllStructSetData(DllStructCreate("USHORT", $pcFound), 1, 0)
	 ;_WinAPI_CreateString()
	DllStructSetData(DllStructCreate("PTR", $pBstrNameInLibrary), 1, $szNameBuf)
	Return $S_OK
EndFunc

Func ITypeLib_Opnum12NotUsedOnWire($pSelf)
	ConsoleWrite("ITypeLib_Opnum12NotUsedOnWire"&@CRLF)
EndFunc