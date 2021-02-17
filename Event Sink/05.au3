#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         genius257

#ce ----------------------------------------------------------------------------

;~ MsgBox(0, "", Hex(-1073741819, 8))

#AutoIt3Wrapper_Run_Au3Check=N

#include <WinAPIDiag.au3>
#include "..\AutoItObject_Internal.au3"

if Not IsDeclared("E_HANDLE") Then Global Const $E_HANDLE = 0x80070006
if Not IsDeclared("E_POINTER") Then Global Const $E_POINTER = 0x80004003
if Not IsDeclared("E_NOINTERFACE") Then Global Const $E_NOINTERFACE = 0x80004002
if Not IsDeclared("S_OK") Then Global Const $S_OK = 0x00000000
if Not IsDeclared("E_INVALIDARG") Then Global Const $E_INVALIDARG = 0x80070057
if Not IsDeclared("E_NOTIMPL") Then Global Const $E_NOTIMPL = 0x80004001

;~ MsgBox(0, "", _WinAPI_GetErrorMessage(0x80028017, 1033))

#include "ITypeInfo.au3"

$AutoItError = ObjEvent("AutoIt.Error", "ErrFunc") ; Install a custom error handler
Func ErrFunc($oError)
	ConsoleWrite("!>COM Error !"&@CRLF&"!>"&@TAB&"Number: "&Hex($oError.Number,8)&@CRLF&"!>"&@TAB&"Windescription: "&StringRegExpReplace($oError.windescription,"\R$","")&@CRLF&"!>"&@TAB&"Source: "&$oError.source&@CRLF&"!>"&@TAB&"Description: "&$oError.description&@CRLF&"!>"&@TAB&"Helpfile: "&$oError.helpfile&@CRLF&"!>"&@TAB&"Helpcontext: "&$oError.helpcontext&@CRLF&"!>"&@TAB&"Lastdllerror: "&$oError.lastdllerror&@CRLF&"!>"&@TAB&"Scriptline: "&$oError.scriptline&@CRLF)
EndFunc ;==>ErrFunc

$IDispatch = IDispatch(QueryInterface2, AddRef2, Release2, GetTypeInfoCount2, GetTypeInfo2, GetIDsOfNames2, Invoke2)
;~ $IDispatch = IDispatch()
;~ ObjName($IDispatch, 1)
;~ ObjName($IDispatch, 2)
;~ ObjName($IDispatch, 3)
;~ ObjName($IDispatch, 4)
;~ ObjName($IDispatch, 5)
;~ ObjName($IDispatch, 6)
;~ ObjName($IDispatch, 7)
;~ MsgBox(0, "", Execute("$IDispatch(0)"))
;~ DllCall("ole32.dll", "none", "OleUninitialize")
;~ DllCall("ole32.dll", "none", "CoUninitialize")
;~ DllCall("ole32.dll", "none", "CoUninitialize")
;~ DllCall("ole32.dll", "none", "CoUninitialize")
ObjEvent($IDispatch, "_event0_", "Event0")
ConsoleWrite("0x"&Hex(@error)&@CRLF)
;~ ObjEvent($IDispatch, "_event1_", "Event1")
;~ ObjEvent($IDispatch, "_event2_", "Event2")
;~ $IDispatch.a = 1
;~ MsgBox(0, "", $IDispatch.a)
;~ $IDispatch.__defineGetter('getter', Getter)
;~ $getter = $IDispatch.getter
;~ ConsoleWrite(Hex(@error)&@CRLF)
;~ MsgBox(0, Hex(@error), $getter)
;~ Sleep(1000)
;~ Sleep(1000)
;~ Sleep(1000)
;~ Sleep(1000)
Exit

Func Getter($o)
	Return SetError(0x000010D2, 0, 0); ERROR_EMPTY
;~ 	Return 1
EndFunc

#Region IConnectionPointContainer interface
	Func IConnectionPointContainer_QueryInterface($pSelf, $pRIID, $pObj)
		ConsoleWrite("IConnectionPointContainer_QueryInterface"&@CRLF)
		Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
		If Not ($sGUID = $IID_IConnectionPointContainer) Then Return $E_NOINTERFACE
		DllStructSetData(DllStructCreate("ptr", $pObj), 1, $pSelf)
		Return $S_OK
	EndFunc

	Func IConnectionPointContainer_AddRef($pSelf)
		ConsoleWrite("IConnectionPointContainer_AddRef"&@CRLF)
		Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
		$tStruct.Ref += 1
		Return $tStruct.Ref
	EndFunc

	Func IConnectionPointContainer_Release($pSelf)
		ConsoleWrite("IConnectionPointContainer_AddRef"&@CRLF)
		Local $tStruct = DllStructCreate("int Ref", $pSelf-8)
		$tStruct.Ref -= 1
		Return $tStruct.Ref
	EndFunc

	Func EnumConnectionPoints($pSelf, $ppEnum)
		ConsoleWrite("IConnectionPointContainer_EnumConnectionPoints"&@CRLF)
		If $ppEnum = 0 Then Return $E_POINTER
		Return $E_NOTIMPL
	EndFunc

	Func FindConnectionPoint($pSelf, $pRIID, $ppvObject)
		ConsoleWrite("IConnectionPointContainer_FindConnectionPoint - "&$pRIID&@CRLF)
		If Not ($ppvObject = 0x00000000) Then Return $E_POINTER
		;$pRIID seems to be 0x00000000 when used by AutoIt's ObjEvent

;~ 		Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
;~ 		ConsoleWrite(DllStructGetData(DllStructCreate("PTR", $ppvObject), 1)&@CRLF)

		Return 0x80040200;$CONNECT_E_NOCONNECTION

;~ 		DllStructSetData(DllStructCreate("PTR", $ppvObject), 1, IConnectionPoint())

;~ 		Return $S_OK
	EndFunc
#EndRegion IConnectionPointContainer interface

#Region IConnectionPoint Interface
	Func IConnectionPoint()
		Return 0
		Local $tagObject = "int RefCount;int Size;ptr Object;ptr Methods[5];int_ptr Callbacks[5];"
		Local $tObject = DllStructCreate($tagObject)

		Local $QueryInterface = DllCallbackRegister(QueryInterface, "LONG", "ptr;ptr;ptr")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($QueryInterface), 1)
		DllStructSetData($tObject, "Callbacks", $QueryInterface, 1)

		Local $AddRef = DllCallbackRegister(AddRef, "dword", "PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($AddRef), 2)
		DllStructSetData($tObject, "Callbacks", $AddRef, 2)

		Local $Release = DllCallbackRegister(Release, "dword", "PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Release), 3)
		DllStructSetData($tObject, "Callbacks", $Release, 3)

		Local $EnumConnectionPoints = DllCallbackRegister(GetConnectionInterface, "dword", "PTR;PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($EnumConnectionPoints), 3)
		DllStructSetData($tObject, "Callbacks", $EnumConnectionPoints, 3)

		Local $FindConnectionPoint = DllCallbackRegister(GetConnectionPointContainer, "dword", "PTR;PTR;PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($FindConnectionPoint), 3)
		DllStructSetData($tObject, "Callbacks", $FindConnectionPoint, 3)

		Local $FindConnectionPoint = DllCallbackRegister(Advise, "dword", "PTR;PTR;PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($FindConnectionPoint), 3)
		DllStructSetData($tObject, "Callbacks", $FindConnectionPoint, 3)

		Local $FindConnectionPoint = DllCallbackRegister(Unadvise, "dword", "PTR;PTR;PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($FindConnectionPoint), 3)
		DllStructSetData($tObject, "Callbacks", $FindConnectionPoint, 3)

		Local $FindConnectionPoint = DllCallbackRegister(EnumConnections, "dword", "PTR;PTR;PTR")
		DllStructSetData($tObject, "Methods", DllCallbackGetPtr($FindConnectionPoint), 3)
		DllStructSetData($tObject, "Callbacks", $FindConnectionPoint, 3)

		Local $pData = MemCloneGlob($tObject)

		Local $tObject = DllStructCreate($tagObject, $pData)

		DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
		$IConnectionPointContainer =DllStructGetPtr($tObject, "Object") ;ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IConnectionPointContainer, Default, True) ; pointer that's wrapped into object
		If @error<>0 Then
			$IConnectionPointContainer=0
			Return $E_NOINTERFACE
		EndIf
	EndFunc

	Func GetConnectionInterface($pSelf, $pIID)

	EndFunc

	Func GetConnectionPointContainer($pSelf, $ppCPC)

	EndFunc

	Func Advise($pSelf, $pUnkSink, $pdwCookie)

	EndFunc

	Func Unadvise($pSelf, $dwCookie)

	EndFunc

	Func EnumConnections($pSelf, $ppEnum)

	EndFunc
#EndRegion IConnectionPoint Interface

Func QueryInterface2($pSelf, $pRIID, $pObj)
	Static $IConnectionPointContainer=0
	Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
	ConsoleWrite("["&@ScriptLineNumber&"]: QueryInterface - "&$sGUID&@CRLF)
	Local $_ = QueryInterface($pSelf, $pRIID, $pObj)
;~ 	If $sGUID="{4C1E39E1-E3E3-4296-AA86-EC938D896E92}" Then Return $S_OK ;more problems
	If $sGUID = $IID_IConnectionPointContainer Then
		If $IConnectionPointContainer=0 Then
			Local $tagObject = "int RefCount;int Size;ptr Object;ptr Methods[5];int_ptr Callbacks[5];"
			Local $tObject = DllStructCreate($tagObject)

			Local $QueryInterface = DllCallbackRegister(IConnectionPointContainer_QueryInterface, "LONG", "ptr;ptr;ptr")
			DllStructSetData($tObject, "Methods", DllCallbackGetPtr($QueryInterface), 1)
			DllStructSetData($tObject, "Callbacks", $QueryInterface, 1)

			Local $AddRef = DllCallbackRegister(IConnectionPointContainer_AddRef, "dword", "PTR")
			DllStructSetData($tObject, "Methods", DllCallbackGetPtr($AddRef), 2)
			DllStructSetData($tObject, "Callbacks", $AddRef, 2)

			Local $Release = DllCallbackRegister(IConnectionPointContainer_Release, "dword", "PTR")
			DllStructSetData($tObject, "Methods", DllCallbackGetPtr($Release), 3)
			DllStructSetData($tObject, "Callbacks", $Release, 3)

			Local $EnumConnectionPoints = DllCallbackRegister(EnumConnectionPoints, "dword", "PTR;PTR")
			DllStructSetData($tObject, "Methods", DllCallbackGetPtr($EnumConnectionPoints), 4)
			DllStructSetData($tObject, "Callbacks", $EnumConnectionPoints, 4)

			Local $FindConnectionPoint = DllCallbackRegister(FindConnectionPoint, "dword", "PTR;PTR;PTR")
			DllStructSetData($tObject, "Methods", DllCallbackGetPtr($FindConnectionPoint), 5)
			DllStructSetData($tObject, "Callbacks", $FindConnectionPoint, 5)

			Local $pData = MemCloneGlob($tObject)

			Local $tObject = DllStructCreate($tagObject, $pData)

			DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
			$IConnectionPointContainer =DllStructGetPtr($tObject, "Object") ;ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $IID_IConnectionPointContainer, Default, True) ; pointer that's wrapped into object
			If @error<>0 Then
				$IConnectionPointContainer=0
				Return $E_NOINTERFACE
			EndIf
		EndIf

		;TODO
		Local $tStruct = DllStructCreate("ptr", $pObj)
		DllStructSetData($tStruct, 1, $IConnectionPointContainer)
		IConnectionPointContainer_AddRef($IConnectionPointContainer)
		Return $S_OK
	EndIf
	Return $_
EndFunc

Func AddRef3()
	AddRef2(ptr($IDispatch))
EndFunc

Func AddRef2($pSelf)
	Local $_=AddRef($pSelf)
	ConsoleWrite("AddRef: "&$_&@CRLF)
	Return $_
EndFunc

Func Release2($pSelf)
	Local $_=Release($pSelf)
	ConsoleWrite("Release: "&$_&@CRLF)
	Return $_
EndFunc

Func GetIDsOfNames2($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
;~ 	AddRef2($pSelf)
	Local $_=GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
	ConsoleWrite("GetIDsOfNames: "&$_&@CRLF)
	Return $_
EndFunc

Func GetTypeInfo2($pSelf, $iTInfo, $lcid, $ppTInfo)
;~ 	$aRet = DllCall("oleaut32.dll","long","CreateDispTypeInfo","struct*",0,"DWORD",$lcid,"ptr*",0)
;~ 	$aRet = DllCall("oleaut32.dll","long","CreateDispTypeInfo","PTR",0,"DWORD",$lcid,"ptr",$ppTInfo)
;~ 	DllStructSetData(DllStructCreate("ptr", $ppTInfo), 1, $aRet[3])
	DllStructSetData(DllStructCreate("ptr", $ppTInfo), 1, ITypeInfo())
	Return $S_OK
	Local $_=GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
	ConsoleWrite("GetTypeInfo: 0x"&Hex($_, 8)&@CRLF)
	ConsoleWrite(@TAB&"iTInfo: "&$iTInfo&@CRLF)
	ConsoleWrite(@TAB&"lcid: "&$lcid&@CRLF)
	ConsoleWrite(@TAB&"ppTInfo: "&DllStructGetData(DllStructCreate("ptr", $ppTInfo), 1)&@CRLF)

;~ 	Local Const $tagInterfacedata = "ptr MethodData;UINT cMembers"
;~ 	Local Const $Struidata = DllStructCreate($tagInterfacedata)
;~ 	Local Const $aRet = DllCall("OleAut32.dll","long","CreateDispTypeInfo","struct*",$Struidata,"DWORD",0,"ptr*",0)
;~ 	Local $tagInterfacedata = "ptr MethodData;UINT cMembers"
;~ 	Local $Struidata = DllStructCreate($tagInterfacedata)
;~ 	DllStructSetData($Struidata,"cMembers",0)
;~ 	Local $aRet = DllCall("OleAut32.dll","long","CreateDispTypeInfo","struct*",$Struidata,"DWORD",0,"ptr*",0)
;~ 	MsgBox(0, "", $aRet[0])
;~ 	$IUnknown = ObjCreateInterface($aRet[3], $IID_IUnknown, "QueryInterface hresult(clsid;ptr*); AddRef dword(); Release dword();", False)
;~ 	$IUnknown.AddRef()
	$aRet = DllCall("oleaut32.dll","long","CreateDispTypeInfo","struct*",0,"DWORD",0,"ptr*",0)
	DllStructSetData(DllStructCreate("ptr", $ppTInfo), 1, $aRet[3])
;~ 	DllStructSetData(DllStructCreate("ptr", $ppTInfo), 1, 0)
	Return $S_OK
;~ 	Return $E_NOTIMPL

	Return $_
EndFunc

Func GetTypeInfoCount2($pSelf, $pctinfo)
	Local $_=GetTypeInfoCount($pSelf, $pctinfo)
;~ 	DllStructSetData(DllStructCreate("UINT",$pctinfo),1, 1)
	ConsoleWrite("GetTypeInfoCount: "&$_&@CRLF)
	Return $_
EndFunc

Func Invoke2($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
;~ 	AddRef2($pSelf)
;~ 	ObjCreateInterface($pSelf, $IID_IDispatch, Default, True) ; pointer that's wrapped into object
;~ 	ConsoleWrite("Invoke "&$dispIdMember&@CRLF)
;~ 	Return 0
	Local $_=Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
	ConsoleWrite("pExcepInfo: "&$pExcepInfo&@CRLF);EXCEPTINFO structure
	ConsoleWrite("Invoke: "&$_&@CRLF)
	Return $_
EndFunc
