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

$IID_IRunningObjectTable = "{00000010-0000-0000-C000-000000000046}"
$IID_IMoniker = "{0000000f-0000-0000-C000-000000000046}"

$IID_IConnectionPoint = "{B196B286-BAB4-101A-B69C-00AA00341D07}"
$IID__DMSMQEventEvents = "{D7D6E078-DCCD-11d0-AA4B-0060970DEBAE}"

$AutoItError = ObjEvent("AutoIt.Error", "ErrFunc") ; Install a custom error handler
Func ErrFunc($oError)
	ConsoleWrite("!>COM Error !"&@CRLF&"!>"&@TAB&"Number: "&Hex($oError.Number,8)&@CRLF&"!>"&@TAB&"Windescription: "&StringRegExpReplace($oError.windescription,"\R$","")&@CRLF&"!>"&@TAB&"Source: "&$oError.source&@CRLF&"!>"&@TAB&"Description: "&$oError.description&@CRLF&"!>"&@TAB&"Helpfile: "&$oError.helpfile&@CRLF&"!>"&@TAB&"Helpcontext: "&$oError.helpcontext&@CRLF&"!>"&@TAB&"Lastdllerror: "&$oError.lastdllerror&@CRLF&"!>"&@TAB&"Scriptline: "&$oError.scriptline&@CRLF)
EndFunc ;==>ErrFunc

$IDispatch = IDispatch(QueryInterface2, AddRef2, Release2, GetTypeInfoCount2, GetTypeInfo2, GetIDsOfNames2, Invoke2)

Func Getter($o)
	Return SetError(0x000010D2, 0, 0); ERROR_EMPTY
;~ 	Return 1
EndFunc

$IDispatch.name = "my name is Danny"
$IRunningObjectTable = DllCall("Ole32.dll","LONG","GetRunningObjectTable","DWORD",0,"PTR*",0)
If @error<>0 Then Exit MsgBox(0, @ScriptLineNumber, @error)
$IRunningObjectTable = ObjCreateInterface($IRunningObjectTable[2],$IID_IRunningObjectTable,"Register HRESULT(DWORD;PTR;PTR;DWORD*);Revoke HRESULT(DWORD);IsRunning HRESULT(PTR*);GetObject HRESULT(PTR;PTR*);NoteChangeTime HRESULT(DWORD;PTR*);GetTimeOfLastChange HRESULT(PTR*;PTR*);EnumRunning HRESULT(PTR*);",True)
If @error<>0 Then Exit MsgBox(0, @ScriptLineNumber, @error)

$sCLSID="AutoIt.COMDemo"
$IMoniker=DllCall("Ole32.dll", "LONG", "CreateFileMoniker", "WSTR", $sCLSID, "PTR*", 0)
If @error<>0 Then Exit MsgBox(0, @ScriptLineNumber, @error)
;~ $IMoniker = ObjCreateInterface($IMoniker[2], $IID_IMoniker, "GetClassID;IsDirty;Load;Save;GetSizeMax;BindToObject;BindToStorage;Reduce;ComposeWith;Enum;IsEqual;Hash;IsRunning;GetTimeOfLastChange;Inverse;CommonPrefixWith;RelativePathTo;GetDisplayName;ParseDisplayName;IsSystemMoniker")

Global Const $ROTFLAGS_REGISTRATIONKEEPSALIVE=0x01
Global Const $ROTFLAGS_ALLOWANYCLIENT=0x02

AddRef(Ptr($IDispatch))

$dwRegister=0
ConsoleWrite("Before $IRunningObjectTable.Register"&@CRLF)
;~ $r=$IRunningObjectTable.Register( $ROTFLAGS_REGISTRATIONKEEPSALIVE, ptr($IDispatch), $IMoniker[2], $dwRegister ) ; use this to allow multiple use for now
$r=$IRunningObjectTable.Register( 0, ptr($IDispatch), $IMoniker[2], $dwRegister )
;~ MsgBox(0, "$IRunningObjectTable.Register", $r)
If @error<>0 Then Exit 1
If $dwRegister=0 Then
	MsgBox(0, "", _WinAPI_GetErrorMessage($r))
	Exit 2
EndIf
ConsoleWrite("After $IRunningObjectTable.Register"&@CRLF)

$IMoniker = ObjCreateInterface($IMoniker[2], $IID_IMoniker)

;~ ConsoleWrite("BeforePadding"&@CRLF);adding extra count, to prevent miscount if object is used before release from ROT
;~ AddRef(Ptr($IDispatch))
;~ AddRef(Ptr($IDispatch))
;~ AddRef(Ptr($IDispatch));TODO: look into RunningObjectTable calling IUnknown:Release 3 times if used one or more times, but only once if not used before release @.@
;~ ConsoleWrite("AfterPadding"&@CRLF)

$IMoniker=0

$IRunningObjectTable=0

Opt("GuiOnEventMode", 1)

$hWnd=GUICreate("Title",700,320)

GUICtrlCreateButton("Ref count", 10, 10, 100, 35)
GUICtrlSetOnEvent(-1, "RefCount")
GUICtrlCreateButton("Ref += 1", 10, 10+35+10, 100, 35)
GUICtrlSetOnEvent(-1, "AddRef3")

GUISetState(@SW_SHOW,$hWnd)

GUISetOnEvent(-3, "_MyExit", $hWnd)

While 1
	Sleep(10)
WEnd

Func _MyExit()
	Local $IRunningObjectTable = DllCall("Ole32.dll","LONG","GetRunningObjectTable","DWORD",0,"PTR*",0)
	If @error<>0 Then Exit MsgBox(0, @ScriptLineNumber, @error)
	$IRunningObjectTable = ObjCreateInterface($IRunningObjectTable[2],$IID_IRunningObjectTable,"Register HRESULT(DWORD;PTR;PTR;DWORD*);Revoke HRESULT(DWORD);IsRunning HRESULT(PTR*);GetObject HRESULT(PTR*;PTR**);NoteChangeTime HRESULT(DWORD;PTR*);GetTimeOfLastChange HRESULT(PTR*;PTR*);EnumRunning HRESULT(PTR*);",True)
	If @error<>0 Then Exit MsgBox(0, @ScriptLineNumber, @error)

	$IRunningObjectTable.Revoke($dwRegister)
	ConsoleWrite("IRunningObjectTable.Revoke"&@CRLF)
	$IRunningObjectTable=0
	$IDispatch=0
	Exit
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
	ConsoleWrite("QueryInterface: "&$_&@CRLF)
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
	;$aRet = DllCall("oleaut32.dll","long","CreateDispTypeInfo","struct*",0,"DWORD",0,"ptr*",0)
	;DllStructSetData(DllStructCreate("ptr", $ppTInfo), 1, $aRet[3])
;~ 	DllStructSetData(DllStructCreate("ptr", $ppTInfo), 1, 0)
	;Return $S_OK
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
	ConsoleWrite("dispIdMember: "&$dispIdMember&@CRLF)
	ConsoleWrite("pExcepInfo: "&$pExcepInfo&@CRLF);EXCEPTINFO structure
	ConsoleWrite("Invoke: "&$_&@CRLF)
	Return $_
EndFunc

Func RefCount()
	ConsoleWrite( DllStructGetData(DllStructCreate("int Ref", Ptr($IDispatch)-8), 1) & @CRLF )
EndFunc