#include "../AutoItObject_Internal.au3"
#include "../AutoItObject_Internal_ROT.au3"

#cs
    possible regestry locations for getting information about undefined IID's
    Computer\HKEY_CLASSES_ROOT\TypeLib\{00000600-0000-0010-8000-00AA006D2EA4}
    Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{00000000-0000-0000-C000-000000000046}
#ce

;MsgBox(0, "", RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{F2153260-232E-4474-9D0A-9F2AB153441D}", ""))
;Exit

$sIIDs = FileRead("iids.txt")

$oIDispatch = IDispatch(QueryInterface2, AddRef2, Release2, GetTypeInfoCount2, GetTypeInfo2, GetIDsOfNames2, Invoke2)

$dwRegister = _AOI_ROT_register($oIDispatch, "AOI.Debug", True)

OnAutoItExitRegister("CleanUp")

ConsoleWrite(@CRLF&@CRLF) ; For seeing what communication is not part of the initial object creation process.

While 1
    Sleep(10)
WEnd

Func QueryInterface2($pSelf, $pRIID, $pObj)
    Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
    $sName = RegRead(StringFormat("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\%s", $sGUID), "")
    if ($sName = "") Then
        Local $aIID = StringRegExp($sIIDs, StringFormat("(?m)^%s\s+(.*)$", StringRegExpReplace($sGUID, "[{}]", "")), 2)
        $sName = Execute("$aIID[1]")
    EndIf
    $sName = $sName == "" ? "?" : $sName
    ConsoleWrite($sGUID&" "&$sName&@CRLF)
    Return QueryInterface($pSelf, $pRIID, $pObj)
EndFunc

Func AddRef2($pSelf)
    Return AddRef($pSelf)
EndFunc

Func Release2($pSelf)
    Return Release($pSelf)
EndFunc

Func GetTypeInfoCount2($pSelf, $pctinfo)
    Return GetTypeInfoCount($pSelf, $pctinfo)
EndFunc

Func GetTypeInfo2($pSelf, $iTInfo, $lcid, $ppTInfo)
    Return GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
EndFunc

Func GetIDsOfNames2($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
    Return GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
EndFunc

Func Invoke2($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
    Return Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
EndFunc

Func CleanUp()
    _AOI_ROT_revoke($dwRegister)
EndFunc
