#include-once

If Not IsDeclared("IID_IRunningObjectTable") Then Global Const $IID_IRunningObjectTable = "{00000010-0000-0000-C000-000000000046}"
If Not IsDeclared("IID_IMoniker") Then Global Const $IID_IMoniker = "{0000000F-0000-0000-C000-000000000046}"

If Not IsDeclared("ROTFLAGS_REGISTRATIONKEEPSALIVE") Then Global Const $ROTFLAGS_REGISTRATIONKEEPSALIVE = 0x01
If Not IsDeclared("ROTFLAGS_ALLOWANYCLIENT") Then Global Const $ROTFLAGS_ALLOWANYCLIENT = 0x02

#cs
# Registers an object and its identifying moniker in the running object table (ROT).
# @param ptr|object $vObject Object that is being registered as running.
# @param string     $sCLSID  The string identifier to use in the ROT
# @param bool|int   $bStrong Indicates a strong registration for the object. If not boolean, the value will be parsed directly to IRunningObjectTable::Register as the grfFlags parameter
# @returns int The identifier of the ROT entry.
#ce
Func _AOI_ROT_register($vObject, $sCLSID, $bStrong = False)
    Local $grfFlags = IsBool($bStrong) ? ($bStrong ? $ROTFLAGS_REGISTRATIONKEEPSALIVE : 0) : $bStrong
    Local $dwRegister = 0
    Local $IRunningObjectTable = __AOI_ROT_GetRunningObjectTable()
    If @error <> 0 Then Return SetError(@error, @extended, $dwRegister)
    Local $IMoniker = __AOI_ROT_CreateFileMoniker($sCLSID)
    If @error <> 0 Then Return SetError(@error, @extended, $dwRegister)
    ;TODO: maybe temp. override ObjEvent for autoit COM Error, to always get @error instead of crash from the $IRunningObjectTable method call
    Local $iRet = $IRunningObjectTable.Register($grfFlags, ptr($vObject), Ptr($IMoniker), $dwRegister)
    If @error <> 0 Then Return SetError(@error, @extended, $dwRegister)
    If $iRet <> 0 Then Return SetError($iRet, @extended, $dwRegister)
    Return $dwRegister
EndFunc

#cs
# Removes an entry from the running object table (ROT) that was previously registered by a call to IRunningObjectTable::Register.
# @param int $dwRegister The identifier of the ROT entry to be revoked.
# @returns bool
#ce
Func _AOI_ROT_revoke($dwRegister)
    Local $IRunningObjectTable = __AOI_ROT_GetRunningObjectTable()
    If @error <> 0 Then Return SetError(@error, @extended, $dwRegister)
    Local $iRet = $IRunningObjectTable.Revoke($dwRegister)
    If @error <> 0 Then Return SetError(@error, @extended, False)
    If $iRet <> 0 Then Return SetError($iRet, @extended, False)
    Return True
EndFunc

#cs
# @internal
#ce
Func __AOI_ROT_GetRunningObjectTable()
    Local $aRet = DllCall("Ole32.dll", "LONG", "GetRunningObjectTable", "DWORD", 0, "PTR*", 0)
    If @error <> 0 Then Return SetError(@error, @extended, 0)
    If $aRet[0] <> 0 Then Return SetError(0xDEAD, $aRet[0], 0)
    Local $IRunningObjectTable = ObjCreateInterface($aRet[2],$IID_IRunningObjectTable,"Register HRESULT(DWORD;PTR;PTR;DWORD*);Revoke HRESULT(DWORD);IsRunning HRESULT(PTR*);GetObject HRESULT(PTR;PTR*);NoteChangeTime HRESULT(DWORD;PTR*);GetTimeOfLastChange HRESULT(PTR*;PTR*);EnumRunning HRESULT(PTR*);",True)
    If @error<>0 Then Return SetError(@error, @extended, 0)
    Return $IRunningObjectTable
EndFunc

#cs
# @internal
#ce
Func __AOI_ROT_CreateFileMoniker($sCLSID)
    Local $aRet=DllCall("Ole32.dll", "LONG", "CreateFileMoniker", "WSTR", $sCLSID, "PTR*", 0)
    If @error <> 0 Then Return SetError(@error, @extended, 0)
    If $aRet[0] <> 0 Then Return SetError(0xDEAD, $aRet[0], 0)
    ;TODO: we should reveal all methods available on IMoniker object to AutoIt3
    ;Local $IMoniker = ObjCreateInterface($aRet[2], $IID_IMoniker, "GetClassID;IsDirty;Load;Save;GetSizeMax;BindToObject;BindToStorage;Reduce;ComposeWith;Enum;IsEqual;Hash;IsRunning;GetTimeOfLastChange;Inverse;CommonPrefixWith;RelativePathTo;GetDisplayName;ParseDisplayName;IsSystemMoniker")
    Local $IMoniker = ObjCreateInterface($aRet[2], $IID_IMoniker)
    If @error<>0 Then Return SetError(@error, @extended, 0)
    Return $IMoniker
EndFunc
