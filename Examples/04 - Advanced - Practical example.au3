#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         genius257

#ce ----------------------------------------------------------------------------
#include <GDIPlus.au3>
#include "..\AutoItObject_Internal.au3"

$AutoItError = ObjEvent("AutoIt.Error", "ErrFunc") ; Install a custom error handler
Func ErrFunc($oError)
	ConsoleWrite("!>COM Error !"&@CRLF&"!>"&@TAB&"Number: "&Hex($oError.Number,8)&@CRLF&"!>"&@TAB&"Windescription: "&StringRegExpReplace($oError.windescription,"\R$","")&@CRLF&"!>"&@TAB&"Source: "&$oError.source&@CRLF&"!>"&@TAB&"Description: "&$oError.description&@CRLF&"!>"&@TAB&"Helpfile: "&$oError.helpfile&@CRLF&"!>"&@TAB&"Helpcontext: "&$oError.helpcontext&@CRLF&"!>"&@TAB&"Lastdllerror: "&$oError.lastdllerror&@CRLF&"!>"&@TAB&"Scriptline: "&$oError.scriptline&@CRLF)
EndFunc ;==>ErrFunc

_GDIPlus_Startup()

$GUI = GUI("old title")
$GUI.bkColor(0xC2E34E).width(200).height(200).title("new title").onExit(_MyExit).Show

$Pen = Pen(0xFF000000)
$Brush = Brush(0xFF000000)

$GUI.graphics.DrawRect(10,10,$GUI.clientWidth-21,$GUI.clientHeight-21,$Pen).FillRect(15, 15, $GUI.clientWidth-30, $GUI.clientHeight-30, $Brush)

Sleep(1000)

$GUI.graphics.FillRect(15, 15, $GUI.clientWidth-30, $GUI.clientHeight-30, $Brush.color(0xFFc3E3c3))

While 1
	Sleep(10)
WEnd

Func _MyExit()
	;remove object references to release resources and activate desctructors
	$Pen=0
	$Brush=0
	$GUI=0
	_GDIPlus_Shutdown()
	Exit
EndFunc

Func GUI($title, $width = Default, $height = Default, $left = Default, $top = Default, $style = Default, $exStyle = Default, $parent = Default)
	Local $IDispatch = IDispatch()
	$IDispatch.hwnd = GUICreate($title, $width, $height, $left, $top, $style, $exStyle, $parent)
	$IDispatch.__defineSetter("hwnd", PrivateProperty);PrivateProperty is defined in AutoItObject_Internal.au3
	$IDispatch.__defineGetter("Show", Wnd_Show)
	$IDispatch.__defineGetter("Hide", Wnd_Hide)
	$IDispatch.__defineGetter("bkColor", Wnd_bkColor)
	$IDispatch.__defineGetter("width", Wnd_width)
	$IDispatch.__defineGetter("height", Wnd_height)
	$IDispatch.__defineGetter("clientWidth", Wnd_clientWidth)
	$IDispatch.__defineGetter("clientHeight", Wnd_clientHeight)
	$IDispatch.__defineGetter("title", Wnd_title)
	$IDispatch.__defineGetter("onExit", Wnd_onExit)
	$IDispatch.__defineGetter("graphics", Wnd_graphics)
	$IDispatch.__destructor(Wnd_Destructor)
	$IDispatch.__preventExtensions
	Return $IDispatch
EndFunc

Func Wnd_Destructor($oSelf)
	$oSelf.graphics=0
EndFunc

Func Wnd_Show($oSelf)
	GUISetState(@SW_SHOW, $oSelf.parent.hwnd)
	Return $oSelf.parent
EndFunc

Func Wnd_Hide($oSelf)
	GUISetState(@SW_HIDE, $oSelf.parent.hwnd)
	Return $oSelf.parent
EndFunc

Func Wnd_bkColor($oSelf)
	If Not ($oSelf.arguments.length==1) Then Return SetError(1, 1, $oSelf.parent)
	GUISetBkColor($oSelf.arguments.values[0], $oSelf.parent.hwnd)
	Return $oSelf.parent
EndFunc

Func Wnd_width($oSelf)
	If Not ($oSelf.arguments.length==1) Then Return SetError(1, 1, $oSelf.parent)
	Local $aPos = WinGetPos(ptr($oSelf.parent.hwnd), "")
	WinMove(ptr($oSelf.parent.hwnd), "", $aPos[0], $aPos[1], $oSelf.arguments.values[0], $aPos[3], 0)
	Return $oSelf.parent
EndFunc

Func Wnd_height($oSelf)
	If Not ($oSelf.arguments.length==1) Then Return SetError(1, 1, $oSelf.parent)
	Local $aPos = WinGetPos(ptr($oSelf.parent.hwnd))
	WinMove(ptr($oSelf.parent.hwnd), "", $aPos[0], $aPos[1], $aPos[2], $oSelf.arguments.values[0], 0)
	Return $oSelf.parent
EndFunc

Func Wnd_clientWidth($oSelf)
	If ($oSelf.arguments.length==1) Then
		Local $aPos = WinGetPos(ptr($oSelf.parent.hwnd))
		Local $tRECT01 = _WinAPI_GetWindowRect($oSelf.parent.hwnd)
		Local $tRECT02 = _WinAPI_GetClientRect($oSelf.parent.hwnd)
		WinMove(ptr($oSelf.parent.hwnd), "", $aPos[0], $aPos[1], (($tRECT01.Right-$tRECT02.Right)-($tRECT01.Left-$tRECT02.Left))+$oSelf.arguments.values[0], $aPos[3], 0)
		Return SetError(@error, @extended, $oSelf.parent)
	EndIf
	Return _WinAPI_GetClientWidth($oSelf.parent.hwnd)
EndFunc


Func Wnd_clientHeight($oSelf)
	If ($oSelf.arguments.length==1) Then
		Local $aPos = WinGetPos(ptr($oSelf.parent.hwnd))
		Local $tRECT01 = _WinAPI_GetWindowRect($oSelf.parent.hwnd)
		Local $tRECT02 = _WinAPI_GetClientRect($oSelf.parent.hwnd)
		WinMove(ptr($oSelf.parent.hwnd), "", $aPos[0], $aPos[1], $aPos[2], (($tRECT01.Bottom-$tRECT02.Bottom)-($tRECT01.Top-$tRECT02.Top))+$oSelf.arguments.values[0], 0)
		Return SetError(@error, @extended, $oSelf.parent)
	EndIf
	Return _WinAPI_GetClientHeight($oSelf.parent.hwnd)
EndFunc

Func Wnd_title($oSelf)
	If Not ($oSelf.arguments.length==1) Then Return SetError(1, 1, $oSelf.parent)
	WinSetTitle(ptr($oSelf.parent.hwnd), "", $oSelf.arguments.values[0])
	Return $oSelf.parent
EndFunc

Func Wnd_onExit($oSelf)
	If Not ($oSelf.arguments.length==1) Then Return SetError(1, 1, $oSelf.parent)
	opt("GuiOnEventMode", 1)
	GUISetOnEvent(-3, $oSelf.arguments.values[0], ptr($oSelf.parent.hwnd))
	Return $oSelf.parent
EndFunc

Func Wnd_graphics($oSelf)
	If Not IsObj($oSelf.val) Then
		Local $IDispatch = IDispatch()
		$IDispatch.hwnd = _GDIPlus_GraphicsCreateFromHWND($oSelf.parent.hwnd)
		$IDispatch.__defineSetter("hwnd", PrivateProperty);PrivateProperty is defined in AutoItObject_Internal.au3
		$IDispatch.__defineSetter("parent", PrivateProperty);PrivateProperty is defined in AutoItObject_Internal.au3
		$IDispatch.__defineGetter("Clear", Graphics_Clear)
		$IDispatch.__defineGetter("DrawRect", Graphics_DrawRect)
		$IDispatch.__defineGetter("FillRect", Graphics_FillRect)
		$IDispatch.__destructor(Graphics_Dispose)
		$IDispatch.__preventExtensions
		$oSelf.val = $IDispatch
		Return $IDispatch
	EndIf
	Return $oSelf.val
EndFunc

Func Graphics_Clear($oSelf)
	_GDIPlus_GraphicsClear($oSelf.parent.hwnd, ($oSelf.arguments.length>0)?$oSelf.arguments.values[0]:Default)
	Return SetError(@error, @extended, $oSelf.parent)
EndFunc

Func Graphics_Dispose($oSelf)
	_GDIPlus_GraphicsDispose($oSelf.hwnd)
EndFunc

Func Graphics_DrawRect($oSelf)
;~ 	MsgBox(0, "", $oSelf.parent.hwnd&@CRLF&$oSelf.arguments.values[4].hwnd)
	_GDIPlus_GraphicsDrawRect($oSelf.parent.hwnd, $oSelf.arguments.values[0], $oSelf.arguments.values[1], $oSelf.arguments.values[2], $oSelf.arguments.values[3], $oSelf.arguments.values[4].hwnd)
	Return SetError(@error, @extended, $oSelf.parent)
EndFunc

Func Graphics_FillRect($oSelf)
	_GDIPlus_GraphicsFillRect($oSelf.parent.hwnd, $oSelf.arguments.values[0], $oSelf.arguments.values[1], $oSelf.arguments.values[2], $oSelf.arguments.values[3], $oSelf.arguments.values[4].hwnd)
	Return SetError(@error, @extended, $oSelf.parent)
EndFunc

Func Pen($color=Default, $width=Default)
	Local $IDispatch = IDispatch()
	$IDispatch.hwnd = _GDIPlus_PenCreate($color, $width)
	$IDispatch.__defineSetter("hwnd", PrivateProperty);PrivateProperty is defined in AutoItObject_Internal.au3
	$IDispatch.__defineGetter("color", Pen_color)
	$IDispatch.__destructor(Pen_Dispose)
	$IDispatch.__preventExtensions
	Return $IDispatch
EndFunc

Func Pen_color($oSelf)
	If $oSelf.arguments.length=0 Then Return _GDIPlus_PenGetColor($oSelf.parent.hwnd)
	_GDIPlus_PenSetColor($oSelf.parent.hwnd, $oSelf.arguments.values[0])
	Return SetError(@error, @extended, $oSelf.parent)
EndFunc

Func Pen_Dispose($oSelf)
	_GDIPlus_PenDispose($oSelf.hwnd)
EndFunc

Func Brush($color=Default)
	Local $IDispatch = IDispatch()
	$IDispatch.hwnd = _GDIPlus_BrushCreateSolid($color)
	$IDispatch.__defineSetter("hwnd", PrivateProperty);PrivateProperty is defined in AutoItObject_Internal.au3
	$IDispatch.__defineGetter("color", Brush_color)
	$IDispatch.__destructor(Brush_Dispose)
	$IDispatch.__preventExtensions
	Return $IDispatch
EndFunc

Func Brush_color($oSelf)
	If $oSelf.arguments.length=0 Then Return _GDIPlus_BrushGetSolidColor($oSelf.parent.hwnd)
	_GDIPlus_BrushSetSolidColor($oSelf.parent.hwnd, $oSelf.arguments.values[0])
	Return SetError(@error, @extended, $oSelf.parent)
EndFunc

Func Brush_Dispose($oSelf)
	_GDIPlus_BrushDispose($oSelf.hwnd)
EndFunc