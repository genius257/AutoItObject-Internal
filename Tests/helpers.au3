#include-once
#include <StringConstants.au3>
;~ Func IDispatch_

;http://php.net/manual/en/function.str-replace.php
Func PHP_str_replace($search, $replace, $subject)
	Local $replacements = 0, $i, $j
	Local $Tsearch = IsArray($search)
	$search = ArrayWrap($search)
	Local $Treplace = IsArray($replace)
	$replace = ArrayWrap($replace)

	If (Not $Tsearch) And $Treplace Then Dim $replace = ["Array"]
	If $Tsearch And (Not $Treplace) Then
		ReDim $replace[UBound($search)]
		For $i=1 To UBound($replace)-1
			$replace[$i]=$replace[0]
		Next
	EndIf
	If UBound($search) > UBound($replace) Then
		Local $Ireplace = UBound($replace)
		ReDim $replace[UBound($search)]
		For $i=$Ireplace To UBound($replace)-1
			$replace[$i] = ""
		Next
	EndIf

	Local $Tsubject = IsArray($subject)
	$subject = ArrayWrap($subject)
	For $i=0 To UBound($subject)-1
		If Not IsString($subject[$i]) Then Return SetError(1, $replacements, $Tsubject?$subject:$subject[0])
		For $j=0 To UBound($search)-1
			$subject[$i] = StringReplace($subject[$i], $search[$j], $replace[$j], 0, $STR_CASESENSE)
			$replacements+=@extended
			If @error<>0 Then Return SetError(@error, $replacements, $Tsubject?$subject:$subject[0])
		Next
	Next

	Return SetError(0, $replacements, $Tsubject?$subject:$subject[0])
EndFunc

Func ArrayWrap($value)
	If IsArray($value) Then Return $value
	Local $return = [$value]
	Return $return
EndFunc