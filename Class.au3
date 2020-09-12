#include-once
#include <Array.au3>

Global Const $__AU3P_ClassRegionPattern = '(?s)(?m)^\h*\QClass \E([a-zA-Z]+)$(.*?)(?=^\h*\QEndClass\E$)EndClass'

If StringRegExp(@ScriptFullPath, '\.au3p$', 0) Then
    Class_Convert_File(@ScriptFullPath)
    Exit
EndIf

Func Class_Convert_File($sFile)
    $sConvertedFile = StringRegExpReplace($sFile, '\.au3p$', '.au3', 1)

    Local $sContent = Class_Parse_File($sFile)

    If @error <> 0 Then Return SetError(@error, @extended, '')

    If FileExists($sConvertedFile) Then Return SetError(10, 0, '')

    $hFile = FileOpen($sConvertedFile, 1)
    FileWrite($hFile, $sContent)
    FileClose($hFile)

    Return $sConvertedFile
EndFunc

Func Class_Parse_File($sFile)
    If Not FileExists($sFile) Then Return SetError(1, 0, '')

    $sContent = FileRead($sFile)

    If @error <> 0 Then Return SetError(@error, @extended, '')

    $sContent = Class_Parse_Content($sContent)

    If @error <> 0 Then Return SetError(@error, @extended, '')

    Return $sContent
EndFunc

Func Class_Parse_Content($sContent)
    Return '#include "AutoItObject_Internal.au3"' & @CRLF & StringRegExpReplaceCallback($sContent, $__AU3P_ClassRegionPattern, "Class_Parse_Region")
EndFunc

#cs
# @internal
#ce
Func Class_Parse_Region($aRegion)
    Local $i, $j

    Local $sClassName = $aRegion[1]

    Local $sResult = ''

    Local Static $sRegex = '(?s)(?m)(^\h*(?:#cs|#comments-start)(?:.*?)(?=^\h*(?:#ce|#comments-end)$)\h*(?:#ce|#comments-end)|^\h*(?:(?:Get|Set)\h*)?\QFunc \E(?:[a-zA-Z0-9_]+)\([^)]*\)$(?:.*?)(?=^\h*\QEndFunc\E$)\h*EndFunc|\h*\$[a-zA-Z0-9_]+(?:\h*=\h*\N*)?$)'
    Local Static $functionPrefix = StringFormat('__Class_%s_', $sClassName)
    Local Static $getterPrefix = StringFormat('__Getter%s', $functionPrefix)
    Local Static $setterPrefix = StringFormat('__Setter%s', $functionPrefix)

    Local $aRegionShards = StringRegExp($aRegion[2], $sRegex, 3)

    Local $iRegionShards = UBound($aRegionShards, 1)

    Local $properties[$iRegionShards + 1]
    $properties[0] = 0
    Local $methods[$iRegionShards + 1]
    $methods[0] = 0
    Local $getters[$iRegionShards + 1]
    $getters[0] = 0
    Local $setters[$iRegionShards + 1]
    $setters[0] = 0
    Local $constructor = Null
    Local $constructorParameters = ""
    Local $deconstructor = Null

    Local $sRegionShard
    For $sRegionShard In $aRegionShards
        If StringRegExp($sRegionShard, "^\h*(?:#cs|#comments-start)", 0) Then ContinueLoop

        Switch (StringLower(StringRegExp($sRegionShard, "^\h*(.)", 1)[0]))
            Case '$' ; Property
                $properties[0]+=1
                $properties[$properties[0]] = $sRegionShard
            Case 'f' ; Method
                $methods[0]+=1
                $methods[$methods[0]] = $sRegionShard

                ;check if method is the constructor and add the index to the constructor ref.
                If StringRegExp($sRegionShard, '^\h*Func __construct\h*\(', 0) Then $constructor = $methods[0]
                ;check if method is the deconstructor and add the index to the constructor ref.
                If StringRegExp($sRegionShard, '^\h*Func __destruct\h*\(', 0) Then $deconstructor = $methods[0]
            Case 'g' ; Getter
                $getters[0]+=1
                $getters[$getters[0]] = $sRegionShard
            Case 's' ; Setter
                $setters[0]+=1
                $setters[$setters[0]] = $sRegionShard
            Case Else
                ;FIXME: add shard ID to the StringFormat
                ConsoleWriteError(StringFormat("WARNING: Shard #%s not recognized and will be ignored.\n", '??'))
                ContinueLoop
        EndSwitch
    Next

    If Not ($constructor = Null) Then
        $constructorParameters = StringRegExp($methods[$constructor], '^\h*Func [a-zA-Z0-9_]+\(([^)]*)\)', 1)[0]
    EndIf

    #Region Main function
    $sResult &= StringFormat('Func %s(%s)\n', $sClassName, $constructorParameters)
    $sResult &= StringFormat('\tLocal $this = IDispatch()\n')
    For $i = 1 To $properties[0] Step +1
        $sResult &= StringFormat('\t$this.%s\n', StringRegExpReplace($properties[$i], '^\h*\$', '', 1))
    Next
    Local $methodName
    For $i = 1 To $getters[0] Step +1
        $methodName = StringRegExp($getters[$i], '^\h*Get\h+Func\h*([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('\t$this.__defineGetter("%s", %s)\n', $methodName, $getterPrefix&$methodName)
    Next
    For $i = 1 To $setters[0] Step +1
        $methodName = StringRegExp($setters[$i], '^\h*Set\h+Func\h*([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('\t$this.__defineSetter("%s", %s)\n', $methodName, $setterPrefix&$methodName)
    Next
    For $i = 1 To $methods[0] Step +1
        If $i = $constructor Or $i = $deconstructor Then ContinueLoop
        $methodName = StringRegExp($methods[$i], '^\h*Func\h*([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('\t$this.__defineGetter("%s", %s)\n', $methodName, $functionPrefix&$methodName)
    Next
    If Not ($deconstructor = Null) Then
        $methodName = StringRegExp($methods[$deconstructor], '^\h*Func\h*([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('\t$this.__destructor(%s)\n', $functionPrefix&$methodName)
    EndIf
    $sResult &= StringFormat('\t$this.__seal()\n')
    If Not ($constructor = Null) Then
        $methodName = StringRegExp($methods[$constructor], '^\h*Func\h*([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('\t%s($this%s)\n', $functionPrefix&$methodName, $constructorParameters == '' ? '' : ', ' & StringRegExpReplace($constructorParameters, '(\$[^\h=,]+)\h*=[^,\)]+', '$1'));FIXME: third argument in StringFormat need to be implemented. Function arguments, without the maybe existing default value definitions
    EndIf
    $sResult &= StringFormat('\tReturn $this\n')
    $sResult &= StringFormat('EndFunc\n\n')
    #EndRegion Main function

    If Not ($constructor = Null) Then
        $methodName = StringRegExp($methods[$constructor], '^\h*Func\h*([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('Func %s($this%s)\n', $functionPrefix&$methodName, $constructorParameters == '' ? '' : ', ' & $constructorParameters)
        $sResult &= StringFormat('\t%s\n', StringRegExpReplace(StringRegExp($methods[$constructor], '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0))
        $sResult &= StringFormat('EndFunc\n\n')
    EndIf

    If Not ($deconstructor = Null) Then
        $methodName = StringRegExp($methods[$deconstructor], '^\h*Func\h*([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('Func %s($this)\n', $functionPrefix&$methodName)
        $sResult &= StringFormat('\t$this = $this.parent\n')
        $sResult &= StringFormat('\t%s\n', StringRegExpReplace(StringRegExp($methods[$deconstructor], '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0))
        $sResult &= StringFormat('EndFunc\n\n')
    EndIf

    For $i = 1 To $getters[0] Step +1
        $methodName = StringRegExp($getters[$i], '^\h*Get\h+Func\h*([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('Func %s($_oAccessorObject)\n', $getterPrefix&$methodName)
        $sResult &= StringFormat('\tLocal $this = $_oAccessorObject.parent\n')
        $sResult &= StringFormat('\t%s\n', StringRegExpReplaceCallbackEx(StringRegExpReplace(StringRegExp($getters[$i], '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0), '\$this\.([a-zA-Z0-9_]+)', 'Class_Replace_AccessorProperty', 0, $methodName))
        $sResult &= StringFormat('EndFunc\n\n')
    Next
    Local $methodParameter
    For $i = 1 To $setters[0] Step +1
        $methodName = StringRegExp($setters[$i], '^\h*Set\h+Func\h+([^\(\h]+)', 1)[0]
        $sResult &= StringFormat('Func %s($_oAccessorObject)\n', $setterPrefix&$methodName)
        $sResult &= StringFormat('\tLocal $this = $_oAccessorObject.parent\n')
        $methodParameter = StringRegExp($setters[$i], '^\h*Set\h+Func\h+[a-zA-Z0-9_]+\(([^)]*)\)', 1)
        $methodParameter = StringRegExp(UBound($methodParameter, 1) > 0 ? $methodParameter[0] : '', '^\h*\$([a-zA-Z0-9_]+)', 1)
        If @error = 0 Then
            $sResult &= StringFormat('\tLocal $%s = $_oAccessorObject.ret\n', $methodParameter[0])
        EndIf
        $sResult &= StringFormat('\t%s\n', StringRegExpReplaceCallbackEx(StringRegExpReplace(StringRegExp($setters[$i], '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0), '\$this\.([a-zA-Z0-9_]+)', 'Class_Replace_AccessorProperty', 0, $methodName))
        $sResult &= StringFormat('EndFunc\n\n')
    Next
    Local $methodParameterShards
    For $i = 1 To $methods[0] Step +1
        If $i = $constructor Or $i = $deconstructor Then ContinueLoop
        $methodName = StringRegExp($methods[$i], '^\h*Func\h*([^\(\h]+)', 1)[0]
        $methodParameter = StringRegExp($methods[$i], '^\h*Func\h+[a-zA-Z0-9_]+\(([^)]*)\)', 1)
        $methodParameter = StringSplit($methodParameter[0], ',')
        ;_ArrayDisplay($methodParameter)
        ;$methodParameter = StringRegExp(UBound($methodParameter, 1) > 0 ? $methodParameter[0] : '', '^\h*\$([a-zA-Z0-9_]+)', 1)
        $sResult &= StringFormat('Func %s($this)\n', $functionPrefix&$methodName)
        If $methodParameter[0] >= 1 And Not ($methodParameter[1] == "") Then
            For $j = 1 To $methodParameter[0] Step +1
                $methodParameterShards = StringRegExp($methodParameter[$j], '^\h*\$([a-zA-Z0-9_]+)(?:\h*=\h*(.*)$)?', 1)
                If UBound($methodParameterShards, 1) <= 1 Or $methodParameterShards[1] == "" Then
                    $sResult &= StringFormat('\tLocal $%s = $this.arguments.values[%s]\n', $methodParameterShards[0], $j-1)
                Else
                    $sResult &= StringFormat('\tLocal $%s = $this.arguments.length >= %s ? $this.arguments.values[%s] : %s\n', $methodParameterShards[0], $j, $j-1, $methodParameterShards[1])
                EndIf
            Next
        EndIf
        $sResult &= StringFormat('\t$this = $this.parent\n')
        $sResult &= StringFormat('\t%s\n', StringRegExpReplace(StringRegExp($methods[$i], '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0))
        $sResult &= StringFormat('EndFunc\n\n')
    Next

    ;_ArrayDisplay($aRegionShards)
    Return $sResult
EndFunc

Func Class_Replace_AccessorProperty($aProperty, $sName)
    Return $aProperty[1] == $sName ? '$_oAccessorObject.val' : $aProperty[0]
EndFunc

Func Class()
    ;
EndFunc

; #FUNCTION# =====================================================================================================================
; Name...........: StringRegExpReplaceCallback
; Description ...:
; Syntax.........: StringRegExpReplaceCallback($sString, $sPattern, $sFunc [, $iLimit ] )
; Parameters ....: $sString       - The input string.
;                  $sPattern      - The regular expression to compare. See StringRegExp for pattern definition characters.
;                  $sFunc         - The name of the user function to call.
;                  $iLimit        - [Optional] The max number of time to call the callback function. Default (zero) is unlimited.
; Return values .: Success        - The new string. The number of callbacks done.
;                  Failure        - Will return the original string and set the @error flag.
; Author ........: Mat
; Modified.......:
; Remarks .......: The callback function should have a single argument. This will be an array of matches, with the complete match
;                  in the first element.
; Related .......:
; Link ..........: http://www.php.net/manual/en/function.preg-replace-callback.php
; Example .......: Yes
; ================================================================================================================================
Func StringRegExpReplaceCallback($sString, $sPattern, $sFunc, $iLimit = 0)
    Local $iOffset = 1, $iDone = 0, $iMatchOffset

    While True
        $aRes = StringRegExp($sString, $sPattern, 2, $iOffset)
        If @error Then ExitLoop

        $sRet = Call($sFunc, $aRes)
        If @error Then Return SetError(@error, $iDone, $sString)

        $iOffset = StringInStr($sString, $aRes[0], 1, 1, $iOffset)
        $sString = StringLeft($sString, $iOffset - 1) & $sRet & StringMid($sString, $iOffset + StringLen($aRes[0]))
        $iOffset += StringLen($sRet)

        $iDone += 1
        If $iDone = $iLimit Then ExitLoop
    WEnd

    Return SetExtended($iDone, $sString)
EndFunc   ;==>StringRegExpReplaceCallback

Func StringRegExpReplaceCallbackEx($sString, $sPattern, $sFunc, $iLimit = 0, $vExtra = Null)
    Local $iOffset = 1, $iDone = 0, $iMatchOffset

    While True
        $aRes = StringRegExp($sString, $sPattern, 2, $iOffset)
        If @error Then ExitLoop

        $sRet = Call($sFunc, $aRes, $vExtra)
        If @error Then Return SetError(@error, $iDone, $sString)

        $iOffset = StringInStr($sString, $aRes[0], 1, 1, $iOffset)
        $sString = StringLeft($sString, $iOffset - 1) & $sRet & StringMid($sString, $iOffset + StringLen($aRes[0]))
        $iOffset += StringLen($sRet)

        $iDone += 1
        If $iDone = $iLimit Then ExitLoop
    WEnd

    Return SetExtended($iDone, $sString)
EndFunc   ;==>StringRegExpReplaceCallback
