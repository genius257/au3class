#include-once
#include <Array.au3>

Global Const $__AU3P_ClassRegionPattern = '(?s)(?m)^\h*\QClass \E([a-zA-Z0-9_]+)$(.*?)(?=^\h*\QEndClass\E$)EndClass'

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
    Return '#include "AutoItObject_Internal.au3"' & @CRLF & StringRegExpReplaceCallback($sContent, $__AU3P_ClassRegionPattern, Class_Parse_Region)
EndFunc

#cs
# @internal
#ce
Func Class_Parse_Region($aRegion)
    Local $i, $j

    Local $sClassName = $aRegion[1]

    Local $sResult = ''

    Local Static $sRegex = '(?s)(?m)(^\h*(?:#cs|#comments-start)(?:.*?)(?=^\h*(?:#ce|#comments-end)$)\h*(?:#ce|#comments-end)|^\h*(?:(?:Get|Set)\h*)?\QFunc \E(?:[a-zA-Z0-9_]+)\(\N*\)$(?:.*?)(?=^\h*\QEndFunc\E$)\h*EndFunc|\h*\$[a-zA-Z0-9_]+(?:\h*=\h*\N*)?$)'
    Local $functionPrefix = StringFormat('__Class_%s_', $sClassName)
    Local $getterPrefix = StringFormat('__Getter%s', $functionPrefix)
    Local $setterPrefix = StringFormat('__Setter%s', $functionPrefix)

    Local $aRegionShards = StringRegExp($aRegion[2], $sRegex, 3)

    Local $properties[], _
        $methods[], _
        $getters[], _
        $setters[]

    Local $constructor = Null
    Local $constructorParameters = ""
    Local $deconstructor = Null

    Local $sRegionShard
    For $sRegionShard In $aRegionShards
        If StringRegExp($sRegionShard, "^\h*(?:#cs|#comments-start)", 0) Then ContinueLoop

        Switch (StringLower(StringRegExp($sRegionShard, "^\h*(.)", 1)[0]))
            Case '$' ; Property
                Local $sName = Class_Property_Get_Name($sRegionShard)
                $properties[$sName] = $sRegionShard
            Case 'f' ; Method
                Local $sMethodName = Class_Function_Get_Name($sRegionShard)
                $methods[$sMethodName] = $sRegionShard

                ;check if method is the constructor and add the index to the constructor ref.
                If StringRegExp($sRegionShard, '^\h*Func __construct\h*\(', 0) Then $constructor = $sMethodName
                ;check if method is the deconstructor and add the index to the constructor ref.
                If StringRegExp($sRegionShard, '^\h*Func __destruct\h*\(', 0) Then $deconstructor = $sMethodName
            Case 'g' ; Getter
                Local $sMethodName = Class_Getter_Get_Name($sRegionShard)
                $getters[$sMethodName] = $sRegionShard
            Case 's' ; Setter
                Local $sMethodName = Class_Setter_Get_Name($sRegionShard)
                $setters[$sMethodName] = $sRegionShard
            Case Else
                ;FIXME: add shard ID to the StringFormat
                ConsoleWriteError(StringFormat("WARNING: Shard #%s not recognized and will be ignored.\n", '??'))
                ContinueLoop
        EndSwitch
    Next

    If Not ($constructor = Null) Then
        $constructorParameters = StringRegExp($methods[$constructor], '^\h*Func [a-zA-Z0-9_]+\((\N*)\)', 1)[0]
    EndIf

    Local $sObjectStruct = StringFormat("'int RefCount;int RefCount;int Size;ptr Object;ptr Methods[7];ptr Properties[%s];'", UBound($properties))

    #Region Main function
    $sResult &= StringFormat('Func %s(%s)\n', $sClassName, $constructorParameters)

    $sResult &= "Local Static $QueryInterface = DllCallbackRegister(__Object__Class_"&$sClassName&"_QueryInterface, 'LONG', 'ptr;ptr;ptr')"
    $sResult &= ", $AddRef = DllCallbackRegister(__Object__Class_"&$sClassName&"_AddRef, 'dword', 'PTR')"
    $sResult &= ", $Release = DllCallbackRegister(__Object__Class_"&$sClassName&"_Release, 'dword', 'PTR')"
    $sResult &= ", $GetTypeInfoCount = DllCallbackRegister(__Object__Class_"&$sClassName&"_GetTypeInfoCount, 'long', 'ptr;ptr')"
    $sResult &= ", $GetTypeInfo = DllCallbackRegister(__Object__Class_"&$sClassName&"_GetTypeInfo, 'long', 'ptr;uint;int;ptr')"
    $sResult &= ", $GetIDsOfNames = DllCallbackRegister(__Object__Class_"&$sClassName&"_GetIDsOfNames, 'long', 'ptr;ptr;ptr;uint;int;ptr')"
    $sResult &= ", $Invoke = DllCallbackRegister(__Object__Class_"&$sClassName&"_Invoke, 'long', 'ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr')"
    $sResult &= @CRLF
    $sResult &= "$tObject = DllStructCreate("&$sObjectStruct&")"&@CRLF
    $sResult &= "DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($QueryInterface), 1)"&@CRLF
	$sResult &= "DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($AddRef), 2)"&@CRLF
	$sResult &= "DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($Release), 3)"&@CRLF
	$sResult &= "DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($GetTypeInfoCount), 4)"&@CRLF
	$sResult &= "DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($GetTypeInfo), 5)"&@CRLF
	$sResult &= "DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($GetIDsOfNames), 6)"&@CRLF
	$sResult &= "DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($Invoke), 7)"&@CRLF
    $sResult &= "DllStructSetData($tObject, 'RefCount', 1)"&@CRLF ; initial ref count is 1
	$sResult &= "DllStructSetData($tObject, 'Size', 7)"&@CRLF ; number of interface methods

    $sResult &= '$pObject = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", DllCall("kernel32.dll", "handle", "GlobalAlloc", "uint", 0x0002, "ulong_ptr", DllStructGetSize($tObject))[0])[0]'&@CRLF
    $sResult &= 'DllCall("kernel32.dll", "none", "RtlMoveMemory", "struct*", $pObject, "struct*", $tObject, "ulong_ptr", DllStructGetSize($tObject))'&@CRLF
    $sResult &= "$tObject = DllStructCreate("&$sObjectStruct&", $pObject)"&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods"))'&@CRLF
    $sResult &= 'Return ObjCreateInterface(DllStructGetPtr($tObject, "Object"), $__AOI_IID_IDispatch, Default, True)'&@CRLF

    $sResult &= StringFormat('\tLocal $this = IDispatch()\n')

    $sResult &= StringFormat("; Properties\n")
    For $property In MapKeys($properties)
        $sResult &= StringFormat('\t$this.%s\n', $property)
    Next

    Local $methodName
    $sResult &= StringFormat("; Getters\n")
    For $getter In MapKeys($getters)
        $methodName = $getter
        $sResult &= StringFormat('\t$this.__defineGetter("%s", %s)\n', $methodName, $getterPrefix&$methodName)
    Next

    $sResult &= StringFormat("; Setters\n")
    For $setter In MapKeys($setters)
        $methodName = $setter
        $sResult &= StringFormat('\t$this.__defineSetter("%s", %s)\n', $methodName, $setterPrefix&$methodName)
    Next

    $sResult &= StringFormat("; Methods\n")
    For $method In MapKeys($methods)
        If $method = $constructor Or $method = $deconstructor Then ContinueLoop
        $sResult &= StringFormat('\t$this.__defineGetter("%s", %s)\n', $method, $functionPrefix&$method)
    Next

    $sResult &= StringFormat("; Deconstructor\n")
    If Not ($deconstructor = Null) Then
        $methodName = $deconstructor
        $sResult &= StringFormat('\t$this.__destructor(%s)\n', $functionPrefix&$methodName)
    EndIf
    
    $sResult &= StringFormat("; Seal object, to prevent dynamic property declaration\n")
    $sResult &= StringFormat('\t$this.__seal()\n')
    $sResult &= StringFormat("; Constructor\n")
    If Not ($constructor = Null) Then
        $methodName = $constructor
        $sResult &= StringFormat('\t%s($this%s)\n', $functionPrefix&$methodName, $constructorParameters == '' ? '' : ', ' & StringRegExpReplace($constructorParameters, '(\$[^\h=,]+)\h*=[^,]+', '$1'));FIXME: third argument in StringFormat need to be implemented. Function arguments, without the maybe existing default value definitions
        $sResult &= StringFormat('\tIf @error <> 0 Then Return SetError(@error, @extended, $this)\n')
    EndIf

    $sResult &= StringFormat('\tReturn $this\n')
    $sResult &= StringFormat('EndFunc\n\n')
    #EndRegion Main function

    If Not ($constructor = Null) Then
        $sResult &= Class_Make_Constructor($methods[$constructor], $functionPrefix, $constructorParameters)
    EndIf

    If Not ($deconstructor = Null) Then
        $sResult &= Class_Make_Desctructor($methods[$deconstructor], $functionPrefix)
    EndIf

    For $getter In $getters
        $sResult &= Class_Make_Getter($getter, $getterPrefix)
    Next

    Local $methodParameter
    For $setter In $setters
        $sResult &= Class_Make_Setter($setter, $setterPrefix)
    Next

    Local $methodParameterShards
    For $method In MapKeys($methods)
        If $method = $constructor Or $method = $deconstructor Then ContinueLoop
        $sResult &= Class_Make_Method($methods[$method], $functionPrefix)
    Next

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

Func Class_Make_Constructor($sSource, $functionPrefix, $constructorParameters)
    $methodName = Class_Function_Get_Name($sSource)
    Return _
        StringFormat('Func %s($this%s)\n', $functionPrefix&$methodName, $constructorParameters == '' ? '' : ', ' & $constructorParameters) & _
        StringFormat('\t%s\n', StringRegExpReplace(StringRegExp($sSource, '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0)) & _
        StringFormat('EndFunc\n\n')
EndFunc

Func Class_Make_Desctructor($sSource, $functionPrefix)
    $methodName = Class_Function_Get_Name($sSource)
    Return _
        StringFormat('Func %s($this)\n', $functionPrefix&$methodName) & _
        StringFormat('\t%s\n', StringRegExpReplace(StringRegExp($sSource, '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0)) & _
        StringFormat('EndFunc\n\n')
EndFunc

Func Class_Make_Getter($sSource, $getterPrefix)
    $methodName = Class_Getter_Get_Name($sSource)
    Return _
        StringFormat('Func %s($_oAccessorObject)\n', $getterPrefix&$methodName) & _
        StringFormat('\tLocal $this = $_oAccessorObject.parent\n') & _
        StringFormat('\t%s\n', StringRegExpReplaceCallbackEx(StringRegExpReplace(StringRegExp($sSource, '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0), '\$this\.([a-zA-Z0-9_]+)', 'Class_Replace_AccessorProperty', 0, $methodName)) & _
        StringFormat('EndFunc\n\n')
EndFunc

Func Class_Make_Setter($sSource, $setterPrefix)
    Local $sResult, _
        $methodName = Class_Setter_Get_Name($sSource)

    $sResult &= StringFormat('Func %s($_oAccessorObject)\n', $setterPrefix&$methodName)
    $sResult &= StringFormat('\tLocal $this = $_oAccessorObject.parent\n')
    Local $methodParameter = StringRegExp($sSource, '^\h*Set\h+Func\h+[a-zA-Z0-9_]+\((\N*)\)', 1)
    $methodParameter = StringRegExp(UBound($methodParameter, 1) > 0 ? $methodParameter[0] : '', '^\h*\$([a-zA-Z0-9_]+)', 1)
    If @error = 0 Then
        $sResult &= StringFormat('\tLocal $%s = $_oAccessorObject.ret\n', $methodParameter[0])
    EndIf
    $sResult &= StringFormat('\t%s\n', StringRegExpReplaceCallbackEx(StringRegExpReplace(StringRegExp($sSource, '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0), '\$this\.([a-zA-Z0-9_]+)', 'Class_Replace_AccessorProperty', 0, $methodName))
    $sResult &= StringFormat('EndFunc\n\n')

    Return $sResult
EndFunc

Func Class_Make_Method($sSource, $functionPrefix)
    Local $sResult, _
        $methodName = Class_Function_Get_Name($sSource), _
        $methodParameter = StringRegExp($sSource, '^\h*Func\h+[a-zA-Z0-9_]+\((\N*)\)', 1)
    
    $methodParameter = StringSplit($methodParameter[0], ',')
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
    $sResult &= StringFormat('\t%s\n', StringRegExpReplace(StringRegExp($sSource, '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0))
    $sResult &= StringFormat('EndFunc\n\n')

    Return $sResult
EndFunc

Func Class_Function_Get_Name($sSource)
    Return StringRegExp($sSource, '^\h*Func\h*([^\(\h]+)', 1)[0]
EndFunc

Func Class_Getter_Get_Name($sSource)
    Return StringRegExp($sSource, '^\h*Get\h+Func\h*([^\(\h]+)', 1)[0]
EndFunc

Func Class_Setter_Get_Name($sSource)
    Return StringRegExp($sSource, '^\h*Set\h+Func\h+([^\(\h]+)', 1)[0]
EndFunc

Func Class_Property_Get_Name($sSource)
    Return StringRegExpReplace($sSource, '^\h*\$', '', 1)
EndFunc
