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
    Return StringRegExpReplaceCallback($sContent, $__AU3P_ClassRegionPattern, Class_Parse_Region)
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

    Local $constructorParameters = ""

    Local $sRegionShard
    For $sRegionShard In $aRegionShards
        If StringRegExp($sRegionShard, "^\h*(?:#cs|#comments-start)", 0) Then ContinueLoop

        Switch (StringLower(StringRegExp($sRegionShard, "^\h*(.)", 1)[0]))
            Case '$' ; Property
                Local $sName = Class_Property_Get_Name($sRegionShard)
                $properties[$sName] = StringRegExpReplace($sRegionShard, '^\h*\$', '', 1)
            Case 'f' ; Method
                Local $sMethodName = Class_Function_Get_Name($sRegionShard)
                $methods[$sMethodName] = $sRegionShard
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

    For $getter In MapKeys($getters)
        If Not MapExists($properties, $getter) Then $properties[$getter] = $getter&"=Null"
    Next
    For $setter In MapKeys($setters)
        If Not MapExists($properties, $setter) Then $properties[$getter] = $setter&"=Null"
    Next

    If MapExists($methods, '__construct') Then
        $constructorParameters = StringRegExp($methods['__construct'], '^\h*Func [a-zA-Z0-9_]+\((\N*)\)', 1)[0]
    EndIf

    #Region Variant Conversion Helper
    $sResult &= "Func ___Class__"&$sClassName&"_VariantHelper()"&@CRLF
    $sResult &= 'Local Static $tVariant = DllStructCreate("ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2")'&@CRLF
    $sResult &= 'Local Static $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;")'&@CRLF
    $sResult &= 'Local Static $oObject = 0'&@CRLF
    $sResult &= 'If $oObject <> 0 Then Return $oObject'&@CRLF
    $sResult &= 'Local $hQueryInterface = DllCallbackRegister(___Class__'&$sClassName&'_VariantHelperQueryInterface, "LONG", "ptr;ptr;ptr")'&@CRLF
    $sResult &= 'Local $hAddRef = DllCallbackRegister(___Class__'&$sClassName&'_VariantHelperAddRef, "dword", "PTR")'&@CRLF
    $sResult &= 'Local $hRelease = DllCallbackRegister(___Class__'&$sClassName&'_VariantHelperRelease, "dword", "PTR")'&@CRLF
    $sResult &= 'Local $hGetTypeInfoCount = DllCallbackRegister(___Class__'&$sClassName&'_VariantHelperGetTypeInfoCount, "long", "ptr;ptr")'&@CRLF
    $sResult &= 'Local $hGetTypeInfo = DllCallbackRegister(___Class__'&$sClassName&'_VariantHelperGetTypeInfo, "long", "ptr;uint;int;ptr")'&@CRLF
    $sResult &= 'Local $hGetIDsOfNames = DllCallbackRegister(___Class__'&$sClassName&'_VariantHelperGetIDsOfNames, "long", "ptr;ptr;ptr;uint;int;ptr")'&@CRLF
    $sResult &= 'Local $hInvoke = DllCallbackRegister(___Class__'&$sClassName&'_VariantHelperInvoke, "long", "ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr")'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hQueryInterface), 1)'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hAddRef), 2)'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hRelease), 3)'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hGetTypeInfoCount), 4)'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hGetTypeInfo), 5)'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hGetIDsOfNames), 6)'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hInvoke), 7)'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "RefCount", 1) ; initial ref count is 1'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Size", 7) ; number of interface methods'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers'&@CRLF
    $sResult &= 'DllStructSetData($tObject, "Variant", DllStructGetPtr($tVariant))'&@CRLF
    $sResult &= '$oObject = ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True) ; pointer that''s wrapped into object'&@CRLF
    $sResult &= 'Return $oObject'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_VariantHelperQueryInterface($pSelf, $pRIID, $pObj)'&@CRLF
    $sResult &= 'If $pObj=0 Then Return $__AOI_E_POINTER'&@CRLF
	$sResult &= 'Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]'&@CRLF
	$sResult &= 'If (Not ($sGUID="{00020400-0000-0000-C000-000000000046}")) And (Not ($sGUID="{00000000-0000-0000-C000-000000000046}")) Then Return 0x80004002'&@CRLF
	$sResult &= 'Local $tStruct = DllStructCreate("ptr", $pObj)'&@CRLF
	$sResult &= 'DllStructSetData($tStruct, 1, $pSelf)'&@CRLF
	$sResult &= '___Class__'&$sClassName&'_VariantHelperAddRef($pSelf)'&@CRLF
	$sResult &= 'Return 0'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_VariantHelperAddRef($pSelf)'&@CRLF
	$sResult &= 'Local $tStruct = DllStructCreate("int Ref", $pSelf - 8)'&@CRLF
	$sResult &= '$tStruct.Ref += 1'&@CRLF
	$sResult &= 'Return $tStruct.Ref'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_VariantHelperRelease($pSelf)'&@CRLF
    $sResult &= 'Return 1'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_VariantHelperGetTypeInfoCount($pSelf, $pctinfo)'&@CRLF
    $sResult &= 'DllStructSetData(DllStructCreate("UINT",$pctinfo),1, 0)'&@CRLF
	$sResult &= 'Return $__AOI_S_OK'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_VariantHelperGetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)'&@CRLF
	$sResult &= 'If $iTInfo<>0 Then Return 0x8002000B'&@CRLF
	$sResult &= 'If $ppTInfo=0 Then Return 0x80070057'&@CRLF
	$sResult &= 'Return 0'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_VariantHelperGetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)'&@CRLF
	$sResult &= 'Local $tIds = DllStructCreate("long i", $rgDispId)'&@CRLF
    $sResult &= 'DllStructSetData($tIds, 1, 1)'&@CRLF
    $sResult &= 'Return 0'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_VariantHelperInvoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)'&@CRLF
    $sResult &= 'Local $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;", $pSelf - 8)'&@CRLF
    $sResult &= 'If BitAND($wFlags, 2) = 2 Then ; DISPATCH_PROPERTYGET'&@CRLF
    $sResult &= 'DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)'&@CRLF
    $sResult &= 'DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "ptr", $tObject.Variant)'&@CRLF
    $sResult &= 'Return 0'&@CRLF
    $sResult &= 'EndIf'&@CRLF
    $sResult &= 'If BitAND($wFlags, 4) = 4 Then ; DISPATCH_PROPERTYPUT'&@CRLF
    $sResult &= '$tParams = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)'&@CRLF
    $sResult &= 'If $tParams.cArgs <> 1 Then Return 0x8002000E ; DISP_E_BADPARAMCOUNT'&@CRLF
    $sResult &= 'DllCall("OleAut32.dll","LONG","VariantClear","ptr",$tObject.Variant)'&@CRLF
    $sResult &= 'DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$tObject.Variant, "ptr", $tParams.rgvargs)'&@CRLF
    $sResult &= 'Return 0'&@CRLF; S_OK
    $sResult &= 'EndIf'&@CRLF
    $sResult &= 'Return 0x80020009 ; DISP_E_EXCEPTION'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_ToVariant($vValue)'&@CRLF
	$sResult &= 'Local $oObject = ___Class__'&$sClassName&'_VariantHelper()'&@CRLF
	$sResult &= '$oObject.a = $vValue'&@CRLF
	$sResult &= 'Local $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;", ptr($oObject) - 8)'&@CRLF
	$sResult &= 'Local $tVariant = DllStructCreate("ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2")'&@CRLF
	$sResult &= 'DllCall("OleAut32.dll","LONG","VariantClear","struct*",$tVariant)'&@CRLF
	$sResult &= 'DllCall("OleAut32.dll","LONG","VariantCopy","struct*",$tVariant, "ptr", $tObject.Variant)'&@CRLF
	$sResult &= 'Return $tVariant'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    $sResult &= 'Func ___Class__'&$sClassName&'_FromVariant($pVariant)'&@CRLF
	$sResult &= 'Local $oObject = ___Class__'&$sClassName&'_VariantHelper()'&@CRLF
	$sResult &= 'Local $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;", ptr($oObject) - 8)'&@CRLF
	$sResult &= 'DllCall("OleAut32.dll","LONG","VariantClear","ptr",$tObject.Variant)'&@CRLF
	$sResult &= 'DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$tObject.Variant, "struct*", $pVariant)'&@CRLF
	$sResult &= 'Return $oObject.a'&@CRLF
    $sResult &= 'EndFunc'&@CRLF
    #EndRegion Variant Conversion Helper

    Local $sObjectStruct = StringFormat("'int RefCount;int Size;ptr Object;ptr Methods[7];ptr Properties[%s];'", UBound($properties))

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

    #Region Initilize class instance property variants
    $sResult &= 'Local Static $tVariant = DllStructCreate("ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2")'&@CRLF
    $sResult &= "DllStructSetData($tVariant, 'vt', 1)"&@CRLF; VT_NULL
    Local $i = 1
    For $property In MapKeys($properties)
        ;'___Class__'&$sClassName&'_ToVariant($vValue)'
        $sResult &= '$pObject = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", DllCall("kernel32.dll", "handle", "GlobalAlloc", "uint", 0x0002, "ulong_ptr", DllStructGetSize($tVariant))[0])[0]'&@CRLF
        If StringInStr($properties[$property], '=') Then
            ;Initial value specified for property
            $sResult &= 'DllCall("kernel32.dll", "none", "RtlMoveMemory", "struct*", $pObject, "struct*", ___Class__'&$sClassName&'_ToVariant('&StringMid($properties[$property], StringInStr($properties[$property], '=')+1)&'), "ulong_ptr", DllStructGetSize($tVariant))'&@CRLF
        Else
            $sResult &= 'DllCall("kernel32.dll", "none", "RtlMoveMemory", "struct*", $pObject, "struct*", $tVariant, "ulong_ptr", DllStructGetSize($tVariant))'&@CRLF
        EndIf
        $sResult &= 'DllStructSetData($tObject, "Properties", $pObject, '&$i&')'&@CRLF
        $i += 1
    Next
    #EndRegion

    $sResult &= 'Local $oObject = ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True)'&@CRLF ; IID_IDispatch

    If MapExists($methods, '__construct') Then
        $sResult &= StringFormat('%s__construct($oObject', $functionPrefix)
        For $parameter In MapKeys(Class_Function_Get_Parameters($methods['__construct']))
            $sResult &= ',' & $parameter
        Next
        $sResult &= ')'&@CRLF
    EndIf
    $sResult &= 'If @error <> 0 Then Return SetError(@error, @extended, 0)'&@CRLF

    $sResult &= 'Return $oObject'&@CRLF

    $sResult &= StringFormat('EndFunc\n\n')
    #EndRegion Main function

    #Region QueryInterface
        $sResult &= "Func __Object__Class_"&$sClassName&"_QueryInterface($pSelf, $pRIID, $pObj)"&@CRLF
        $sResult &= '___Class__'&$sClassName&'_VariantHelperQueryInterface($pSelf, $pRIID, $pObj)'&@CRLF
        $sResult &= 'EndFunc'&@CRLF
    #EndRegion

    #Region AddRef
        $sResult &= "Func __Object__Class_"&$sClassName&"_AddRef($pSelf)"&@CRLF
        $sResult &= 'Return ___Class__'&$sClassName&'_VariantHelperAddRef($pSelf)'&@CRLF
        $sResult &= 'EndFunc'&@CRLF
    #EndRegion

    #Region Release
        $sResult &= "Func __Object__Class_"&$sClassName&"_Release($pSelf)"&@CRLF
        $sResult &= 'Return 1'&@CRLF; FIXME: implement release memory cleanup
        $sResult &= 'EndFunc'&@CRLF
    #EndRegion

    #Region GetTypeInfoCount
        $sResult &= "Func __Object__Class_"&$sClassName&"_GetTypeInfoCount($pSelf, $pctinfo)"&@CRLF
        $sResult &= 'Return ___Class__'&$sClassName&'_VariantHelperGetTypeInfoCount($pSelf, $pctinfo)'&@CRLF
        $sResult &= 'EndFunc'&@CRLF
    #EndRegion

    #Region GetTypeInfo
        $sResult &= "Func __Object__Class_"&$sClassName&"_GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)"&@CRLF
        $sResult &= 'Return ___Class__'&$sClassName&'_VariantHelperGetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)'&@CRLF
        $sResult &= 'EndFunc'&@CRLF
    #EndRegion

    #Region GetIDsOfNames
        $sResult &= "Func __Object__Class_"&$sClassName&"_GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)"&@CRLF
        $sResult &= 'Local $tId = DllStructCreate("long i", $rgDispId)'&@CRLF
        $sResult &= 'Local $pStr = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)'&@CRLF
	    $sResult &= 'Local $s_rgszName = DllStructGetData(DllStructCreate("WCHAR[255]", $pStr), 1)'&@CRLF
        $sResult &= 'Switch $s_rgszName'&@CRLF
        Local $i = 1
        For $property In MapKeys($properties)
            $sResult &= StringFormat('Case "%s"\n', $property)
            $sResult &= StringFormat('DllStructSetData($tId, 1, %s)\n', $i)
            $i += 1
        Next

        For $method In MapKeys($methods)
            Switch $method
                Case "__construct", "__destruct"
                    ContinueLoop
                Case Else
                    $sResult &= StringFormat('Case "%s"\n', $method)
                    $sResult &= StringFormat('DllStructSetData($tId, 1, %s)\n', $i)
                    $i += 1
            EndSwitch
        Next
        $sResult &= 'Case Else'&@CRLF
        $sResult &= 'DllStructSetData($tId, 1, -1)'&@CRLF
        $sResult &= 'Return 0x80020006'&@CRLF; DISP_E_UNKNOWNNAME
        $sResult &= 'EndSwitch'&@CRLF
        $sResult &= 'Return 0'&@CRLF; S_OK
        $sResult &= 'EndFunc'&@CRLF
    #EndRegion

    #Region Invoke
        $sResult &= "Func __Object__Class_"&$sClassName&"_Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)"&@CRLF
        $sResult &= 'If $dispIdMember=-1 Then Return 0x80020003'&@CRLF; DISP_E_MEMBERNOTFOUND
        $sResult &= "$tObject = DllStructCreate("&$sObjectStruct&", $pSelf - 8)"&@CRLF
        $sResult &= 'Local Static $iVariant = DllStructGetSize(DllStructCreate("ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2"))'&@CRLF
        $sResult &= 'Switch $dispIdMember'&@CRLF
        Local $i = 1
        For $property In MapKeys($properties)
            $sResult &= StringFormat('Case %s\n', $i)
            $sResult &= 'If BitAND($wFlags, 2)=2 Then'&@CRLF ; DISPATCH_PROPERTYGET
            If MapExists($getters, $property) Then
                $soObject = 'ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True)' ; IID_IDispatch
                $sResult &= StringFormat('Local $vValue = %s%s(%s)\n', $getterPrefix, $property, $soObject)
                $sResult &= 'If @error <> 0 Then Return 0x80020009'&@CRLF; DISP_E_EXCEPTION
                $sResult &= '$tVariant = ___Class__'&$sClassName&'_ToVariant($vValue)'&@CRLF
                $sResult &= 'DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)'&@CRLF
                $sResult &= 'DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "struct*", $tVariant)'&@CRLF
            Else
                ;$sResult &= 'Local $tParams = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)'&@CRLF
                ;$sResult &= 'If $tParams.cArgs <> 1 Then Return 0x8002000E ; DISP_E_BADPARAMCOUNT'&@CRLF
                $sResult &= 'DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)'&@CRLF
                $sResult &= 'DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "ptr", DllStructGetData($tObject, "Properties", '&$i&'))'&@CRLF
            EndIf
            $sResult &= "Return 0"&@CRLF; S_OK
            $sResult &= 'EndIf'&@CRLF
            $sResult &= 'Local $tParams = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)'&@CRLF
            $sResult &= 'If $tParams.cArgs <> 1 Then Return 0x8002000E ; DISP_E_BADPARAMCOUNT'&@CRLF
            If MapExists($setters, $property) Then
                $soObject = 'ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True)' ; IID_IDispatch
                $parameter = '___Class__'&$sClassName&'_FromVariant($tParams.rgvargs)'
                $sResult &= StringFormat('%s%s(%s, %s)\n', $setterPrefix, $property, $soObject, $parameter)
                $sResult &= 'If @error <> 0 Then Return 0x80020009'&@CRLF; DISP_E_EXCEPTION
            Else
                $sResult &= 'DllCall("OleAut32.dll","LONG","VariantClear","ptr",DllStructGetData($tObject, "Properties", '&$i&'))'&@CRLF
                $sResult &= 'DllCall("OleAut32.dll","LONG","VariantCopy","ptr",DllStructGetData($tObject, "Properties", '&$i&'), "ptr", $tParams.rgvargs)'&@CRLF
            EndIf
            $sResult &= "Return 0"&@CRLF; S_OK
            $i += 1
        Next
        For $method In MapKeys($methods)
            Switch $method
                Case '__construct', '__destruct'
                    ContinueLoop
                Case Else
                    Local $parameters = Class_Function_Get_Parameters($methods[$method])
                    Local $iRequiredParameters = @extended
                    $sResult &= StringFormat('Case %s\n', $i)
                    $sResult &= 'If BitAND($wFlags, 4) = 4 Or BitAND($wFlags, 8) = 8 Then Return 0x80020009'&@CRLF; DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISP_E_EXCEPTION
                    $sResult &= '$tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)'&@CRLF; tagDISPPARAMS
                    $sResult &= 'If $tDISPPARAMS.cArgs < '&$iRequiredParameters&' Or $tDISPPARAMS.cArgs > '&UBound($parameters)&' Then Return 0x8002000E'&@CRLF; DISP_E_BADPARAMCOUNT
                    $soObject = 'ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True)' ; IID_IDispatch
                    $sResult &= 'Local $parameters[$tDISPPARAMS.cArgs + 2] = ["CallArgArray", '&$soObject&']'&@CRLF
                    $sResult &= 'Local $j = 2'&@CRLF
                    $sResult &= 'For $i=$tDISPPARAMS.cArgs-1 To 0 Step -1'&@CRLF
                    $sResult &= '$parameters[$j] = ___Class__'&$sClassName&'_FromVariant($tDISPPARAMS.rgvargs+$iVariant*$i)'&@CRLF
                    $sResult &= 'Next'&@CRLF
                    $sResult &= StringFormat('Local $vValue = Call(%s%s, $parameters)\n', $functionPrefix, $method, $soObject)
                    $sResult &= 'If @error <> 0 Then Return 0x80020009'&@CRLF; DISP_E_EXCEPTION
                    $sResult &= '$tVariant = ___Class__'&$sClassName&'_ToVariant($vValue)'&@CRLF
                    $sResult &= 'DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)'&@CRLF
                    $sResult &= 'DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "struct*", $tVariant)'&@CRLF
                    $sResult &= "Return 0"&@CRLF; S_OK
                    $i += 1
            EndSwitch
        Next
        $sResult &= 'EndSwitch'&@CRLF
        $sResult &= 'EndFunc'&@CRLF
    #EndRegion

    If MapExists($methods, '__construct') Then
        $sResult &= Class_Make_Constructor($methods['__construct'], $functionPrefix, $constructorParameters)
    EndIf

    If MapExists($methods, '__destruct') Then
        $sResult &= Class_Make_Desctructor($methods['__destruct'], $functionPrefix)
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
        Switch $method
            Case '__construct', '__destruct'
                ContinueLoop
            Case Else
        $sResult &= Class_Make_Method($methods[$method], $functionPrefix)
        EndSwitch
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
    Local $sResult = '', _
        $methodName = Class_Getter_Get_Name($sSource), _
        $parameters = Class_Function_Get_Parameters(StringRegExpReplace($sSource,'(?i)^\h*Get', '')), _
        $iRequiredParameters = @extended

    If $iRequiredParameters > 0 Then ConsoleWriteError(StringFormat("WARNING! getter for ""%s"" have one or more required parameters. This will cause a crash when called!\n", $methodName))

    $sResult &= StringFormat('Func %s($this', $getterPrefix&$methodName)
    For $parameter In MapKeys($parameters)
        $sResult &= ',' & $parameter
        If $parameters[$parameter] = "" Then ContinueLoop
        $sResult &= '=' & $parameters[$parameter]
    Next
    $sResult &= ')'&@CRLF
    $sResult &= StringRegExpReplace(StringRegExp($sSource, '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0)
    $sResult &= StringFormat('\nEndFunc\n\n')

    Return $sResult
EndFunc

Func Class_Make_Setter($sSource, $setterPrefix)
    Local $sResult, _
        $methodName = Class_Setter_Get_Name($sSource), _
        $parameters = Class_Function_Get_Parameters(StringRegExpReplace($sSource,'^\h*Set', '')), _
        $iRequiredParameters = @extended

    If $iRequiredParameters > 1 Then ConsoleWriteError(StringFormat("WARNING! setter for ""%s"" have more than one required parameter. This will cause a crash when called!\n", $methodName))

    $sResult &= StringFormat('Func %s($this', $setterPrefix&$methodName)
    For $parameter In MapKeys($parameters)
        $sResult &= ',' & $parameter
        If $parameters[$parameter] = "" Then ContinueLoop
        $sResult &= '=' & $parameters[$parameter]
    Next
    $sResult &= ')'&@CRLF
    Local $methodParameter = StringRegExp($sSource, '^\h*Set\h+Func\h+[a-zA-Z0-9_]+\((\N*)\)', 1)
    $methodParameter = StringRegExp(UBound($methodParameter, 1) > 0 ? $methodParameter[0] : '', '^\h*\$([a-zA-Z0-9_]+)', 1)
    $sResult &= StringRegExpReplace(StringRegExp($sSource, '(?s)^.*?\N+(.*)\N+\h*EndFunc\h*$', 1)[0], '(^(\h|\R)*|(\h|\R)*$)', '', 0)
    $sResult &= StringFormat('\nEndFunc\n\n')

    Return $sResult
EndFunc

Func Class_Make_Method($sSource, $functionPrefix)
    Local $sResult, _
        $methodName = Class_Function_Get_Name($sSource), _
        $methodParameters = Class_Function_Get_Parameters($sSource)

    $sResult &= StringFormat('Func %s($this', $functionPrefix&$methodName)
    For $parameter In MapKeys($methodParameters)
        $sResult &= ',' & $parameter & ($methodParameters[$parameter] = "" ? '' : '='&$methodParameters[$parameter])
    Next
    $sResult &= ')'&@CRLF
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
    Return StringRegExp($sSource, '^\h*\$([_a-zA-Z0-9]+)', 1)[0]
EndFunc

Func Class_Function_Get_Parameters($sSource)
    ;Local Static $sDefine = '(?(DEFINE)(?<method>[_a-zA-Z][_a-zA-Z0-9]*(?&ws)?\(((?&parameter)((?&ws)?,(?&ws)?(?&parameter))*)?\))(?<parameter>(?&method)|[^(),]+)(?<ws>\h+))'
    ;Local Static $sDefine = '(?(DEFINE)(?<method>[_a-zA-Z][_a-zA-Z0-9]*(?&ws)?\(((?&parameter)((?&ws)?,(?&ws)?(?&parameter))*)?\))(?<parameter>(?&variable)((?&ws)?=(?&ws)?(?&value))?)(?<variable>\$[_a-zA-Z0-9]+)(?<value>(?&method)|[^(),]+)(?<ws>\h+))'
    Local Static $sDefine = '(?(DEFINE)(?<function>[_a-zA-Z][_a-zA-Z0-9]*(?&ws)?\(((?&parameter)((?&ws)?,(?&ws)?(?&parameter))*)?\))(?<parameter>(?&variable)((?&ws)?=(?&ws)?(?&value))?)(?<variable>\$[_a-zA-Z0-9]+)(?<value>(?&call)|[^(),]+)(?<call>[_a-zA-Z][_a-zA-Z0-9]*(?&ws)?\(((?&value)((?&ws)?,(?&ws)?(?&value))*)?\))(?<ws>\h+))'
    Local $mParameters[]
    ; Extract function name and parameters part of string
    Local $parameters = StringRegExp($sSource, $sDefine&'^\h*Func(?&ws)((?&function))', 1)
    ; Extract parameters part of string
    $parameters = StringRegExp($parameters[UBound($parameters) - 1], $sDefine&'^[_a-zA-Z]+(?&ws)?\(((?&parameter)(?:(?&ws)?,(?&ws)?(?&parameter))*)\)', 1)
    If @error = 1 Then Return $mParameters; No parameters in method
    ; Extract each parameter as key and value
    $parameters = StringRegExp($parameters[UBound($parameters) - 1], $sDefine&'\G(?:^|, )((?&variable))(?:(?&ws)?=(?&ws)?((?&value)))?', 3)

    $iRequiredParameters = 0
    ; Due to bug/feature of current PCRE implementation, the groups within DEFINE still produces empty matches, per group. Currently that is 11 empty matches before the expected match.
    For $i = 0 To UBound($parameters) - 1 Step + 13
        ; NOTE: Using execute, to prevent out of bounds crashing, if last parameter does not have a default value. I could also use tenerary statement with count, but execute is used for now.
        $mParameters[$parameters[$i+11]] = Execute('$parameters[$i+12]')
        If Execute('$parameters[$i+12]') = "" Then
            $i -= 1
            $iRequiredParameters += 1
        EndIf
    Next

    Return SetExtended($iRequiredParameters, $mParameters)
EndFunc
