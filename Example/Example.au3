#include-once

Func ___Class__Example_VariantHelper()
Local Static $tVariant = DllStructCreate("ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2")
Local Static $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;")
Local Static $oObject = 0
If $oObject <> 0 Then Return $oObject
Local $hQueryInterface = DllCallbackRegister(___Class__Example_VariantHelperQueryInterface, "LONG", "ptr;ptr;ptr")
Local $hAddRef = DllCallbackRegister(___Class__Example_VariantHelperAddRef, "dword", "PTR")
Local $hRelease = DllCallbackRegister(___Class__Example_VariantHelperRelease, "dword", "PTR")
Local $hGetTypeInfoCount = DllCallbackRegister(___Class__Example_VariantHelperGetTypeInfoCount, "long", "ptr;ptr")
Local $hGetTypeInfo = DllCallbackRegister(___Class__Example_VariantHelperGetTypeInfo, "long", "ptr;uint;int;ptr")
Local $hGetIDsOfNames = DllCallbackRegister(___Class__Example_VariantHelperGetIDsOfNames, "long", "ptr;ptr;ptr;uint;int;ptr")
Local $hInvoke = DllCallbackRegister(___Class__Example_VariantHelperInvoke, "long", "ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr")
DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hQueryInterface), 1)
DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hAddRef), 2)
DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hRelease), 3)
DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hGetTypeInfoCount), 4)
DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hGetTypeInfo), 5)
DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hGetIDsOfNames), 6)
DllStructSetData($tObject, "Methods", DllCallbackGetPtr($hInvoke), 7)
DllStructSetData($tObject, "RefCount", 1) ; initial ref count is 1
DllStructSetData($tObject, "Size", 7) ; number of interface methods
DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods")) ; Interface method pointers
DllStructSetData($tObject, "Variant", DllStructGetPtr($tVariant))
$oObject = ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True) ; pointer that's wrapped into object
Return $oObject
EndFunc
Func ___Class__Example_VariantHelperQueryInterface($pSelf, $pRIID, $pObj)
If $pObj=0 Then Return $__AOI_E_POINTER
Local $sGUID=DllCall("ole32.dll", "int", "StringFromGUID2", "PTR", $pRIID, "wstr", "", "int", 40)[2]
If (Not ($sGUID="{00020400-0000-0000-C000-000000000046}")) And (Not ($sGUID="{00000000-0000-0000-C000-000000000046}")) Then Return -2147467262
Local $tStruct = DllStructCreate("ptr", $pObj)
DllStructSetData($tStruct, 1, $pSelf)
___Class__Example_VariantHelperAddRef($pSelf)
Return 0
EndFunc
Func ___Class__Example_VariantHelperAddRef($pSelf)
Local $tStruct = DllStructCreate("int Ref", $pSelf - 8)
$tStruct.Ref += 1
Return $tStruct.Ref
EndFunc
Func ___Class__Example_VariantHelperRelease($pSelf)
Local $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;", $pSelf - 8)
$tObject.RefCount -= 1
If $tObject.RefCount > 0 Then Return $tObject.RefCount
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$tObject.Variant)
DllCall("kernel32.dll", "ptr", "GlobalFree", "handle", DllCall("kernel32.dll", "ptr", "GlobalHandle", "ptr", $tObject.Variant)[0])
DllCall("kernel32.dll", "ptr", "GlobalFree", "handle", DllCall("kernel32.dll", "ptr", "GlobalHandle", "ptr", DllStructGetPtr($tObject))[0])
Return 0
EndFunc
Func ___Class__Example_VariantHelperGetTypeInfoCount($pSelf, $pctinfo)
DllStructSetData(DllStructCreate("UINT",$pctinfo),1, 0)
Return 0
EndFunc
Func ___Class__Example_VariantHelperGetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
If $iTInfo<>0 Then Return -2147352565
If $ppTInfo=0 Then Return -2147024809
Return 0
EndFunc
Func ___Class__Example_VariantHelperGetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
Local $tIds = DllStructCreate("long i", $rgDispId)
DllStructSetData($tIds, 1, 1)
Return 0
EndFunc
Func ___Class__Example_VariantHelperInvoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
Local $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;", $pSelf - 8)
If BitAND($wFlags, 2) = 2 Then
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "ptr", $tObject.Variant)
Return 0
EndIf
If BitAND($wFlags, 4) = 4 Then
$tParams = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
If $tParams.cArgs <> 1 Then Return -2147352562
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$tObject.Variant)
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$tObject.Variant, "ptr", $tParams.rgvargs)
Return 0
EndIf
Return -2147352567
EndFunc
Func ___Class__Example_ToVariant($vValue)
Local $oObject = ___Class__Example_VariantHelper()
$oObject.a = $vValue
Local $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;", ptr($oObject) - 8)
Local $tVariant = DllStructCreate("ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2")
DllCall("OleAut32.dll","LONG","VariantClear","struct*",$tVariant)
DllCall("OleAut32.dll","LONG","VariantCopy","struct*",$tVariant, "ptr", $tObject.Variant)
Return $tVariant
EndFunc
Func ___Class__Example_FromVariant($pVariant)
Local $oObject = ___Class__Example_VariantHelper()
Local $tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];ptr Variant;", ptr($oObject) - 8)
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$tObject.Variant)
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$tObject.Variant, "struct*", $pVariant)
Return $oObject.a
EndFunc
Func Example($ts = 'now')
Local Static $QueryInterface = DllCallbackRegister(__Object__Class_Example_QueryInterface, 'LONG', 'ptr;ptr;ptr'), $AddRef = DllCallbackRegister(__Object__Class_Example_AddRef, 'dword', 'PTR'), $Release = DllCallbackRegister(__Object__Class_Example_Release, 'dword', 'PTR'), $GetTypeInfoCount = DllCallbackRegister(__Object__Class_Example_GetTypeInfoCount, 'long', 'ptr;ptr'), $GetTypeInfo = DllCallbackRegister(__Object__Class_Example_GetTypeInfo, 'long', 'ptr;uint;int;ptr'), $GetIDsOfNames = DllCallbackRegister(__Object__Class_Example_GetIDsOfNames, 'long', 'ptr;ptr;ptr;uint;int;ptr'), $Invoke = DllCallbackRegister(__Object__Class_Example_Invoke, 'long', 'ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr')
$tObject = DllStructCreate('int RefCount;int Size;ptr Object;ptr Methods[7];ptr Properties[2];')
DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($QueryInterface), 1)
DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($AddRef), 2)
DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($Release), 3)
DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($GetTypeInfoCount), 4)
DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($GetTypeInfo), 5)
DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($GetIDsOfNames), 6)
DllStructSetData($tObject, 'Methods', DllCallbackGetPtr($Invoke), 7)
DllStructSetData($tObject, 'RefCount', 1)
DllStructSetData($tObject, 'Size', 7)
$pObject = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", DllCall("kernel32.dll", "handle", "GlobalAlloc", "uint", 0x0002, "ulong_ptr", DllStructGetSize($tObject))[0])[0]
DllCall("kernel32.dll", "none", "RtlMoveMemory", "struct*", $pObject, "struct*", $tObject, "ulong_ptr", DllStructGetSize($tObject))
$tObject = DllStructCreate('int RefCount;int Size;ptr Object;ptr Methods[7];ptr Properties[2];', $pObject)
DllStructSetData($tObject, "Object", DllStructGetPtr($tObject, "Methods"))
Local Static $tVariant = DllStructCreate("ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2")
DllStructSetData($tVariant, 'vt', 1)
$pObject = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", DllCall("kernel32.dll", "handle", "GlobalAlloc", "uint", 0x0002, "ulong_ptr", DllStructGetSize($tVariant))[0])[0]
DllCall("kernel32.dll", "none", "RtlMoveMemory", "struct*", $pObject, "struct*", ___Class__Example_ToVariant( Null), "ulong_ptr", DllStructGetSize($tVariant))
DllStructSetData($tObject, "Properties", $pObject, 1)
$pObject = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", DllCall("kernel32.dll", "handle", "GlobalAlloc", "uint", 0x0002, "ulong_ptr", DllStructGetSize($tVariant))[0])[0]
DllCall("kernel32.dll", "none", "RtlMoveMemory", "struct*", $pObject, "struct*", ___Class__Example_ToVariant(Null), "ulong_ptr", DllStructGetSize($tVariant))
DllStructSetData($tObject, "Properties", $pObject, 2)
Local $oObject = ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True)
__Class_Example___construct($oObject,$ts)
If @error <> 0 Then Return SetError(@error, @extended, 0)
Return $oObject
EndFunc

Func __Object__Class_Example_QueryInterface($pSelf, $pRIID, $pObj)
___Class__Example_VariantHelperQueryInterface($pSelf, $pRIID, $pObj)
EndFunc
Func __Object__Class_Example_AddRef($pSelf)
Return ___Class__Example_VariantHelperAddRef($pSelf)
EndFunc
Func __Object__Class_Example_Release($pSelf)
$tObject = DllStructCreate('int RefCount;int Size;ptr Object;ptr Methods[7];ptr Properties[2];', $pSelf - 8)
$tObject.RefCount -= 1
If $tObject.RefCount > 0 Then Return $tObject.RefCount
__Object__Class_Example_AddRef($pSelf)
$tObject.RefCount += 1
__Class_Example___destruct(ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True))
$tObject.RefCount -= 1
$pProperty = DllStructGetData($tObject, "Properties", 1)
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pProperty)
DllCall("kernel32.dll", "ptr", "GlobalFree", "handle", DllCall("kernel32.dll", "ptr", "GlobalHandle", "ptr", $pProperty)[0])
$pProperty = DllStructGetData($tObject, "Properties", 2)
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pProperty)
DllCall("kernel32.dll", "ptr", "GlobalFree", "handle", DllCall("kernel32.dll", "ptr", "GlobalHandle", "ptr", $pProperty)[0])
DllCall("kernel32.dll", "ptr", "GlobalFree", "handle", DllCall("kernel32.dll", "ptr", "GlobalHandle", "ptr", DllStructGetPtr($tObject))[0])
Return 0
EndFunc
Func __Object__Class_Example_GetTypeInfoCount($pSelf, $pctinfo)
Return ___Class__Example_VariantHelperGetTypeInfoCount($pSelf, $pctinfo)
EndFunc
Func __Object__Class_Example_GetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
Return ___Class__Example_VariantHelperGetTypeInfo($pSelf, $iTInfo, $lcid, $ppTInfo)
EndFunc
Func __Object__Class_Example_GetIDsOfNames($pSelf, $riid, $rgszNames, $cNames, $lcid, $rgDispId)
Local $tId = DllStructCreate("long i", $rgDispId)
Local $pStr = DllStructGetData(DllStructCreate("ptr", $rgszNames), 1)
Local $s_rgszName = DllStructGetData(DllStructCreate("WCHAR[255]", $pStr), 1)
Switch $s_rgszName
Case "property"
DllStructSetData($tId, 1, 1)
Case "dynamic"
DllStructSetData($tId, 1, 2)
Case "method"
DllStructSetData($tId, 1, 3)
Case Else
DllStructSetData($tId, 1, -1)
Return -2147352570
EndSwitch
Return 0
EndFunc
Func __Object__Class_Example_Invoke($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
If $dispIdMember=-1 Then Return -2147352573
$tObject = DllStructCreate('int RefCount;int Size;ptr Object;ptr Methods[7];ptr Properties[2];', $pSelf - 8)
Local Static $iVariant = DllStructGetSize(DllStructCreate("ushort vt;ushort r1;ushort r2;ushort r3;PTR data;PTR data2"))
Local Static $Invoke = DllCallbackRegister(__Object__Class_Example_InvokeAccessor, 'long', 'ptr;int;ptr;int;ushort;ptr;ptr;ptr;ptr')
Switch $dispIdMember
Case 1
If BitAND($wFlags, 2)=2 Then
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "ptr", DllStructGetData($tObject, "Properties", 1))
Return 0
EndIf
Local $tParams = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
If $tParams.cArgs <> 1 Then Return -2147352562
DllCall("OleAut32.dll","LONG","VariantClear","ptr",DllStructGetData($tObject, "Properties", 1))
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",DllStructGetData($tObject, "Properties", 1), "ptr", $tParams.rgvargs)
Return 0
Case 2
If BitAND($wFlags, 2)=2 Then
Local $_tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];int PropertyIndex;ptr OriginalObject;")
$_tObject.RefCount = 2; > 1 to prevent release to be triggered
$_tObject.Size = 7
$_tObject.Object = DllStructGetPtr($_tObject, "Methods")
$_tObject.PropertyIndex = 2
$_tObject.OriginalObject = $pSelf
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 1), 1)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 2), 2)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 3), 3)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 4), 4)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 5), 5)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 6), 6)
DllStructSetData($_tObject, "Methods", DllCallbackGetPtr($Invoke), 7)
Local $vValue = __Getter__Class_Example_dynamic(ObjCreateInterface(DllStructGetPtr($_tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True))
If @error <> 0 Then Return -2147352567
$tVariant = ___Class__Example_ToVariant($vValue)
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "struct*", $tVariant)
Return 0
EndIf
Local $tParams = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
If $tParams.cArgs <> 1 Then Return -2147352562
Local $_tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];int PropertyIndex;ptr OriginalObject;")
$_tObject.RefCount = 2; > 1 to prevent release to be triggered
$_tObject.Size = 7
$_tObject.Object = DllStructGetPtr($_tObject, "Methods")
$_tObject.PropertyIndex = 2
$_tObject.OriginalObject = $pSelf
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 1), 1)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 2), 2)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 3), 3)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 4), 4)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 5), 5)
DllStructSetData($_tObject, "Methods", DllStructGetData($tObject, "Methods", 6), 6)
DllStructSetData($_tObject, "Methods", DllCallbackGetPtr($Invoke), 7)
__Setter__Class_Example_dynamic(ObjCreateInterface(DllStructGetPtr($_tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True), ___Class__Example_FromVariant($tParams.rgvargs))
If @error <> 0 Then Return -2147352567
Return 0
Case 3
If BitAND($wFlags, 4) = 4 Or BitAND($wFlags, 8) = 8 Then Return -2147352567
$tDISPPARAMS = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
If $tDISPPARAMS.cArgs < 0 Or $tDISPPARAMS.cArgs > 0 Then Return -2147352562
__Object__Class_Example_AddRef($pSelf)
Local $parameters[$tDISPPARAMS.cArgs + 2] = ["CallArgArray", ObjCreateInterface(DllStructGetPtr($tObject, "Object"), "{00020400-0000-0000-C000-000000000046}", Default, True)]
Local $j = 2
For $i=$tDISPPARAMS.cArgs-1 To 0 Step -1
$parameters[$j] = ___Class__Example_FromVariant($tDISPPARAMS.rgvargs+$iVariant*$i)
Next
Local $vValue = Call(__Class_Example_method, $parameters)
If @error <> 0 Then Return -2147352567
$tVariant = ___Class__Example_ToVariant($vValue)
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "struct*", $tVariant)
Return 0
EndSwitch
EndFunc
Func __Object__Class_Example_InvokeAccessor($pSelf, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
Local $_tObject = DllStructCreate("int RefCount;int Size;ptr Object;ptr Methods[7];int PropertyIndex;ptr OriginalObject;", $pSelf - 8)
If $dispIdMember = $_tObject.PropertyIndex Then
$tObject = DllStructCreate('int RefCount;int Size;ptr Object;ptr Methods[7];ptr Properties[2];', $_tObject.OriginalObject - 8)
If BitAND($wFlags, 2)=2 Then
DllCall("OleAut32.dll","LONG","VariantClear","ptr",$pVarResult)
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",$pVarResult, "ptr", DllStructGetData($tObject, "Properties", $dispIdMember))
Return 0
EndIf
Local $tParams = DllStructCreate("ptr rgvargs;ptr rgdispidNamedArgs;dword cArgs;dword cNamedArgs;", $pDispParams)
If $tParams.cArgs <> 1 Then Return -2147352562
DllCall("OleAut32.dll","LONG","VariantClear","ptr",DllStructGetData($tObject, "Properties", $dispIdMember))
DllCall("OleAut32.dll","LONG","VariantCopy","ptr",DllStructGetData($tObject, "Properties", $dispIdMember), "ptr", $tParams.rgvargs)
Return 0
EndIf
Return __Object__Class_Example_Invoke($_tObject.OriginalObject, $dispIdMember, $riid, $lcid, $wFlags, $pDispParams, $pVarResult, $pExcepInfo, $puArgErr)
EndFunc
Func __Class_Example___construct($this, $ts = 'now')
	$this.property = 0
EndFunc

Func __Class_Example___destruct($this)
	; Destructor code here.
EndFunc

Func __Getter__Class_Example_dynamic($this)
Return $this.dynamic & $this.property
EndFunc

Func __Setter__Class_Example_dynamic($this,$value)
$this.property += 1
        $this.dynamic = $value
EndFunc

Func __Class_Example_method($this)
	Return "something"
EndFunc



$oExample = Example()

$oExample.dynamic = 12

MsgBox(0, "", $oExample.dynamic)

MsgBox(0, "", $oExample.method())

$oExample = Null
