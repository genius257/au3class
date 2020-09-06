#include "AutoItObject_Internal.au3"
#include-once

Func Example($ts = 'now')
	Local $this = IDispatch()
	$this.property = Null
	$this.__defineGetter("dynamic", __Getter__Class_Example_dynamic)
	$this.__defineSetter("dynamic", __Setter__Class_Example_dynamic)
	$this.__defineGetter("method", __Class_Example_method)
	$this.__destructor(__Class_Example___destruct)
	$this.__seal()
	__Class_Example___construct($this, $ts)
	Return $this
EndFunc

Func __Class_Example___construct($this, $ts = 'now')
	$this.property = 0
EndFunc

Func __Class_Example___destruct($this)
	$this = $this.parent
	; Destructor code here.
EndFunc

Func __Getter__Class_Example_dynamic($_oAccessorObject)
	Local $this = $_oAccessorObject.parent
	Return $_oAccessorObject.val & $this.property
EndFunc

Func __Setter__Class_Example_dynamic($_oAccessorObject)
	Local $this = $_oAccessorObject.parent
	Local $value = $_oAccessorObject.ret
	$this.property += 1
        $_oAccessorObject.val = $value
EndFunc

Func __Class_Example_method($this)
	$this = $this.parent
	Return "something"
EndFunc



$oExample = Example()

$oExample.dynamic = 12

MsgBox(0, "", $oExample.dynamic)

MsgBox(0, "", $oExample.method())
