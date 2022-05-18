package ole

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
	"log"
	"unsafe"
)

type SafeArrayInterface interface {
	SafeArrayPtr() *win32.SAFEARRAY
}

type Array[T any] struct {
	Psa *win32.SAFEARRAY
}

func NewArray[T any](count int32, scoped bool) Array[T] {
	return NewArrayWithBase[T](0, count, scoped)
}

func NewArrayWithBase[T any](base int32, count int32, scoped bool) Array[T] {
	return NewArrayWithBounds[T]([]win32.SAFEARRAYBOUND{
		{LLbound: base, CElements: uint32(count)}}, scoped)
}

func NewArray2D[T any](count1 int32, count2 int32, scoped bool) Array[T] {
	return NewArray2DWithBase[T](0, count1, 0, count2, scoped)
}

func NewArray2DWithBase[T any](base1, count1, base2, count2 int32, scoped bool) Array[T] {
	return NewArrayWithBounds[T]([]win32.SAFEARRAYBOUND{
		{LLbound: base1, CElements: uint32(count1)},
		{LLbound: base2, CElements: uint32(count2)}}, scoped)
}

func NewArrayWithBounds[T any](bounds []win32.SAFEARRAYBOUND, scoped bool) Array[T] {
	var t T
	vt := CheckVarType(t)
	return NewArrayWithVt[T](vt, bounds, scoped)
}

func NewArrayWithVt[T any](varType win32.VARENUM, bounds []win32.SAFEARRAYBOUND, scoped bool) Array[T] {
	Psa := win32.SafeArrayCreate(uint16(varType), uint32(len(bounds)), &bounds[0])
	a := Array[T]{Psa}
	a.lock()
	if scoped {
		com.CurrentScope.AddArray(Psa)
	}
	return a
}

func NewArrayByAttach[T any](Psa *win32.SAFEARRAY) Array[T] {
	a := Array[T]{Psa}
	a.lock()
	return a
}

func (me Array[T]) SafeArrayPtr() *win32.SAFEARRAY {
	return me.Psa
}

func (me Array[T]) Detach() *win32.SAFEARRAY {
	me.unlock()
	return me.Psa
}

func (me Array[T]) Destroy() {
	me.unlock()
	win32.SafeArrayDestroy(me.Psa)
}

func (me Array[T]) GetLowerBound(dim int) int {
	var lBound int32
	win32.SafeArrayGetLBound(me.Psa, uint32(dim+1), &lBound)
	return int(lBound)
}

func (me Array[T]) GetUpperBound(dim int) int {
	var uBound int32
	win32.SafeArrayGetUBound(me.Psa, uint32(dim+1), &uBound)
	return int(uBound)
}

func (me Array[T]) GetCount(dim int) int {
	return me.GetUpperBound(dim) - me.GetLowerBound(dim) + 1
}

func (me Array[T]) GetDimCount() int {
	return int(win32.SafeArrayGetDim(me.Psa))
}

func (me Array[T]) GetVarType() win32.VARENUM {
	var vt uint16
	win32.SafeArrayGetVartype(me.Psa, &vt)
	return win32.VARENUM(vt)
}

func (me Array[T]) GetShape() []int32 {
	cDims := me.Psa.CDims
	shape := make([]int32, cDims)
	bounds := unsafe.Slice(&me.Psa.Rgsabound[0], cDims)
	for n := uint16(0); n < cDims; n++ {
		bound := bounds[n]
		shape[cDims-n-1] = int32(bound.CElements)
	}
	return shape
}

func (me Array[T]) GetAt(index ...int32) (ret T) {
	if len(index) == 1 {
		index0 := index[0] - me.Psa.Rgsabound[0].LLbound
		ret = *(*T)(unsafe.Pointer(uintptr(me.Psa.PvData) +
			unsafe.Sizeof(ret)*uintptr(index0)))
	} else {
		win32.SafeArrayGetElement(me.Psa, &index[0], unsafe.Pointer(&ret))
	}
	return
}

func (me Array[T]) SetAt(index int, value T) {
	me.SetAtEx(index, value, true)
}

func (me Array[T]) SetAtEx(index int, value T, copy bool) {
	index0 := int32(index) - me.Psa.Rgsabound[0].LLbound
	p := unsafe.Pointer(uintptr(me.Psa.PvData) + unsafe.Sizeof(value)*uintptr(index0))
	if *(*uintptr)(p) != 0 {
		switch v := any(value).(type) {
		case win32.VARIANT, Variant:
			if copy {
				win32.VariantCopyInd((*win32.VARIANT)(p),
					(*win32.VARIANT)(unsafe.Pointer(&value)))
				return //
			}
			(*Variant)(p).Clear()
		case *win32.IUnknown:
			(*(**win32.IUnknown)(p)).Release()
			if copy {
				v.AddRef()
			}
		case *win32.IDispatch:
			(*(**win32.IDispatch)(p)).Release()
			if copy {
				v.AddRef()
			}
		case win32.BSTR, com.BStr:
			(*com.BStr)(p).Free()
			if copy {
				*(*win32.BSTR)(p) = win32.SysAllocString(
					(win32.BSTR)(unsafe.Pointer(&value)))
				return //
			}
		}
	}
	*(*T)(p) = value
}

func (me Array[T]) SetAt2(rowIndex int, colIndex int, value T) {
	index := []int32{int32(rowIndex), int32(colIndex)}
	win32.SafeArrayPutElement(me.Psa, &index[0], unsafe.Pointer(&value))
}

func (me Array[T]) SetAtMd(index []int32, value T) {
	win32.SafeArrayPutElement(me.Psa, &index[0], unsafe.Pointer(&value))
	return
}

func (me Array[T]) lock() {
	win32.SafeArrayLock(me.Psa)
}

func (me Array[T]) unlock() {
	win32.SafeArrayUnlock(me.Psa)
}

func (me Array[T]) Copy() Array[T] {
	var Psa2 *win32.SAFEARRAY
	hr := win32.SafeArrayCopy(me.Psa, &Psa2)
	if win32.FAILED(hr) {
		log.Panic(win32.HRESULT_ToString(hr))
	}
	return NewArrayByAttach[T](Psa2)
}

func (me Array[T]) ToVar() Variant {
	return me.ToVarEx(false, false)
}

func (me Array[T]) ToVarEx(copy bool, scoped bool) Variant {
	var v Variant
	v.Vt = uint16(win32.VT_ARRAY | me.GetVarType())
	if copy {
		*v.Parray() = me.Copy().Detach()
	} else {
		*v.Parray() = me.Psa
	}
	if scoped {
		com.AddToScope(v)
	}
	return v
}
