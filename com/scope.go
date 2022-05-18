package com

import (
	"github.com/zzl/go-win32api/win32"
	"unsafe"
)

type scopedObject struct {
	Ptr  unsafe.Pointer
	Type int //0:com interface, 1:bstr, 2:*variant, 3:safearray
}

var scopedObjs []scopedObject

var CurrentScope *Scope

type Scope struct {
	index       int
	ParentScope *Scope
}

func NewScope() *Scope {
	scope := &Scope{
		index:       len(scopedObjs),
		ParentScope: CurrentScope,
	}
	CurrentScope = scope
	return scope
}

func (this *Scope) Add(pUnknown unsafe.Pointer) {
	scopedObjs = append(scopedObjs, scopedObject{Ptr: pUnknown})
}

func (this *Scope) AddBstr(bstr win32.BSTR) {
	scopedObjs = append(scopedObjs, scopedObject{Ptr: unsafe.Pointer(bstr), Type: 1})
}

func (this *Scope) AddVar(pVar *win32.VARIANT) {
	scopedObjs = append(scopedObjs, scopedObject{Ptr: unsafe.Pointer(pVar), Type: 2})
}

func (this *Scope) AddVarIfNeeded(pVar *win32.VARIANT) {
	switch win32.VARENUM(pVar.Vt) {
	case win32.VT_UNKNOWN, win32.VT_DISPATCH, win32.VT_BSTR, win32.VT_SAFEARRAY:
		break
	default:
		return
	}
	scopedObjs = append(scopedObjs, scopedObject{Ptr: unsafe.Pointer(pVar), Type: 2})
}

func (this *Scope) AddArray(psa *win32.SAFEARRAY) {
	scopedObjs = append(scopedObjs, scopedObject{Ptr: unsafe.Pointer(psa), Type: 3})
}

func AddToScope(value interface{}) {
	if CurrentScope == nil {
		panic("no current scope") //?
	}
	switch v := value.(type) {
	case win32.IUnknownObject:
		CurrentScope.Add(unsafe.Pointer(v.GetIUnknown()))
	case win32.BSTR:
		CurrentScope.AddBstr(v)
	case win32.VARIANT:
		CurrentScope.AddVarIfNeeded(&v)
	case *win32.SAFEARRAY:
		CurrentScope.AddArray(v)
	default:
		println("?")
	}
}

func WithScope(action func()) {
	scope := NewScope()
	defer scope.Leave()
	action()
}

func (this *Scope) Leave() {
	this.Clear()
	CurrentScope = this.ParentScope
}

func (this *Scope) Clear() {
	count := len(scopedObjs)
	for n := this.index; n < count; n++ {
		obj := scopedObjs[n]
		if obj.Ptr != nil {
			if obj.Type == 0 {
				(*win32.IUnknown)(obj.Ptr).Release()
			} else if obj.Type == 1 {
				win32.SysFreeString((win32.BSTR)(obj.Ptr))
			} else if obj.Type == 2 {
				win32.VariantClear((*win32.VARIANT)(obj.Ptr))
			} else if obj.Type == 3 {
				win32.SafeArrayDestroy((*win32.SAFEARRAY)(obj.Ptr))
			}
		}
	}
	scopedObjs = scopedObjs[:this.index]
}
