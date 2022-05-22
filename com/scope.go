package com

import (
	"github.com/zzl/go-win32api/win32"
	"unsafe"
	"log"
)

type scopedObject struct {
	Ptr  unsafe.Pointer
	Type int //0:com interface, 1:bstr, 2:*variant, 3:safearray
}

var CurrentScope *Scope

type Scope struct {
	scopedObjs []scopedObject
	ParentScope *Scope
}

func NewScope() *Scope {
	scope := &Scope{
		ParentScope: CurrentScope,
	}
	CurrentScope = scope
	return scope
}

func (this *Scope) Add(pUnknown unsafe.Pointer) {
	this.scopedObjs = append(this.scopedObjs, scopedObject{Ptr: pUnknown})
}

func (this *Scope) AddComPtr(iunknownObj win32.IUnknownObject, addRef ...bool) {
	pUnknown := iunknownObj.GetIUnknown()
	this.scopedObjs = append(this.scopedObjs, scopedObject{Ptr: unsafe.Pointer(pUnknown), Type: 0})
	if len(addRef) == 1 && addRef[0] {
		pUnknown.AddRef()
	}
}

func (this *Scope) AddBstr(bstr win32.BSTR) {
	this.scopedObjs = append(this.scopedObjs, scopedObject{Ptr: unsafe.Pointer(bstr), Type: 1})
}

func (this *Scope) AddVar(pVar *win32.VARIANT) {
	this.scopedObjs = append(this.scopedObjs, scopedObject{Ptr: unsafe.Pointer(pVar), Type: 2})
}

func (this *Scope) AddVarIfNeeded(pVar *win32.VARIANT) {
	switch win32.VARENUM(pVar.Vt) {
	case win32.VT_UNKNOWN, win32.VT_DISPATCH, win32.VT_BSTR, win32.VT_SAFEARRAY:
		break
	default:
		return
	}
	this.scopedObjs = append(this.scopedObjs, scopedObject{Ptr: unsafe.Pointer(pVar), Type: 2})
}

func (this *Scope) AddArray(psa *win32.SAFEARRAY) {
	this.scopedObjs = append(this.scopedObjs, scopedObject{Ptr: unsafe.Pointer(psa), Type: 3})
}

func AddToScope(value interface{}, scope ...*Scope) {
	var s *Scope
	if len(scope) != 0 {
		s = scope[0]
	} else {
		if CurrentScope == nil {
			log.Panic("no current scope")
		}
		s = CurrentScope
	}
	switch v := value.(type) {
	case win32.IUnknownObject:
		s.Add(unsafe.Pointer(v.GetIUnknown()))
	case win32.BSTR:
		s.AddBstr(v)
	case win32.VARIANT:
		s.AddVarIfNeeded(&v)
	case *win32.SAFEARRAY:
		s.AddArray(v)
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
	count := len(this.scopedObjs)
	for n := 0; n < count; n++ {
		obj := this.scopedObjs[n]
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
	this.scopedObjs = nil
}
