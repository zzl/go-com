package com

import (
	"github.com/zzl/go-win32api/v2/win32"
	"log"
	"sync"
	"unsafe"
)

type scopedObject struct {
	Ptr  unsafe.Pointer
	Type int //0:com interface, 1:bstr, 2:*variant, 3:safearray
}

var muScope sync.Mutex

type Scope struct {
	scopedObjs  []scopedObject
	ParentScope *Scope
}

func NewScope() *Scope {
	context := GetContext()
	scope := &Scope{
		ParentScope: context.GetCurrentScope(),
	}
	context.SetCurrentScope(scope)
	return scope
}

func (this *Scope) _add(so scopedObject) {
	muScope.Lock()
	this.scopedObjs = append(this.scopedObjs, so)
	muScope.Unlock()
}

func (this *Scope) Add(pUnknown unsafe.Pointer) {
	this._add(scopedObject{Ptr: pUnknown})
}

func (this *Scope) AddComPtr(iunknownObj win32.IUnknownObject, addRef ...bool) {
	pUnknown := iunknownObj.GetIUnknown()
	if len(addRef) == 1 && addRef[0] {
		pUnknown.AddRef()
	}
	this._add(scopedObject{Ptr: unsafe.Pointer(pUnknown), Type: 0})
}

func (this *Scope) AddBstr(bstr win32.BSTR) {
	this._add(scopedObject{Ptr: unsafe.Pointer(bstr), Type: 1})
}

func (this *Scope) AddVar(pVar *win32.VARIANT) {
	this._add(scopedObject{Ptr: unsafe.Pointer(pVar), Type: 2})
}

func (this *Scope) AddVarIfNeeded(pVar *win32.VARIANT) {
	switch win32.VARENUM(pVar.Vt) {
	case win32.VT_UNKNOWN, win32.VT_DISPATCH, win32.VT_BSTR, win32.VT_SAFEARRAY:
		break
	default:
		return
	}
	this._add(scopedObject{Ptr: unsafe.Pointer(pVar), Type: 2})
}

func (this *Scope) AddArray(psa *win32.SAFEARRAY) {
	this._add(scopedObject{Ptr: unsafe.Pointer(psa), Type: 3})
}

type VariantCompatible interface {
	AsVARIANT() *win32.VARIANT
}

func AddToScope(value interface{}, scope ...*Scope) {
	var s *Scope
	if len(scope) != 0 {
		s = scope[0]
	} else {
		currentScope := GetContext().GetCurrentScope()
		if currentScope == nil {
			log.Panic("no current scope")
		}
		s = currentScope
	}
	switch v := value.(type) {
	case win32.IUnknownObject:
		s.Add(unsafe.Pointer(v.GetIUnknown()))
	case win32.BSTR:
		s.AddBstr(v)
	case BStr:
		s.AddBstr(v.bs)
	case *win32.VARIANT:
		s.AddVarIfNeeded(v)
	case win32.VARIANT:
		s.AddVarIfNeeded(&v)
	case VariantCompatible:
		s.AddVarIfNeeded(v.AsVARIANT())
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
	GetContext().SetCurrentScope(this.ParentScope)
}

func (this *Scope) Clear() {
	muScope.Lock()
	count := len(this.scopedObjs)
	scopedObjs := make([]scopedObject, count)
	copy(scopedObjs, this.scopedObjs)
	this.scopedObjs = nil
	muScope.Unlock()

	for n := count - 1; n >= 0; n-- {
		obj := scopedObjs[n]
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
