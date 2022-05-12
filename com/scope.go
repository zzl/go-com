package com

import (
	"github.com/zzl/go-win32api/win32"
	"unsafe"
)

var scopedComPtrs []unsafe.Pointer

var CurrentScope *Scope

type Scope struct {
	index       int
	ParentScope *Scope
}

func NewScope() *Scope {
	scope := &Scope{
		index:       len(scopedComPtrs),
		ParentScope: CurrentScope,
	}
	CurrentScope = scope
	return scope
}

func (this *Scope) Add(pUnknown unsafe.Pointer) {
	scopedComPtrs = append(scopedComPtrs, pUnknown)
}

func AddScopedComPtr(pUnknown *win32.IUnknown) {
	if CurrentScope == nil {
		panic("no current scope") //?
	}
	CurrentScope.Add(unsafe.Pointer(pUnknown))
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
	count := len(scopedComPtrs)
	for n := this.index; n < count; n++ {
		pointer := scopedComPtrs[n]
		if pointer != nil {
			(*win32.IUnknown)(pointer).Release()
		}
	}
	scopedComPtrs = scopedComPtrs[:this.index]
}
