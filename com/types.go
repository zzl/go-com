package com

import (
	"unsafe"

	"github.com/zzl/go-win32api/win32"
)

type Disposable interface {
	Dispose()
}

type UnknownClass struct {
	win32.IUnknown
}

func (this *UnknownClass) Dispose() {
	//
}

func NewUnknownClass(pUnk *win32.IUnknown, scoped bool) *UnknownClass {
	p := (*UnknownClass)(unsafe.Pointer(pUnk))
	if scoped {
		AddScopedComPtr(&p.IUnknown)
	}
	return p
}
