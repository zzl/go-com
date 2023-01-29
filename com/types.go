package com

import (
	"syscall"
	"unsafe"

	"github.com/zzl/go-win32api/v2/win32"
)

type Disposable interface {
	Dispose()
}

type IIDProvider interface {
	IID() *syscall.GUID
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
		AddToScope(p)
	}
	return p
}
