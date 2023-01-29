package com

import (
	"syscall"
	"unsafe"

	"github.com/zzl/go-win32api/v2/win32"
)

type BStr struct {
	bs win32.BSTR
}

func NewBStr(bs win32.BSTR) *BStr {
	return &BStr{bs}
}

func NewBStringFromStr(str string) *BStr {
	pwsz, _ := syscall.UTF16PtrFromString(str)
	bs := win32.SysAllocString(pwsz)
	return NewBStr(bs)
}

func (this *BStr) String() string {
	len := win32.SysStringLen(this.bs)
	if len == 0 {
		return ""
	}
	ws := unsafe.Slice(this.bs, len)
	s := syscall.UTF16ToString(ws)
	return s
}

func (this *BStr) ToStringAndFree() string {
	str := this.String()
	this.Free()
	return str
}

func (this *BStr) BSTR() win32.BSTR {
	return this.bs
}

func (this *BStr) PBSTR() *win32.BSTR {
	return &this.bs
}

func (this *BStr) Addr() uintptr {
	return uintptr(unsafe.Pointer(this.bs))
}

func (this *BStr) Free() {
	if this.bs != nil {
		win32.SysFreeString(this.bs)
		this.bs = nil
	}
}

type Bstrs struct {
	bss []*BStr
}

func NewBStrs() *Bstrs {
	return &Bstrs{}
}

func (this *Bstrs) Dispose() {
	for _, bs := range this.bss {
		bs.Free()
	}
	this.bss = nil
}

func (this *Bstrs) Add(str string) *BStr {
	bs := NewBStringFromStr(str)
	this.bss = append(this.bss, bs)
	return bs
}
