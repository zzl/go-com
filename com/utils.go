package com

import (
	"fmt"
	"github.com/zzl/go-win32api/v2/win32"
	"log"
	"unsafe"
)

func HresultFromWin32(err win32.WIN32_ERROR) win32.HRESULT {
	hr := win32.HRESULT(err)
	if hr > 0 {
		hr = win32.HRESULT((uint32(err) & 0x0000FFFF) |
			(uint32(win32.FACILITY_WIN32) << 16) | 0x80000000)
	}
	return hr
}

func PrintHr(hr win32.HRESULT) {
	fmt.Println(win32.HRESULT_ToString(hr))
}

func Succeeded(hr win32.HRESULT, logFail ...bool) bool {
	return !Failed(hr, logFail...)
}

func Failed(hr win32.HRESULT, logFail ...bool) bool {
	if win32.FAILED(hr) {
		if len(logFail) == 1 && logFail[0] {
			log.Println(win32.HRESULT_ToString(hr))
		}
		return true
	}
	return false
}

type QueryInterfaceResultType interface {
	IIDProvider
	win32.IUnknownObject
}

func QueryInterface[T QueryInterfaceResultType](pObj win32.IUnknownObject, scoped bool) T {
	var p T
	pObj.GetIUnknown().QueryInterface(p.IID(), unsafe.Pointer(&p))
	if scoped {
		AddToScope(p)
	}
	return p
}

func MessageLoop() {
	var msg win32.MSG
	for {
		ret, _ := win32.GetMessage(&msg, 0, 0, 0)
		if ret == 0 {
			break
		}
		win32.DispatchMessage(&msg)
	}
}