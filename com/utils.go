package com

import (
	"runtime"
	"sync"

	"github.com/zzl/go-win32api/win32"
)

var comThreadIds sync.Map

func Initialize() {
	runtime.LockOSThread()
	win32.CoInitialize(nil)
	tId := win32.GetCurrentThreadId()
	comThreadIds.Store(tId, true)
}

func EnsureThreadCoInitialized() {
	tId := win32.GetCurrentThreadId()
	if _, loaded := comThreadIds.LoadOrStore(tId, true); !loaded {
		win32.CoInitialize(nil)
	}
}

func HresultFromWin32(err win32.WIN32_ERROR) win32.HRESULT {
	hr := win32.HRESULT(err)
	if hr > 0 {
		hr = win32.HRESULT((uint32(err) & 0x0000FFFF) |
			(uint32(win32.FACILITY_WIN32) << 16) | 0x80000000)
	}
	return hr
}
