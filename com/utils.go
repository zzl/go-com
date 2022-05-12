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
