package ole

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"runtime"
)

func Initialize() {
	runtime.LockOSThread()
	win32.OleInitialize(nil)

	com.InitializeContext()
}

func Uninitialize() {
	runtime.GC()
	com.UninitializeContext()
	win32.OleUninitialize()
}
