package oleimpl

import (
	"github.com/zzl/go-win32api/win32"
	"syscall"
	"unsafe"
)

type DualWithEventsDispImpl struct {
	DualDispImpl
}

func (this *DualWithEventsDispImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IConnectionPointContainer {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IUnknownImpl.QueryInterface(riid, ppvObject)
}
