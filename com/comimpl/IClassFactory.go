package comimpl

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

type IClassFactoryImpl struct {
	com.IUnknownImpl
}

func (this *IClassFactoryImpl) CreateInstance(pUnkOuter *win32.IUnknown, riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IClassFactoryImpl) LockServer(fLock win32.BOOL) win32.HRESULT {
	return win32.E_NOTIMPL
}

//
type IClassFactoryComObj struct {
	com.IUnknownComObj
}

func (this *IClassFactoryComObj) IID() *syscall.GUID {
	return &win32.IID_IClassFactory
}

func (this *IClassFactoryComObj) impl() win32.IClassFactoryInterface {
	return this.Impl().(win32.IClassFactoryInterface)
}

func (this *IClassFactoryComObj) CreateInstance(pUnkOuter *win32.IUnknown, riid *syscall.GUID, ppvObject unsafe.Pointer) uintptr {
	return (uintptr)(this.impl().CreateInstance(pUnkOuter, riid, ppvObject))
}

func (this *IClassFactoryComObj) LockServer(fLock int32) uintptr {
	return (uintptr)(this.impl().LockServer(fLock))
}

var _pIClassFactoryVtbl *win32.IClassFactoryVtbl

func (this *IClassFactoryComObj) BuildVtbl(lock bool) *win32.IClassFactoryVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIClassFactoryVtbl != nil {
		return _pIClassFactoryVtbl
	}
	_pIClassFactoryVtbl = (*win32.IClassFactoryVtbl)(
		com.Malloc(unsafe.Sizeof(*_pIClassFactoryVtbl)))
	*_pIClassFactoryVtbl = win32.IClassFactoryVtbl{
		IUnknownVtbl:   *this.IUnknownComObj.BuildVtbl(false),
		CreateInstance: syscall.NewCallback((*IClassFactoryComObj).CreateInstance),
		LockServer:     syscall.NewCallback((*IClassFactoryComObj).LockServer),
	}
	return _pIClassFactoryVtbl
}

func (this *IClassFactoryComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}
