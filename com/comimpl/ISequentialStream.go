package comimpl

import (
	"syscall"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
)

//
type ISequentialStreamImpl struct {
	com.IUnknownImpl
}

func (this *ISequentialStreamImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_ISequentialStream {
		//*(**com.IUnknownComObj)(ppvObject) = this.ComObject
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IUnknownImpl.QueryInterface(riid, ppvObject)
}

func (this *ISequentialStreamImpl) Read(pv unsafe.Pointer, cb uint32, pcbRead *uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *ISequentialStreamImpl) Write(pv unsafe.Pointer, cb uint32, pcbWritten *uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

//
type ISequentialStreamComObj struct {
	com.IUnknownComObj
}

func (this *ISequentialStreamComObj) impl() win32.ISequentialStreamInterface {
	return this.Impl().(win32.ISequentialStreamInterface)
}

func (this *ISequentialStreamComObj) Read(pv unsafe.Pointer, cb uint32, pcbRead *uint32) uintptr {
	return (uintptr)(this.impl().Read(pv, cb, pcbRead))
}

func (this *ISequentialStreamComObj) Write(pv unsafe.Pointer, cb uint32, pcbWritten *uint32) uintptr {
	return (uintptr)(this.impl().Write(pv, cb, pcbWritten))
}

var _pISequentialStreamVtbl *win32.ISequentialStreamVtbl

func (this *ISequentialStreamComObj) BuildVtbl(lock bool) *win32.ISequentialStreamVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pISequentialStreamVtbl != nil {
		return _pISequentialStreamVtbl
	}
	_pISequentialStreamVtbl = (*win32.ISequentialStreamVtbl)(
		com.Malloc(unsafe.Sizeof(*_pISequentialStreamVtbl)))

	*_pISequentialStreamVtbl = win32.ISequentialStreamVtbl{
		IUnknownVtbl: *this.IUnknownComObj.BuildVtbl(false),
		Read:         syscall.NewCallback((*ISequentialStreamComObj).Read),
		Write:        syscall.NewCallback((*ISequentialStreamComObj).Write),
	}
	return _pISequentialStreamVtbl
}

func (this *ISequentialStreamComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}
