package comimpl

import (
	"syscall"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
)

type IDropSourceImpl struct {
	com.IUnknownImpl
}

func (this *IDropSourceImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IDropSource {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IUnknownImpl.QueryInterface(riid, ppvObject)
}

func (this *IDropSourceImpl) QueryContinueDrag(fEscapePressed win32.BOOL, grfKeyState uint32) win32.HRESULT {
	return win32.S_OK
}

func (this *IDropSourceImpl) GiveFeedback(dwEffect uint32) win32.HRESULT {
	return win32.S_OK
}

//
type IDropSourceComObj struct {
	com.IUnknownComObj
}

func (this *IDropSourceComObj) impl() win32.IDropSourceInterface {
	return this.Impl().(win32.IDropSourceInterface)
}

func (this *IDropSourceComObj) QueryContinueDrag(fEscapePressed win32.BOOL, grfKeyState win32.MODIFIERKEYS_FLAGS) uintptr {
	return uintptr(this.impl().QueryContinueDrag(fEscapePressed, grfKeyState))
}

func (this *IDropSourceComObj) GiveFeedback(dwEffect win32.DROPEFFECT) uintptr {
	return uintptr(this.impl().GiveFeedback(dwEffect))
}

var _pIDropSourceVtbl *win32.IDropSourceVtbl

func (this *IDropSourceComObj) BuildVtbl(lock bool) *win32.IDropSourceVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIDropSourceVtbl != nil {
		return _pIDropSourceVtbl
	}
	_pIDropSourceVtbl = &win32.IDropSourceVtbl{
		IUnknownVtbl:      *this.IUnknownComObj.BuildVtbl(false),
		QueryContinueDrag: syscall.NewCallback((*IDropSourceComObj).QueryContinueDrag),
		GiveFeedback:      syscall.NewCallback((*IDropSourceComObj).GiveFeedback),
	}
	return _pIDropSourceVtbl
}

func (this *IDropSourceComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

func (this *IDropSourceComObj) IDropSource() *win32.IDropSource {
	return (*win32.IDropSource)(unsafe.Pointer(this))
}

func NewIDropSourceComObj(impl win32.IDropSourceInterface) *IDropSourceComObj {
	comObj := com.NewComObj[IDropSourceComObj](impl)
	return comObj
}
