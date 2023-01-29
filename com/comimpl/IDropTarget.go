package comimpl

import (
	"syscall"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
)

type IDropTargetImpl struct {
	com.IUnknownImpl
}

func (this *IDropTargetImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IDropTarget {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IUnknownImpl.QueryInterface(riid, ppvObject)
}

func (this *IDropTargetImpl) DragEnter(pDataObj *win32.IDataObject, grfKeyState uint32, pt win32.POINTL, pdwEffect *uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IDropTargetImpl) DragOver(grfKeyState uint32, pt win32.POINTL, pdwEffect *uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IDropTargetImpl) DragLeave() win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IDropTargetImpl) Drop(pDataObj *win32.IDataObject, grfKeyState uint32, pt win32.POINTL, pdwEffect *uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

//
type IDropTargetComObj struct {
	com.IUnknownComObj
}

func (this *IDropTargetComObj) impl() win32.IDropTargetInterface {
	return this.Impl().(win32.IDropTargetInterface)
}

func (this *IDropTargetComObj) DragEnter(pDataObj *win32.IDataObject, grfKeyState win32.MODIFIERKEYS_FLAGS,
	pt win32.POINTL, pdwEffect *win32.DROPEFFECT) uintptr {
	return uintptr(this.impl().DragEnter(pDataObj, grfKeyState, pt, pdwEffect))
}

func (this *IDropTargetComObj) DragOver(grfKeyState win32.MODIFIERKEYS_FLAGS,
	pt win32.POINTL, pdwEffect *win32.DROPEFFECT) uintptr {
	return uintptr(this.impl().DragOver(grfKeyState, pt, pdwEffect))
}

func (this *IDropTargetComObj) DragLeave() uintptr {
	return uintptr(this.impl().DragLeave())
}

func (this *IDropTargetComObj) Drop(pDataObj *win32.IDataObject, grfKeyState win32.MODIFIERKEYS_FLAGS,
	pt win32.POINTL, pdwEffect *win32.DROPEFFECT) uintptr {
	return uintptr(this.impl().Drop(pDataObj, grfKeyState, pt, pdwEffect))
}

var _pIDropTargetVtbl *win32.IDropTargetVtbl

func (this *IDropTargetComObj) BuildVtbl(lock bool) *win32.IDropTargetVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIDropTargetVtbl != nil {
		return _pIDropTargetVtbl
	}
	_pIDropTargetVtbl = &win32.IDropTargetVtbl{
		IUnknownVtbl: *this.IUnknownComObj.BuildVtbl(false),
		DragEnter:    syscall.NewCallback((*IDropTargetComObj).DragEnter),
		DragOver:     syscall.NewCallback((*IDropTargetComObj).DragOver),
		DragLeave:    syscall.NewCallback((*IDropTargetComObj).DragLeave),
		Drop:         syscall.NewCallback((*IDropTargetComObj).Drop),
	}
	return _pIDropTargetVtbl
}

func (this *IDropTargetComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

func (this *IDropTargetComObj) IDropTarget() *win32.IDropTarget {
	return (*win32.IDropTarget)(unsafe.Pointer(this))
}

func NewIDropTargetComObj(impl win32.IDropTargetInterface) *IDropTargetComObj {
	comObj := com.NewComObj[IDropTargetComObj](impl)
	return comObj
}
