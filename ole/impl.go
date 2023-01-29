package ole

import (
	"syscall"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
)

type IDispatchImpl struct {
	com.IUnknownImpl
}

func (this *IDispatchImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IDispatch {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IUnknownImpl.QueryInterface(riid, ppvObject)
}

func (this *IDispatchImpl) GetTypeInfoCount(pctinfo *uint32) win32.HRESULT {
	*pctinfo = 0
	return win32.S_OK
}

func (this *IDispatchImpl) GetTypeInfo(iTInfo uint32, lcid uint32, ppTInfo **win32.ITypeInfo) win32.HRESULT {
	*ppTInfo = nil
	return win32.E_NOTIMPL
}

func (this *IDispatchImpl) GetIDsOfNames(riid *syscall.GUID, rgszNames *win32.PWSTR, cNames uint32, lcid uint32, rgDispId *int32) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IDispatchImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT, pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

type IDispatchComObj struct {
	com.IUnknownComObj
}

func (this *IDispatchComObj) impl() win32.IDispatchInterface {
	return this.Impl().(win32.IDispatchInterface)
}

func (this *IDispatchComObj) IDispatch() *win32.IDispatch {
	return (*win32.IDispatch)(unsafe.Pointer(this))
}

func (this *IDispatchComObj) GetTypeInfoCount(pctinfo *uint32) uintptr {
	return uintptr(this.impl().GetTypeInfoCount(pctinfo))
}

func (this *IDispatchComObj) GetTypeInfo(iTInfo uint32, lcid uint32,
	ppTInfo **win32.ITypeInfo) uintptr {
	return uintptr(this.impl().GetTypeInfo(iTInfo, lcid, ppTInfo))
}

func (this *IDispatchComObj) GetIDsOfNames(riid *syscall.GUID,
	rgszNames *win32.PWSTR, cNames uint32, lcid uint32, rgDispId *int32) uintptr {
	return uintptr(this.impl().GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId))
}

func (this *IDispatchComObj) Invoke(dispIdMember int32, riid *syscall.GUID,
	lcid uint32, wFlags win32.DISPATCH_FLAGS, pDispParams *win32.DISPPARAMS,
	pVarResult *win32.VARIANT, pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) uintptr {
	return uintptr(this.impl().Invoke(dispIdMember, riid, lcid,
		wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr))
}

var _pIDispatchVtbl *win32.IDispatchVtbl

func (this *IDispatchComObj) BuildVtbl(lock bool) *win32.IDispatchVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIDispatchVtbl != nil {
		return _pIDispatchVtbl
	}
	_pIDispatchVtbl = &win32.IDispatchVtbl{
		IUnknownVtbl:     *this.IUnknownComObj.BuildVtbl(false),
		GetTypeInfoCount: syscall.NewCallback((*IDispatchComObj).GetTypeInfoCount),
		GetTypeInfo:      syscall.NewCallback((*IDispatchComObj).GetTypeInfo),
		GetIDsOfNames:    syscall.NewCallback((*IDispatchComObj).GetIDsOfNames),
		Invoke:           syscall.NewCallback((*IDispatchComObj).Invoke),
	}
	return _pIDispatchVtbl
}

func (this *IDispatchComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

func NewIDispatchComObject(impl win32.IDispatchInterface) *IDispatchComObj {
	comObj := com.NewComObj[IDispatchComObj](impl)
	return comObj
}

func NewIDispatch(impl win32.IDispatchInterface) *win32.IDispatch {
	return NewIDispatchComObject(impl).IDispatch()
}
