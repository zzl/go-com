package oleimpl

import (
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/win32"
	"syscall"
	"unsafe"
)

type DualDispImpl struct {
	ole.IDispatchImpl
	TypeInfo *win32.ITypeInfo
}

func (this *DualDispImpl) GetTypeInfoCount(pctinfo *uint32) win32.HRESULT {
	*pctinfo = 1
	return win32.S_OK
}

func (this *DualDispImpl) GetTypeInfo(iTInfo uint32, lcid uint32, ppTInfo **win32.ITypeInfo) win32.HRESULT {
	*ppTInfo = nil
	if iTInfo != 0 {
		return win32.DISP_E_BADINDEX
	}
	this.TypeInfo.AddRef()
	*ppTInfo = this.TypeInfo
	return win32.E_NOTIMPL
}

func (this *DualDispImpl) GetIDsOfNames(riid *syscall.GUID, rgszNames *win32.PWSTR, cNames uint32, lcid uint32, rgDispId *int32) win32.HRESULT {
	if *riid != win32.IID_NULL {
		return win32.DISP_E_UNKNOWNINTERFACE
	}
	hr := win32.DispGetIDsOfNames(this.TypeInfo, rgszNames, cNames, rgDispId)
	return hr
}

func (this *DualDispImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT, pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {
	if *riid != win32.IID_NULL {
		return win32.DISP_E_UNKNOWNINTERFACE
	}
	hr := win32.DispInvoke(unsafe.Pointer(this.ComObj.GetIUnknownComObj()),
		this.TypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr)
	return hr
}

func (this *DualDispImpl) OnComObjFree() {
	if this.TypeInfo != nil {
		this.TypeInfo.Release()
		this.TypeInfo = nil
	}
}
