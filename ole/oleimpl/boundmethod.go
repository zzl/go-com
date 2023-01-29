package oleimpl

import (
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

type BoundMethodDispImpl struct {
	ole.IDispatchImpl
	ownerDispObj *win32.IDispatch
	memberDispId int32
}

func (this *BoundMethodDispImpl) GetIDsOfNames(riid *syscall.GUID, rgszNames *win32.PWSTR,
	cNames uint32, lcid uint32, rgDispId *int32) win32.HRESULT {
	return win32.DISP_E_UNKNOWNNAME
}

func (this *BoundMethodDispImpl) Invoke(dispIdMember int32, riid *syscall.GUID,
	lcid uint32, wFlags win32.DISPATCH_FLAGS, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT,
	pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {
	if dispIdMember == int32(win32.DISPID_VALUE) {
		return this.ownerDispObj.Invoke(this.memberDispId, riid, lcid,
			wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr)
	}
	return win32.E_NOTIMPL
}

func (this *BoundMethodDispImpl) OnComObjFree() {
	this.ownerDispObj.Release()
}

func NewBoundMethodDispatch(pDispOwner *win32.IDispatch, memberDispId int32) *win32.IDispatch {
	pDisp := ole.NewIDispatchComObject(&BoundMethodDispImpl{
		ownerDispObj: pDispOwner,
		memberDispId: memberDispId,
	}).IDispatch()
	return pDisp
}
