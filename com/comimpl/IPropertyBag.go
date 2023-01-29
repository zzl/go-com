package comimpl

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

type IPropertyBagImpl struct {
	com.IUnknownImpl
}

func (this *IPropertyBagImpl) Read(pszPropName win32.PWSTR,
	pVar *win32.VARIANT, pErrorLog *win32.IErrorLog) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IPropertyBagImpl) Write(pszPropName win32.PWSTR, pVar *win32.VARIANT) win32.HRESULT {
	return win32.E_NOTIMPL
}

//
type IPropertyBagComObj struct {
	com.IUnknownComObj
}

func (this *IPropertyBagComObj) IID() *syscall.GUID {
	return &win32.IID_IPropertyBag
}

func (this *IPropertyBagComObj) impl() win32.IPropertyBagInterface {
	return this.Impl().(win32.IPropertyBagInterface)
}

func (this *IPropertyBagComObj) Read(pszPropName win32.PWSTR,
	pVar *win32.VARIANT, pErrorLog *win32.IErrorLog) uintptr {
	return (uintptr)(this.impl().Read(pszPropName, pVar, pErrorLog))
}

func (this *IPropertyBagComObj) Write(pszPropName win32.PWSTR, pVar *win32.VARIANT) uintptr {
	return (uintptr)(this.impl().Write(pszPropName, pVar))
}

var _pIPropertyBagVtbl *win32.IPropertyBagVtbl

func (this *IPropertyBagComObj) BuildVtbl(lock bool) *win32.IPropertyBagVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIPropertyBagVtbl != nil {
		return _pIPropertyBagVtbl
	}
	_pIPropertyBagVtbl = (*win32.IPropertyBagVtbl)(
		com.Malloc(unsafe.Sizeof(*_pIPropertyBagVtbl)))
	*_pIPropertyBagVtbl = win32.IPropertyBagVtbl{
		IUnknownVtbl: *this.IUnknownComObj.BuildVtbl(false),
		Read:         syscall.NewCallback((*IPropertyBagComObj).Read),
		Write:        syscall.NewCallback((*IPropertyBagComObj).Write),
	}
	return _pIPropertyBagVtbl
}

func (this *IPropertyBagComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}
