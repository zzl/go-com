package oleimpl

import (
	"strings"
	"syscall"
	"unsafe"

	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/win32"
)

type VariantFunc func(args ...*ole.Variant) *ole.Variant

type VariantProp struct {
	Get VariantFunc
	Set VariantFunc
}

type FuncMapDispImpl struct {
	ole.IDispatchImpl

	fNames []string
	funcs  []VariantFunc
	pNames []string
	props  []VariantProp

	OnFinalize func()
}

func (this *FuncMapDispImpl) OnComObjFree() {
	if this.OnFinalize != nil {
		this.OnFinalize()
	}
}

func (this *FuncMapDispImpl) GetIDsOfNames(riid *syscall.GUID, rgszNames *win32.PWSTR,
	cNames uint32, lcid uint32, rgDispId *int32) win32.HRESULT {
	if cNames != 1 {
		return win32.E_INVALIDARG
	}
	name := win32.PwstrToStr(*rgszNames)
	name = strings.ToLower(name)
	for n, fName := range this.fNames {
		if fName == name {
			*rgDispId = int32(n) + 1
			return win32.S_OK
		}
	}
	for n, pName := range this.pNames {
		if pName == name {
			*rgDispId = int32(n+len(this.fNames)) + 1
			return win32.S_OK
		}
	}
	return win32.DISP_E_UNKNOWNNAME
}

func (this *FuncMapDispImpl) Invoke(dispIdMember int32, riid *syscall.GUID,
	lcid uint32, wFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT,
	pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {

	vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 9)

	funcIdx := int(dispIdMember) - 1
	if funcIdx >= 0 && funcIdx < len(this.funcs) {
		if wFlags == uint16(win32.DISPATCH_PROPERTYGET) { //
			if pDispParams.CArgs == 0 {
				pDispThis := (*win32.IDispatch)(unsafe.Pointer(this.ComObj.GetIUnknownComObj()))
				pDisp := NewBoundMethodDispatch(pDispThis, dispIdMember)
				*(*ole.Variant)(pVarResult) = *ole.NewVariantDispatch(pDisp)
				return win32.S_OK
			} else {
				return win32.E_NOTIMPL
			}
		}

		pvRet := this.funcs[funcIdx](vArgs...)
		if pVarResult != nil && pvRet != nil {
			*pVarResult = win32.VARIANT(*pvRet)
		}
		return win32.S_OK
	} else if propIdx := funcIdx - len(this.funcs); propIdx >= 0 && propIdx < len(this.props) {
		prop := this.props[propIdx]
		var f VariantFunc
		if wFlags == uint16(win32.DISPATCH_PROPERTYGET) {
			f = prop.Get
		} else if wFlags == uint16(win32.DISPATCH_PROPERTYPUT) ||
			wFlags == uint16(win32.DISPATCH_PROPERTYPUTREF) {
			f = prop.Set
		}
		if f == nil {
			return win32.E_NOTIMPL
		} else {
			pvRet := f(vArgs...)
			if pVarResult != nil && pvRet != nil {
				*pVarResult = win32.VARIANT(*pvRet)
			}
			return win32.S_OK
		}
	}
	return win32.E_NOTIMPL
}

//
func NewFuncMapDispatch(funcMap map[string]VariantFunc,
	propMap map[string]VariantProp, onFinalize func()) *win32.IDispatch {
	var fNames []string
	var funcs []VariantFunc
	for name, f := range funcMap {
		fNames = append(fNames, strings.ToLower(name))
		funcs = append(funcs, f)
	}
	var pNames []string
	var props []VariantProp
	for name, p := range propMap {
		pNames = append(pNames, strings.ToLower(name))
		props = append(props, p)
	}
	pDisp := ole.NewIDispatch(&FuncMapDispImpl{
		fNames:     fNames,
		funcs:      funcs,
		pNames:     pNames,
		props:      props,
		OnFinalize: onFinalize,
	})
	return pDisp
}
