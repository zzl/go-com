package ole

import (
	"time"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
)

type IDispatchProvider interface {
	GetIDispatch(addRef bool) *win32.IDispatch
}

type OleClientProvider interface {
	GetOleClient() *OleClient
}

type OleClient struct {
	*win32.IDispatch
}

func CreateClient(clsid *win32.CLSID) (*OleClient, error) {
	var pDisp *win32.IDispatch
	hr := win32.CoCreateInstance(clsid, nil, win32.CLSCTX_LOCAL_SERVER,
		&win32.IID_IDispatch, unsafe.Pointer(&pDisp))
	if win32.FAILED(hr) {
		err := com.NewError(hr)
		com.SetLastError(err)
		return nil, err
	}
	return &OleClient{pDisp}, nil
}

func (this *OleClient) Dispose() {
	//
}

func (this *OleClient) GetOleClient() *OleClient {
	return this
}

func (this *OleClient) PropPut(dispId win32.DISPID,
	reqArgs []interface{}, optArgs ...interface{}) error {

	dispParams, unwrapActions := this.buildDispParams(optArgs, reqArgs)

	named := win32.DISPID_PROPERTYPUT

	dispParams.CNamedArgs = 1
	dispParams.RgdispidNamedArgs = &named

	hr := this.Invoke(dispId, &win32.IID_NULL, win32.LOCALE_INVARIANT,
		win32.DISPATCH_PROPERTYPUT, &dispParams, nil, nil, nil)
	unwrapActions.Execute()
	if win32.FAILED(hr) {
		err := com.NewError(hr)
		com.SetLastError(err)
		return err
	}
	return nil
}

func (this *OleClient) PropPutRef(dispId win32.DISPID,
	reqArgs []interface{}, optArgs ...interface{}) error {

	dispParams, unwrapActions := this.buildDispParams(optArgs, reqArgs)

	named := win32.DISPID_PROPERTYPUT
	dispParams.CNamedArgs = 1
	dispParams.RgdispidNamedArgs = &named

	hr := this.Invoke(dispId, &win32.IID_NULL, win32.LOCALE_INVARIANT,
		win32.DISPATCH_PROPERTYPUTREF, &dispParams, nil, nil, nil)
	unwrapActions.Execute()
	if win32.FAILED(hr) {
		err := com.NewError(hr)
		com.SetLastError(err)
		return err
	}
	return nil
}

func (this *OleClient) PropGet(dispId win32.DISPID,
	reqArgs []interface{}, optArgs ...interface{}) (*Variant, error) {

	dispParams, unwrapActions := this.buildDispParams(optArgs, reqArgs)

	var result Variant
	hr := this.Invoke(dispId, &win32.IID_NULL, win32.LOCALE_INVARIANT,
		win32.DISPATCH_PROPERTYGET, &dispParams, (*win32.VARIANT)(&result), nil, nil)
	unwrapActions.Execute()
	if win32.FAILED(hr) {
		err := com.NewError(hr)
		com.SetLastError(err)
		return &result, err
	}
	return &result, nil
}

func (this *OleClient) Call(dispId win32.DISPID,
	reqArgs []interface{}, optArgs ...interface{}) (*Variant, error) {

	dispParams, unwrapActions := this.buildDispParams(optArgs, reqArgs)
	var result Variant

	hr := this.Invoke(dispId, &win32.IID_NULL, win32.LOCALE_INVARIANT,
		win32.DISPATCH_METHOD, &dispParams, (*win32.VARIANT)(&result), nil, nil)

	unwrapActions.Execute()
	if win32.FAILED(hr) {
		err := com.NewError(hr)
		com.SetLastError(err)
		return &result, err
	}
	return &result, nil
}

func (this *OleClient) buildDispParams(optArgs []interface{},
	reqArgs []interface{}) (win32.DISPPARAMS, Actions) {

	optArgc := len(optArgs)
	totalArgc := len(reqArgs) + optArgc
	vs := make([]Variant, totalArgc)
	var unwrapActions Actions
	for n, a := range reqArgs {
		SetVariantParam(&vs[totalArgc-n-1], a, &unwrapActions)
	}
	for n, a := range optArgs {
		SetVariantParam(&vs[optArgc-n-1], a, &unwrapActions)
	}
	dispParams := win32.DISPPARAMS{
		CArgs: uint32(totalArgc),
	}
	if totalArgc > 0 {
		dispParams.Rgvarg = (*win32.VARIANT)(&vs[0])
	}
	return dispParams, unwrapActions
}

func VariantFromValue(value interface{}) *Variant {
	var v Variant
	if value == nil {
		return &v
	}
	SetVariantParam(&v, value, nil)
	return &v
}

func SetVariantParam(v *Variant, value interface{}, unwrapActions *Actions) {
	if value == nil {
		v.Vt = win32.VT_ERROR
		*v.Scode() = win32.DISP_E_PARAMNOTFOUND
		return
	}
	switch val := value.(type) {
	case Variant:
		*v = val
	case *Variant:
		v.Vt = win32.VT_VARIANT | win32.VT_BYREF
		*v.PvarVal() = (*win32.VARIANT)(val)
	case int8:
		v.Vt = win32.VT_I1
		*v.CVal() = win32.CHAR(val)
	case uint8:
		v.Vt = win32.VT_UI1
		*v.BVal() = val
	case int16:
		v.Vt = win32.VT_I2
		*v.IVal() = val
	case uint16:
		v.Vt = win32.VT_UI2
		*v.UiVal() = val
	case int32:
		v.Vt = win32.VT_I4
		*v.LVal() = val
	case uint32:
		v.Vt = win32.VT_UI4
		*v.UlVal() = val
	case int64:
		v.Vt = win32.VT_I8
		*v.LlVal() = val
	case uint64:
		v.Vt = win32.VT_UI8
		*v.UllVal() = val
	case float32:
		v.Vt = win32.VT_R4
		*v.FltVal() = val
	case float64:
		v.Vt = win32.VT_R8
		*v.DblVal() = val
	case time.Time:
		v.Vt = win32.VT_DATE
		//?*v.DateVal() =
	case string:
		v.Vt = win32.VT_BSTR
		bs := win32.StrToBstr(val)
		*v.BstrVal() = bs
		unwrapActions.Add(func() {
			win32.SysFreeString(bs)
		})
	case *int8:
		v.Vt = win32.VT_I1 | win32.VT_BYREF
		*v.PcVal() = (*win32.CHAR)(unsafe.Pointer(val))
	case *uint8:
		v.Vt = win32.VT_UI1 | win32.VT_BYREF
		*v.PbVal() = val
	case *int16:
		v.Vt = win32.VT_UI2 | win32.VT_BYREF
		*v.PiVal() = val
	case *uint16:
		v.Vt = win32.VT_UI2 | win32.VT_BYREF
		*v.PuiVal() = val
	case *int32:
		v.Vt = win32.VT_I4 | win32.VT_BYREF
		*v.PlVal() = val
	case *uint32:
		v.Vt = win32.VT_UI4 | win32.VT_BYREF
		*v.PulVal() = val
	case *int64:
		v.Vt = win32.VT_I8 | win32.VT_BYREF
		*v.PllVal() = val
	case *uint64:
		v.Vt = win32.VT_UI8 | win32.VT_BYREF
		*v.PullVal() = val
	case *float32:
		v.Vt = win32.VT_R4 | win32.VT_BYREF
		*v.PfltVal() = val
	case *float64:
		v.Vt = win32.VT_R8 | win32.VT_BYREF
		*v.PdblVal() = val
	case *time.Time:
		//v.Vt = win32.VT_DATE
	case *string:
		v.Vt = win32.VT_BSTR | win32.VT_BYREF
		bs := com.NewBStringFromStr(*val)
		*v.PbstrVal() = bs.PBSTR()
		unwrapActions.Add(func() {
			*val = bs.ToStringAndFree()
		})
	case bool:
		v.Vt = win32.VT_BOOL
		if val {
			*v.BoolVal() = win32.VARIANT_TRUE
		} else {
			*v.BoolVal() = win32.VARIANT_FALSE
		}
	case *bool:
		v.Vt = win32.VT_BOOL | win32.VT_BYREF
		var val2 win32.VARIANT_BOOL
		if *val {
			val2 = win32.VARIANT_TRUE
		}
		*v.PboolVal() = &val2
		unwrapActions.Add(func() {
			*val = val2 == win32.VARIANT_TRUE
		})
	case *win32.BSTR:
		*v.PbstrVal() = val
	case int:
		//v.Vt = win32.VT_INT
		v.Vt = win32.VT_I4
		*v.IntVal() = int32(val)
	case uint:
		//v.Vt = win32.VT_UINT
		v.Vt = win32.VT_UI4
		*v.UintVal() = uint32(val)
	case *int:
		//v.Vt = win32.VT_INT | win32.VT_BYREF
		v.Vt = win32.VT_I4 | win32.VT_BYREF
		val2 := int32(*val)
		*v.PintVal() = &val2
		unwrapActions.Add(func() {
			*val = int(val2)
		})
	case *uint:
		//v.Vt = win32.VT_UINT | win32.VT_BYREF
		v.Vt = win32.VT_UI4 | win32.VT_BYREF
		val2 := uint32(*val)
		*v.PuintVal() = &val2
		unwrapActions.Add(func() {
			*val = uint(val2)
		})

	case *win32.IUnknown:
		v.Vt = win32.VT_UNKNOWN
		*v.PunkVal() = val
	case **win32.IUnknown:
		v.Vt = win32.VT_UNKNOWN | win32.VT_BYREF
		*v.PpunkVal() = val

	case *win32.IDispatch:
		v.Vt = win32.VT_DISPATCH
		*v.PdispVal() = val
	case **win32.IDispatch:
		v.Vt = win32.VT_DISPATCH | win32.VT_BYREF
		*v.PpdispVal() = val
	case IDispatchProvider:
		v.Vt = win32.VT_DISPATCH
		*v.PdispVal() = val.GetIDispatch(false) //?
	case SafeArrayInterface:
		psa := val.SafeArrayPtr()
		var vt win32.VARENUM
		win32.SafeArrayGetVartype(psa, &vt)
		v.Vt = win32.VT_ARRAY | vt
		*v.Parray() = psa
	case *win32.SAFEARRAY:
		v.Vt = win32.VT_ARRAY
		*v.Parray() = val
	default:
		panic("??")
	}
}

type Action func()
type Actions struct {
	actions []Action
}

func (this *Actions) Add(action Action) {
	this.actions = append(this.actions, action)
}

func (this *Actions) Execute() {
	for _, a := range this.actions {
		a()
	}
}

func getNameIndex(names []string, name string) int {
	for n, it := range names {
		if it == name {
			return n
		}
	}
	return -1
}

func ProcessOptArgs(argNames []string, optArgs []interface{}) []interface{} {
	for n, a := range optArgs {
		if namedArgs, ok := a.(NamedArgs); ok {
			optArgs = optArgs[:n]
			for k, v := range namedArgs {
				index := getNameIndex(argNames, k)
				if index == -1 {
					panic("invalid argument name: " + k)
				}
				for m := len(optArgs); m <= index; m++ {
					optArgs = append(optArgs, nil)
				}
				optArgs[index] = v
			}
			break
		}
	}
	for n := len(optArgs); n < len(argNames); n++ {
		optArgs = append(optArgs, nil)
	}
	return optArgs
}
