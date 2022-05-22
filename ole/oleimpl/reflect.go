package oleimpl

import (
	"reflect"
	"strings"
	"syscall"

	"github.com/zzl/go-com/ole"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
)

type ReflectDispMember struct {
	Name string
	//
	CallFuncValue *reflect.Value
	GetFuncValue  *reflect.Value
	SetFuncValue  *reflect.Value
}

type ReflectDispImpl struct {
	ole.IDispatchImpl
	Obj interface{}

	memIdMap map[string]int //name:dispid(1 base)
	members  []*ReflectDispMember
}

func (this *ReflectDispImpl) Init() {
	objVal := reflect.ValueOf(this.Obj)
	objType := objVal.Type()
	funcCount := objVal.NumMethod()
	funcCount2 := objType.NumMethod()
	if funcCount != funcCount2 {
		panic("??")
	}
	this.memIdMap = make(map[string]int)
	memMap := make(map[string]*ReflectDispMember)
	for n := 0; n < funcCount; n++ {
		methodVal := objVal.Method(n)
		_ = methodVal
		method := objType.Method(n)
		name := method.Name

		var namePrefix string
		if len(name) > 3 {
			namePrefix = name[:3]
			if namePrefix == "Get" || namePrefix == "Set" {
				name = name[3:]
			} else {
				namePrefix = ""
			}
		}
		mem := memMap[name]
		if mem == nil {
			mem = &ReflectDispMember{}
			memMap[name] = mem
			this.members = append(this.members, mem)
			this.memIdMap[strings.ToLower(name)] = len(this.members)
		}
		if namePrefix == "Get" {
			mem.GetFuncValue = &methodVal
			mem.Name = name
		} else if namePrefix == "Set" {
			mem.SetFuncValue = &methodVal
			mem.Name = name
		} else {
			mem.CallFuncValue = &methodVal
			mem.Name = name
		}
	}
	println("?")
}

func (this *ReflectDispImpl) GetIDsOfNames(riid *syscall.GUID, rgszNames *win32.PWSTR,
	cNames uint32, lcid uint32, rgDispId *int32) win32.HRESULT {
	if cNames != 1 {
		return win32.E_INVALIDARG
	}
	name := win32.PwstrToStr(*rgszNames)
	name = strings.ToLower(name)
	if dispId, ok := this.memIdMap[name]; ok {
		*rgDispId = int32(dispId)
		return win32.S_OK
	}
	return win32.DISP_E_UNKNOWNNAME
}

func (this *ReflectDispImpl) Invoke(dispIdMember int32, riid *syscall.GUID,
	lcid uint32, wFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT,
	pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {
	dispId := int(dispIdMember)
	if dispId == 0 {
		return win32.E_NOTIMPL //?
	} else if dispId > len(this.members) {
		return win32.E_INVALIDARG
	}
	member := this.members[dispId-1]
	var funcValue *reflect.Value
	if wFlags&uint16(win32.DISPATCH_METHOD) != 0 {
		funcValue = member.CallFuncValue
	} else if wFlags == uint16(win32.DISPATCH_PROPERTYGET) {
		funcValue = member.GetFuncValue
		if funcValue == nil {
			if member.CallFuncValue != nil && pDispParams.CArgs == 0 {
				pDispThis := (*win32.IDispatch)(this.ComObject.Pointer())
				pDisp := NewBoundMethodDispatch(pDispThis, dispIdMember)
				*(*ole.Variant)(pVarResult) = *ole.NewVariantDispatch(pDisp)
				return win32.S_OK
			} else {
				return win32.E_NOTIMPL
			}
		}
	} else {
		funcValue = member.SetFuncValue
	}
	if funcValue == nil {
		return win32.DISP_E_MEMBERNOTFOUND
	}

	ft := funcValue.Type()
	numIn := ft.NumIn()
	numOut := ft.NumOut()

	vArgs, srcArgIndexes := ole.ProcessInvokeArgs(pDispParams, numIn)
	argVals := make([]reflect.Value, numIn)
	for n, vArg := range vArgs {
		arg, err := this.readVariantArgValue(ft.In(n), vArg)
		if err != nil {
			if puArgErr != nil {
				*puArgErr = uint32(srcArgIndexes[n])
			}
			if ce, ok := err.(com.Error); ok {
				return ce.HRESULT()
			} else {
				return win32.DISP_E_TYPEMISMATCH
			}
		}
		argVals[n] = reflect.ValueOf(arg)
	}
	retVals := funcValue.Call(argVals)

	if numOut == 0 {
		//
	} else if numOut == 1 {
		var vResult ole.Variant
		var unwrapActions ole.Actions
		ole.SetVariantParam(&vResult, retVals[0].Interface(), &unwrapActions)
		*(*ole.Variant)(pVarResult) = *vResult.Copy()
		unwrapActions.Execute()
	}
	println("?")

	return win32.S_OK
}

func (this *ReflectDispImpl) readVariantArgValue(typ reflect.Type, v *ole.Variant) (interface{}, error) {
	return v.ValueOfType(typ)
}

func NewReflectDispImpl(obj interface{}) *ReflectDispImpl {
	impl := &ReflectDispImpl{Obj: obj}
	impl.Init()
	return impl
}

func NewReflectDispatch(obj interface{}) *win32.IDispatch {
	pDisp := ole.NewIDispatch(NewReflectDispImpl(obj))
	return pDisp
}
