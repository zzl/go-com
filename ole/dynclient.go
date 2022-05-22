package ole

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
	"log"
)

type DynClient struct {
	OleClient
	dispIdCache map[string]int32
}

func (this *DynClient) getDispId(name string) (int32, error) {
	if dispId, ok := this.dispIdCache[name]; ok {
		return dispId, nil
	}
	var dispId int32
	pwszName := win32.StrToPwstr(name)
	hr := this.GetIDsOfNames(&win32.IID_NULL, &pwszName, 1, 0, &dispId)
	if win32.FAILED(hr) {
		return 0, com.NewError(hr)
	}
	this.dispIdCache[name] = dispId
	return dispId, nil
}

func (this *DynClient) Get(name string) (*Variant, error) {
	dispId, err := this.getDispId(name)
	if err != nil {
		return nil, err
	}
	result, err := this.PropGet(dispId, nil, nil)
	return result, err
}

func (this *DynClient) Set(name string, value interface{}) error {
	dispId, err := this.getDispId(name)
	if err != nil {
		return err
	}
	err = this.PropPut(dispId, []interface{}{value})
	return err
}

func (this *DynClient) SetRef(name string, value interface{}) error {
	dispId, err := this.getDispId(name)
	if err != nil {
		return err
	}
	this.PropPutRef(dispId, []interface{}{value})
	return nil
}

func (this *DynClient) Call(name string, args ...interface{}) (*Variant, error) {
	dispId, err := this.getDispId(name)
	if err != nil {
		return nil, err
	}
	argCount := len(args)
	if argCount > 0 {
		lastArg := args[argCount-1]
		if namedArgs, ok := lastArg.(NamedArgs); ok {
			args = args[:argCount-1]
			result, err := this.OleClient.Call(dispId, args, namedArgs)
			return result, err
		}
	}
	result, err := this.OleClient.Call(dispId, args)
	return result, err
}

func NewDynClient(source interface{}, addRef bool, scoped bool) *DynClient {
	var pDisp *win32.IDispatch
	var err error
	switch v := source.(type) {
	case win32.IDispatchObject:
		pDisp = v.GetIDispatch_()
	case win32.VARIANT:
		pDisp, err = (*Variant)(&v).ToIDispatch()
	case *win32.VARIANT:
		pDisp, err = (*Variant)(v).ToIDispatch()
	case Variant:
		pDisp, err = v.ToIDispatch()
	case *Variant:
		pDisp, err = v.ToIDispatch()
	}
	if err != nil {
		log.Fatal(err)
	}
	if addRef {
		pDisp.AddRef()
	}
	return &DynClient{OleClient: OleClient{pDisp}, dispIdCache: make(map[string]int32)}
}
