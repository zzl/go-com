package ole

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
	"syscall"
	"unsafe"
)

type IConnectionPointContainerImplSupport struct {
	//com.IUnknownImpl without this, how to access IUnknownImpl related data?, with this, comes ambiguous
	//unnecessary pUnknownImpl *com.IUnknownImpl

	ConnectPoints []*IConnectionPointComObj
}

func (this *IConnectionPointContainerImplSupport) OnComObjFree() {
	for _, cp := range this.ConnectPoints {
		cp.Release()
	}
}

func (this *IConnectionPointContainerImplSupport) EnumConnectionPoints(ppEnum **win32.IEnumConnectionPoints) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IConnectionPointContainerImplSupport) FindConnectionPoint(riid *syscall.GUID, ppCP **win32.IConnectionPoint) win32.HRESULT {
	for _, cp := range this.ConnectPoints {
		var sourceId syscall.GUID
		cp.GetConnectionInterface(&sourceId)
		if sourceId == *riid {
			hr := cp.QueryInterface(&win32.IID_IConnectionPoint, unsafe.Pointer(ppCP))
			return win32.HRESULT(hr)
		}
	}
	return win32.CONNECT_E_NOCONNECTION
}

//
type IConnectionPointContainerComObj struct {
	com.IUnknownComObj
}

func (this *IConnectionPointContainerComObj) IID() *syscall.GUID {
	return &win32.IID_IConnectionPointContainer
}

func (this *IConnectionPointContainerComObj) impl() win32.IConnectionPointContainerInterface {
	return this.Impl().(win32.IConnectionPointContainerInterface)
}

func (this *IConnectionPointContainerComObj) EnumConnectionPoints(ppEnum **win32.IEnumConnectionPoints) uintptr {
	return uintptr(this.impl().EnumConnectionPoints(ppEnum))
}

func (this *IConnectionPointContainerComObj) FindConnectionPoint(riid *syscall.GUID, ppCP **win32.IConnectionPoint) uintptr {
	return uintptr(this.impl().FindConnectionPoint(riid, ppCP))
}

var _pIConnectionPointContainerVtbl *win32.IConnectionPointContainerVtbl

func (this *IConnectionPointContainerComObj) BuildVtbl(lock bool) *win32.IConnectionPointContainerVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIConnectionPointContainerVtbl != nil {
		return _pIConnectionPointContainerVtbl
	}
	_pIConnectionPointContainerVtbl = &win32.IConnectionPointContainerVtbl{
		IUnknownVtbl:         *this.IUnknownComObj.BuildVtbl(false),
		EnumConnectionPoints: syscall.NewCallback((*IConnectionPointContainerComObj).EnumConnectionPoints),
		FindConnectionPoint:  syscall.NewCallback((*IConnectionPointContainerComObj).FindConnectionPoint),
	}
	return _pIConnectionPointContainerVtbl
}

func (this *IConnectionPointContainerComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

//
type IConnectionPointImpl struct {
	com.IUnknownImpl
	Container *win32.IConnectionPointContainer
	SourceIID *syscall.GUID
	Sinks     []*OleClient
}

func (this *IConnectionPointImpl) OnComObjFree() {
	for _, s := range this.Sinks {
		if s != nil {
			s.Release()
		}
	}
	this.Sinks = nil
}

func (this *IConnectionPointImpl) GetConnectionInterface(pIID *syscall.GUID) win32.HRESULT {
	*pIID = *this.SourceIID
	return win32.S_OK
}

func (this *IConnectionPointImpl) GetConnectionPointContainer(ppCPC **win32.IConnectionPointContainer) win32.HRESULT {
	*ppCPC = this.Container
	this.Container.AddRef()
	return win32.S_OK
}

func (this *IConnectionPointImpl) Advise(pUnkSink *win32.IUnknown, pdwCookie *uint32) win32.HRESULT {
	pSink := &OleClient{}
	//win32.IID_IDispatch?
	hr := pUnkSink.QueryInterface(this.SourceIID, unsafe.Pointer(&pSink.IDispatch))
	if win32.FAILED(hr) {
		return hr
	}
	cookie := -1
	for n, s := range this.Sinks {
		if s == nil {
			this.Sinks[n] = pSink
			cookie = n
			break
		}
	}
	if cookie == -1 {
		this.Sinks = append(this.Sinks, pSink)
		cookie = len(this.Sinks) - 1
	}
	return win32.S_OK
}

func (this *IConnectionPointImpl) Unadvise(dwCookie uint32) win32.HRESULT {
	index := int(dwCookie)
	if index < 0 || index > len(this.Sinks)-1 {
		return win32.E_INVALIDARG
	}
	pSink := this.Sinks[index]
	this.Sinks[index] = nil
	pSink.Release()
	return win32.S_OK
}

func (this *IConnectionPointImpl) EnumConnections(ppEnum **win32.IEnumConnections) win32.HRESULT {
	var cds []win32.CONNECTDATA
	for n, sink := range this.Sinks {
		if sink == nil {
			continue
		}
		cds = append(cds, win32.CONNECTDATA{
			DwCookie: uint32(n),
			PUnk:     sink.GetIUnknown(),
		})
		sink.AddRef()
	}
	impl := &IEnumConnectionsImpl{cds: cds}
	obj := com.NewComObj[IEnumConnectionsComObj](impl)
	obj.AssignPpvObject(unsafe.Pointer(ppEnum))
	return win32.S_OK
}

//
type IConnectionPointComObj struct {
	com.IUnknownComObj
}

func (this *IConnectionPointComObj) IID() *syscall.GUID {
	return &win32.IID_IConnectionPoint
}

func (this *IConnectionPointComObj) impl() win32.IConnectionPointInterface {
	return this.Impl().(win32.IConnectionPointInterface)
}

func (this *IConnectionPointComObj) GetConnectionInterface(pIID *syscall.GUID) uintptr {
	return uintptr(this.impl().GetConnectionInterface(pIID))
}

func (this *IConnectionPointComObj) GetConnectionPointContainer(ppCPC **win32.IConnectionPointContainer) uintptr {
	return uintptr(this.impl().GetConnectionPointContainer(ppCPC))
}

func (this *IConnectionPointComObj) Advise(pUnkSink *win32.IUnknown, pdwCookie *uint32) uintptr {
	return uintptr(this.impl().Advise(pUnkSink, pdwCookie))
}

func (this *IConnectionPointComObj) Unadvise(dwCookie uint32) uintptr {
	return uintptr(this.impl().Unadvise(dwCookie))
}

func (this *IConnectionPointComObj) EnumConnections(ppEnum **win32.IEnumConnections) uintptr {
	return uintptr(this.impl().EnumConnections(ppEnum))
}

var _pIConnectionPointVtbl *win32.IConnectionPointVtbl

func (this *IConnectionPointComObj) BuildVtbl(lock bool) *win32.IConnectionPointVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIConnectionPointVtbl != nil {
		return _pIConnectionPointVtbl
	}
	_pIConnectionPointVtbl = &win32.IConnectionPointVtbl{
		IUnknownVtbl:                *this.IUnknownComObj.BuildVtbl(false),
		GetConnectionInterface:      syscall.NewCallback((*IConnectionPointComObj).GetConnectionInterface),
		GetConnectionPointContainer: syscall.NewCallback((*IConnectionPointComObj).GetConnectionPointContainer),
		Advise:                      syscall.NewCallback((*IConnectionPointComObj).Advise),
		Unadvise:                    syscall.NewCallback((*IConnectionPointComObj).Unadvise),
		EnumConnections:             syscall.NewCallback((*IConnectionPointComObj).EnumConnections),
	}
	return _pIConnectionPointVtbl
}

func (this *IConnectionPointComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

//
type IProvideClassInfoImplSupport struct {
	//unnecessary pUnknownImpl *com.IUnknownImpl
	TypeInfo *win32.ITypeInfo
}

func (this *IProvideClassInfoImplSupport) GetClassInfo(ppTI **win32.ITypeInfo) win32.HRESULT {
	*ppTI = this.TypeInfo
	this.TypeInfo.AddRef()
	return win32.S_OK
}

func (this *IProvideClassInfoImplSupport) OnComObjFree() {
	this.TypeInfo.Release()
}

//
type IProvideClassInfoComObj struct {
	com.IUnknownComObj
}

func (this *IProvideClassInfoComObj) IID() *syscall.GUID {
	return &win32.IID_IProvideClassInfo
}

func (this *IProvideClassInfoComObj) impl() win32.IProvideClassInfoInterface {
	return this.Impl().(win32.IProvideClassInfoInterface)
}

func (this *IProvideClassInfoComObj) GetClassInfo(ppTI **win32.ITypeInfo) uintptr {
	return uintptr(this.impl().GetClassInfo(ppTI))
}

var _pIProvideClassInfoVtbl *win32.IProvideClassInfoVtbl

func (this *IProvideClassInfoComObj) BuildVtbl(lock bool) *win32.IProvideClassInfoVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIProvideClassInfoVtbl != nil {
		return _pIProvideClassInfoVtbl
	}
	_pIProvideClassInfoVtbl = &win32.IProvideClassInfoVtbl{
		IUnknownVtbl: *this.IUnknownComObj.BuildVtbl(false),
		GetClassInfo: syscall.NewCallback((*IProvideClassInfoComObj).GetClassInfo),
	}
	return _pIProvideClassInfoVtbl
}

func (this *IProvideClassInfoComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

//
type IProvideClassInfo2Impl struct {
	IProvideClassInfoImplSupport
	DefaultSourceIID *syscall.GUID
}

func (this *IProvideClassInfo2Impl) GetGUID(dwGuidKind uint32, pGUID *syscall.GUID) win32.HRESULT {
	if dwGuidKind != uint32(win32.GUIDKIND_DEFAULT_SOURCE_DISP_IID) || pGUID == nil {
		return win32.E_INVALIDARG
	}
	*pGUID = *this.DefaultSourceIID
	return win32.S_OK
}

//
type IProvideClassInfo2ComObj struct {
	IProvideClassInfoComObj
}

func (this *IProvideClassInfo2ComObj) IID() *syscall.GUID {
	return &win32.IID_IProvideClassInfo2
}

func (this *IProvideClassInfo2ComObj) impl() win32.IProvideClassInfo2Interface {
	return this.Impl().(win32.IProvideClassInfo2Interface)
}

func (this *IProvideClassInfo2ComObj) GetGUID(dwGuidKind uint32, pGUID *syscall.GUID) uintptr {
	return uintptr(this.impl().GetGUID(dwGuidKind, pGUID))
}

var _pIProvideClassInfo2Vtbl *win32.IProvideClassInfo2Vtbl

func (this *IProvideClassInfo2ComObj) BuildVtbl(lock bool) *win32.IProvideClassInfo2Vtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIProvideClassInfo2Vtbl != nil {
		return _pIProvideClassInfo2Vtbl
	}
	_pIProvideClassInfo2Vtbl = &win32.IProvideClassInfo2Vtbl{
		IProvideClassInfoVtbl: *this.IProvideClassInfoComObj.BuildVtbl(false),
		GetGUID:               syscall.NewCallback((*IProvideClassInfo2ComObj).GetGUID),
	}
	return _pIProvideClassInfo2Vtbl
}

func (this *IProvideClassInfo2ComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

//
type IEnumConnectionsImpl struct {
	com.IUnknownImpl
	cds       []win32.CONNECTDATA
	nextIndex int
}

func (this *IEnumConnectionsImpl) OnComObjFree() {
	for _, cd := range this.cds {
		cd.PUnk.Release()
	}
	this.cds = nil
}

func (this *IEnumConnectionsImpl) Next(cConnections uint32, rgcd *win32.CONNECTDATA, pcFetched *uint32) win32.HRESULT {
	if cConnections == 0 {
		return win32.S_OK
	} else if cConnections > 1 && pcFetched == nil {
		return win32.S_FALSE
	}
	rgcds := unsafe.Slice(rgcd, cConnections)
	index := this.nextIndex
	for n := 0; n < int(cConnections); n++ {
		if index > len(this.cds)-1 {
			break
		}
		rgcds[n] = this.cds[index]
		rgcds[n].PUnk.AddRef() //
		index += 1
	}
	if index == this.nextIndex {
		return win32.S_FALSE
	}
	if pcFetched != nil {
		*pcFetched = uint32(index - this.nextIndex)
	}
	this.nextIndex = index
	return win32.S_OK
}

func (this *IEnumConnectionsImpl) Skip(cConnections uint32) win32.HRESULT {
	newNextIndex := this.nextIndex + int(cConnections)
	if newNextIndex <= len(this.cds) {
		this.nextIndex = newNextIndex
		return win32.S_OK
	} else {
		return win32.S_FALSE
	}
}

func (this *IEnumConnectionsImpl) Reset() win32.HRESULT {
	this.nextIndex = 0
	return win32.S_OK
}

func (this *IEnumConnectionsImpl) Clone(ppEnum **win32.IEnumConnections) win32.HRESULT {
	var cdsClone []win32.CONNECTDATA
	for _, cd := range this.cds {
		cd.PUnk.AddRef()
		cdsClone = append(cdsClone, cd)
	}
	implClone := &IEnumConnectionsImpl{cds: cdsClone}
	objClone := com.NewComObj[IEnumConnectionsComObj](implClone)
	objClone.AssignPpvObject(unsafe.Pointer(ppEnum))
	return win32.S_OK
}

//
type IEnumConnectionsComObj struct {
	com.IUnknownComObj
}

func (this *IEnumConnectionsComObj) IID() *syscall.GUID {
	return &win32.IID_IEnumConnections
}

func (this *IEnumConnectionsComObj) impl() win32.IEnumConnectionsInterface {
	return this.Impl().(win32.IEnumConnectionsInterface)
}

func (this *IEnumConnectionsComObj) Next(cConnections uint32, rgcd *win32.CONNECTDATA, pcFetched *uint32) uintptr {
	return uintptr(this.impl().Next(cConnections, rgcd, pcFetched))
}

func (this *IEnumConnectionsComObj) Skip(cConnections uint32) uintptr {
	return uintptr(this.impl().Skip(cConnections))
}

func (this *IEnumConnectionsComObj) Reset() uintptr {
	return uintptr(this.impl().Reset())
}

func (this *IEnumConnectionsComObj) Clone(ppEnum **win32.IEnumConnections) uintptr {
	return uintptr(this.impl().Clone(ppEnum))
}

var _pIEnumConnectionsVtbl *win32.IEnumConnectionsVtbl

func (this *IEnumConnectionsComObj) BuildVtbl(lock bool) *win32.IEnumConnectionsVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIEnumConnectionsVtbl != nil {
		return _pIEnumConnectionsVtbl
	}
	_pIEnumConnectionsVtbl = &win32.IEnumConnectionsVtbl{
		IUnknownVtbl: *this.IUnknownComObj.BuildVtbl(false),
		Next:         syscall.NewCallback((*IEnumConnectionsComObj).Next),
		Skip:         syscall.NewCallback((*IEnumConnectionsComObj).Skip),
		Reset:        syscall.NewCallback((*IEnumConnectionsComObj).Reset),
		Clone:        syscall.NewCallback((*IEnumConnectionsComObj).Clone),
	}
	return _pIEnumConnectionsVtbl
}

func (this *IEnumConnectionsComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}
