package com

import (
	"sync"
	"sync/atomic"
	"syscall"
	"unsafe"

	"github.com/zzl/go-win32api/win32"
)

var MuVtbl sync.Mutex

var impls []win32.IUnknownInterface
var freeImplSlots []int32
var hHeap win32.HANDLE

func init() {
	hHeap, _ = win32.GetProcessHeap()
}

type ComObjConstraint[T any] interface {
	ComObjInterface
	*T
}

type RealObjectAware interface {
	SetRealObject(obj interface{})
}

type Initializable interface {
	Initialize()
}

type Finalizable interface {
	Finalize()
}

type ComObjLifecycleAware interface {
	OnComObjCreate()
	OnComObjFree()
}

func NewComObj[T any, PT ComObjConstraint[T]](impl win32.IUnknownInterface) *T {
	if roa, ok := impl.(RealObjectAware); ok {
		roa.SetRealObject(impl)
	}

	size := unsafe.Sizeof(*(*T)(nil))
	//println("SIZE of comobj:", size)
	p := Malloc(size)
	pComObj := (*T)(p)

	comObj := PT(pComObj)
	pUnknownComObj := (*IUnknownComObj)(p)
	pUnknownComObj.LpVtbl = comObj.GetVtbl() //
	pUnknownComObj.implSlot = AddImpl(impl)
	pUnknownComObj.cRef = 1

	//
	impl.(comObjAware).SetComObj(comObj)
	initializeComObj(comObj)

	if lsnr, ok := impl.(ComObjLifecycleAware); ok {
		lsnr.OnComObjCreate()
	}

	return pComObj
}

func Malloc(size uintptr) unsafe.Pointer {
	return win32.HeapAlloc(hHeap, win32.HEAP_ZERO_MEMORY, size)
}

func AddImpl(impl win32.IUnknownInterface) int32 {
	freeCount := len(freeImplSlots)
	if freeCount != 0 {
		index := freeImplSlots[freeCount-1]
		freeImplSlots = freeImplSlots[:freeCount-1]
		if impls[index] != nil {
			panic("?")
		}
		impls[index] = impl
		return index
	} else {
		impls = append(impls, impl)
		return int32(len(impls) - 1)
	}
}

type ComObjInterface interface {
	AddRef() uintptr
	Release() uintptr

	GetVtbl() *win32.IUnknownVtbl
	GetIUnknownComObj() *IUnknownComObj
	GetSubComObjs() []ComObjInterface
	IID() *syscall.GUID
	AssignPpvObject(ppvObject unsafe.Pointer)
}

type comObjAware interface {
	SetComObj(obj ComObjInterface)
	GetComObj() ComObjInterface
}

//
type IUnknownImpl struct {
	ComObj ComObjInterface
}

func (this *IUnknownImpl) OnComObjCreate() {
	//
}

func (this *IUnknownImpl) SetComObj(obj ComObjInterface) {
	this.ComObj = obj
}

func (this *IUnknownImpl) GetComObj() ComObjInterface {
	return this.ComObj
}

func (this *IUnknownImpl) AssignPpvObject(ppvObject unsafe.Pointer, iid ...*syscall.GUID) bool {
	if len(iid) == 0 {
		this.ComObj.AssignPpvObject(ppvObject)
		return true
	}
	theIid := *iid[0]
	var assigned bool
	if theIid == *this.ComObj.IID() {
		this.ComObj.AssignPpvObject(ppvObject)
		assigned = true
	} else {
		for _, subComObj := range this.ComObj.GetSubComObjs() {
			if *subComObj.IID() == theIid {
				subComObj.AssignPpvObject(ppvObject)
				assigned = true
				break
			}
		}
	}
	return assigned
}

func (this *IUnknownImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IUnknown {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	//println(win32.GuidToStr(riid))
	if this.AssignPpvObject(ppvObject, riid) {
		this.AddRef()
		return win32.S_OK
	}
	return win32.E_NOINTERFACE
}

func (this *IUnknownImpl) AddRef() uint32 {
	return uint32(this.ComObj.AddRef())
}

func (this *IUnknownImpl) Release() uint32 {
	return uint32(this.ComObj.Release())
}

func (this *IUnknownImpl) OnComObjFree() {
	//
}

type IUnknownComObj struct {
	LpVtbl   *win32.IUnknownVtbl
	cRef     int32
	implSlot int32
	Parent   *IUnknownComObj
}

func (this *IUnknownComObj) AssignPpvObject(ppvObject unsafe.Pointer) {
	*(*unsafe.Pointer)(ppvObject) = unsafe.Pointer(this)
}

func (this *IUnknownComObj) IID() *syscall.GUID {
	return &win32.IID_IUnknown
}

func (this *IUnknownComObj) GetIUnknownComObj() *IUnknownComObj {
	return this
}

func initializeComObj(comObj ComObjInterface) {
	if initializable, ok := comObj.(Initializable); ok {
		initializable.Initialize()
	}
	pComObj := comObj.GetIUnknownComObj()
	for _, subComObj := range comObj.GetSubComObjs() {
		pSubComObj := subComObj.GetIUnknownComObj()
		pSubComObj.Parent = pComObj
		pSubComObj.LpVtbl = subComObj.GetVtbl()
		pSubComObj.implSlot = pComObj.implSlot //?
		initializeComObj(pSubComObj)
	}
}

func (this *IUnknownComObj) GetSubComObjs() []ComObjInterface {
	return nil
}

func (this *IUnknownComObj) SetVtbl(vtbl *win32.IUnknownVtbl) {
	this.LpVtbl = vtbl
}

func (this *IUnknownComObj) Impl() win32.IUnknownInterface {
	return impls[this.implSlot]
}

func (this *IUnknownComObj) free() {
	impl := this.Impl()
	if lsnr, ok := impl.(ComObjLifecycleAware); ok {
		lsnr.OnComObjFree()
	}
	comObj := impl.(comObjAware).GetComObj()
	if finalizable, ok := comObj.(Finalizable); ok {
		finalizable.Finalize()
	}
	//
	impls[this.implSlot] = nil
	freeImplSlots = append(freeImplSlots, this.implSlot)
	bOk, err := win32.HeapFree(hHeap, 0, unsafe.Pointer(this))
	if err != win32.NO_ERROR || bOk != win32.TRUE {
		println("??")
	}
}

func (this *IUnknownComObj) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) uintptr {
	if this.Parent != nil {
		return this.Parent.QueryInterface(riid, ppvObject)
	}
	return (uintptr)(this.Impl().QueryInterface(riid, ppvObject))
}

func (this *IUnknownComObj) AddRef() uintptr {
	if this.Parent != nil {
		return this.Parent.AddRef()
	}
	cRef := atomic.AddInt32(&this.cRef, 1)
	return uintptr(cRef)
}

func (this *IUnknownComObj) Release() uintptr {
	if this.Parent != nil {
		return this.Parent.Release()
	}
	cRef := atomic.AddInt32(&this.cRef, -1)
	if cRef == 0 {
		this.free()
	}
	return uintptr(cRef)
}

func (this *IUnknownComObj) IUnknown() *win32.IUnknown {
	if this.Parent != nil {
		return this.Parent.IUnknown()
	}
	return (*win32.IUnknown)(unsafe.Pointer(this))
}

var _pIUnknownVtbl *win32.IUnknownVtbl

func (this *IUnknownComObj) BuildVtbl(lock bool) *win32.IUnknownVtbl {
	if lock {
		MuVtbl.Lock()
		defer MuVtbl.Unlock()
	}
	if _pIUnknownVtbl != nil {
		return _pIUnknownVtbl
	}
	_pIUnknownVtbl = (*win32.IUnknownVtbl)(Malloc(unsafe.Sizeof(*_pIUnknownVtbl)))
	*_pIUnknownVtbl = win32.IUnknownVtbl{
		QueryInterface: syscall.NewCallback((*IUnknownComObj).QueryInterface),
		AddRef:         syscall.NewCallback((*IUnknownComObj).AddRef),
		Release:        syscall.NewCallback((*IUnknownComObj).Release),
	}
	return _pIUnknownVtbl
}

func (this *IUnknownComObj) GetVtbl() *win32.IUnknownVtbl {
	return this.BuildVtbl(true)
}
