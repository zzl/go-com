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

const ComObjSize = unsafe.Sizeof(IUnknownComObj{})

func init() {
	hHeap, _ = win32.GetProcessHeap()
}

type ComObjConstraint[T any] interface {
	GetVtbl() *win32.IUnknownVtbl
	AddRef() uintptr
	Release() uintptr
	Pointer() unsafe.Pointer
	*T
}

type RealObjectAware interface {
	SetRealObject(obj interface{})
}

type Initializable interface {
	Initialize()
}

func NewComObj[T any, PT ComObjConstraint[T]](impl win32.IUnknownInterface) *T {
	if roa, ok := impl.(RealObjectAware); ok {
		roa.SetRealObject(impl)
	}

	size := unsafe.Sizeof(*(*T)(nil))
	p := Malloc(size)
	pComObj := (*T)(p)

	comObj := PT(pComObj)
	pUnknownComObj := (*IUnknownComObj)(p)
	pUnknownComObj.LpVtbl = comObj.GetVtbl()
	pUnknownComObj.implSlot = AddImpl(impl)
	pUnknownComObj.cRef = 1

	//
	impl.(comObjAware).SetComObj(comObj)

	if initializable, ok := any(pComObj).(Initializable); ok {
		initializable.Initialize()
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
	Pointer() unsafe.Pointer
}

type comObjAware interface {
	SetComObj(obj ComObjInterface)
}

//
type IUnknownImpl struct {
	ComObject ComObjInterface
}

func (this *IUnknownImpl) SetComObj(obj ComObjInterface) {
	this.ComObject = obj
}

func (this *IUnknownImpl) AssignPpvObject(ppvObject unsafe.Pointer) {
	*(*unsafe.Pointer)(ppvObject) = this.ComObject.Pointer()
}

func (this *IUnknownImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IUnknown {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return win32.E_NOINTERFACE
}

func (this *IUnknownImpl) AddRef() uint32 {
	return uint32(this.ComObject.AddRef())
}

func (this *IUnknownImpl) Release() uint32 {
	return uint32(this.ComObject.Release())
}

//
type IUnknownComObj struct {
	LpVtbl *win32.IUnknownVtbl
	cRef   int32
	//Impl   win32.IUnknownInterface
	implSlot int32
}

func (this *IUnknownComObj) Impl() win32.IUnknownInterface {
	return impls[this.implSlot] //GetImpl(this.implSlot)
}

func (this *IUnknownComObj) free() {
	impls[this.implSlot] = nil
	freeImplSlots = append(freeImplSlots, this.implSlot)
	bOk, err := win32.HeapFree(hHeap, 0, unsafe.Pointer(this))
	if err != win32.NO_ERROR || bOk != win32.TRUE {
		println("??")
	}
}

func (this *IUnknownComObj) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) uintptr {
	return (uintptr)(this.Impl().QueryInterface(riid, ppvObject))
}

func (this *IUnknownComObj) AddRef() uintptr {
	cRef := atomic.AddInt32(&this.cRef, 1)
	return uintptr(cRef)
}

func (this *IUnknownComObj) Release() uintptr {
	cRef := atomic.AddInt32(&this.cRef, -1)
	if cRef == 0 {
		this.free()
	}
	return uintptr(cRef)
}

func (this *IUnknownComObj) IUnknown() *win32.IUnknown {
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

func (this *IUnknownComObj) Pointer() unsafe.Pointer {
	return unsafe.Pointer(this)
}

//
type SubComObj struct {
	LpVtbl *win32.IUnknownVtbl
	Parent *IUnknownComObj
}

func (this *SubComObj) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) uintptr {
	return this.Parent.QueryInterface(riid, ppvObject)
}

func (this *SubComObj) AddRef() uintptr {
	return this.Parent.AddRef()
}

func (this *SubComObj) Release() uintptr {
	return this.Parent.Release()
}

var _pSubIUnknownVtbl *win32.IUnknownVtbl

func (this *SubComObj) BuildVtbl(lock bool) *win32.IUnknownVtbl {
	if lock {
		MuVtbl.Lock()
		defer MuVtbl.Unlock()
	}
	if _pSubIUnknownVtbl != nil {
		return _pSubIUnknownVtbl
	}
	_pSubIUnknownVtbl = (*win32.IUnknownVtbl)(Malloc(unsafe.Sizeof(*_pIUnknownVtbl)))
	*_pSubIUnknownVtbl = win32.IUnknownVtbl{
		QueryInterface: syscall.NewCallback((*SubComObj).QueryInterface),
		AddRef:         syscall.NewCallback((*SubComObj).AddRef),
		Release:        syscall.NewCallback((*SubComObj).Release),
	}
	return _pSubIUnknownVtbl
}

func (this *SubComObj) Pointer() unsafe.Pointer {
	return unsafe.Pointer(this)
}
