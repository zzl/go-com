package com

import (
	"github.com/zzl/go-win32api/win32"
	"log"
	"runtime"
	"sync"
	"sync/atomic"
	"unsafe"
)

var tlsIndex uint32

type Context struct {
	ID           int32 //could be reused
	TID          uint32
	CurrentScope *Scope
	LastError    *ErrorInfo
}

var contexts []*Context
var muContext sync.Mutex

var theOnlyContextPtr unsafe.Pointer

func init() {
	var errno win32.WIN32_ERROR
	tlsIndex, errno = win32.TlsAlloc()
	if errno != win32.NO_ERROR {
		log.Fatal("?")
	}
}

func InitializeContext() {
	//win32.TlsSetValue()
	index := -1
	context := &Context{}
	context.TID = win32.GetCurrentThreadId()

	muContext.Lock()
	defer muContext.Unlock()
	for n, ctx := range contexts {
		if ctx == nil {
			contexts[n] = context
			index = n
			break
		}
	}
	count := len(contexts)
	if index == -1 {
		index = count
		contexts = append(contexts, context)
	}
	context.ID = int32(index + 1)
	win32.TlsSetValue(tlsIndex, unsafe.Pointer(uintptr(index)))

	//
	compactContexts()
}

func GetContext() *Context {
	pContext := atomic.LoadPointer(&theOnlyContextPtr)
	if pContext != nil {
		//println("GOT ONLY CONTEXT")
		return (*Context)(pContext)
	}

	index := int(win32.TlsGetValueAlt(tlsIndex))
	return contexts[index]
}

func UninitializeContext() {
	muContext.Lock()
	defer muContext.Unlock()
	index := int(win32.TlsGetValueAlt(tlsIndex))
	contexts[index] = nil

	//
	compactContexts()
}

func compactContexts() {
	nonNilIndex := -1
	nonNilCount := 0
	for n, context := range contexts {
		if context != nil {
			nonNilCount++
			nonNilIndex = n
		}
	}
	if nonNilCount == 1 {
		ptrContext := unsafe.Pointer(contexts[nonNilIndex])
		atomic.StorePointer(&theOnlyContextPtr, ptrContext)
	} else {
		atomic.StorePointer(&theOnlyContextPtr, nil)
	}
	contexts = contexts[:nonNilIndex+1]
}

type Initialized struct {
}

func (me Initialized) Uninitialize() {
	Uninitialize()
}

func Initialize() Initialized {
	runtime.LockOSThread()

	InitializeContext()
	win32.CoInitialize(nil)

	//atomic.LoadInt32()
	//tId := win32.GetCurrentThreadId()
	//comThreadIds.Store(tId, true)
	return Initialized{}
}

func Uninitialize() {
	win32.CoUninitialize()
	UninitializeContext()
	runtime.UnlockOSThread()
}

func EnsureThreadCoInitialized() {
	//tId := win32.GetCurrentThreadId()
	//if _, loaded := comThreadIds.LoadOrStore(tId, true); !loaded {
	//	win32.CoInitialize(nil)
	//}
}
