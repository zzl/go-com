package comimpl

import (
	"syscall"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
)

type IDataObjectImpl struct {
	com.IUnknownImpl

	Text       string
	DataHandle uintptr
	shellDo    *win32.IDataObject
}

func (this *IDataObjectImpl) Dispose() {
	do := this.shellDo
	if do != nil {
		this.shellDo = nil
		do.Release()
	}
}

func (this *IDataObjectImpl) lazyLoadShellDo() win32.HRESULT {
	if this.shellDo != nil {
		return win32.S_OK
	}
	var pdo *win32.IDataObject
	hr := win32.SHCreateDataObject(nil, 0, nil, nil,
		&win32.IID_IDataObject, unsafe.Pointer(&pdo))
	if win32.SUCCEEDED(hr) {
		this.shellDo = pdo
	}
	return hr
}

func (this *IDataObjectImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IDataObject {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IUnknownImpl.QueryInterface(riid, ppvObject)
}

func (this *IDataObjectImpl) GetData(pformatetcIn *win32.FORMATETC, pmedium *win32.STGMEDIUM) win32.HRESULT {
	*pmedium = win32.STGMEDIUM{}
	hr := win32.DATA_E_FORMATETC

	if pformatetcIn.CfFormat >= uint16(win32.CF_PRIVATEFIRST) &&
		pformatetcIn.CfFormat <= uint16(win32.CF_PRIVATELAST) {
		*pmedium.HGlobal() = this.DataHandle
		return win32.S_OK
	}

	if pformatetcIn.CfFormat == uint16(win32.CF_UNICODETEXT) {
		if pformatetcIn.Tymed&(uint32)(win32.TYMED_HGLOBAL) != 0 {
			text := this.Text
			wsz, _ := syscall.UTF16FromString(text)
			cb := len(wsz) * 2
			h, _ := win32.GlobalAlloc(win32.GPTR, uintptr(cb))
			if h == 0 {
				hr = win32.E_OUTOFMEMORY
			} else {
				hr = win32.S_OK
			}
			if win32.SUCCEEDED(hr) {
				bts := unsafe.Slice((*uint16)(unsafe.Pointer(h)), cb)
				copy(bts, wsz)
				*pmedium.HGlobal() = h
				pmedium.Tymed = (uint32)(win32.TYMED_HGLOBAL)
			}
		}
	} else if win32.SUCCEEDED(this.lazyLoadShellDo()) {
		hr = this.shellDo.GetData(pformatetcIn, pmedium)
	}
	return hr
}

func (this *IDataObjectImpl) GetDataHere(pformatetc *win32.FORMATETC, pmedium *win32.STGMEDIUM) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IDataObjectImpl) QueryGetData(pformatetc *win32.FORMATETC) win32.HRESULT {
	hr := win32.S_FALSE
	if pformatetc.CfFormat >= uint16(win32.CF_PRIVATEFIRST) &&
		pformatetc.CfFormat <= uint16(win32.CF_PRIVATELAST) {
		hr = win32.S_OK
	} else if pformatetc.CfFormat == uint16(win32.CF_UNICODETEXT) {
		hr = win32.S_OK
	} else if win32.SUCCEEDED(this.lazyLoadShellDo()) {
		hr = this.shellDo.QueryGetData(pformatetc)
	}
	return hr
}

func (this *IDataObjectImpl) GetCanonicalFormatEtc(pformatectIn *win32.FORMATETC, pformatetcOut *win32.FORMATETC) win32.HRESULT {
	*pformatetcOut = *pformatectIn
	pformatetcOut.Ptd = nil
	return win32.DATA_S_SAMEFORMATETC
}

func (this *IDataObjectImpl) SetData(pformatetc *win32.FORMATETC, pmedium *win32.STGMEDIUM, fRelease win32.BOOL) win32.HRESULT {
	if pformatetc.CfFormat >= uint16(win32.CF_PRIVATEFIRST) &&
		pformatetc.CfFormat <= uint16(win32.CF_PRIVATELAST) {
		this.DataHandle = *pmedium.HGlobal()
		return win32.S_OK
	}
	hr := this.lazyLoadShellDo()
	if win32.SUCCEEDED(hr) {
		hr = this.shellDo.SetData(pformatetc, pmedium, fRelease)
	}
	return hr
}

func (this *IDataObjectImpl) EnumFormatEtc(dwDirection uint32, ppenumFormatEtc **win32.IEnumFORMATETC) win32.HRESULT {
	*ppenumFormatEtc = nil
	hr := win32.E_NOTIMPL
	if dwDirection == (uint32)(win32.DATADIR_GET) {
		rgfmtetc := []win32.FORMATETC{
			{uint16(win32.CF_UNICODETEXT), nil, 0, 0, (uint32)(win32.TYMED_HGLOBAL)},
		}
		hr = win32.SHCreateStdEnumFmtEtc(1, &rgfmtetc[0], ppenumFormatEtc)
	}
	return hr
}

func (this *IDataObjectImpl) DAdvise(pformatetc *win32.FORMATETC, advf uint32, pAdvSink *win32.IAdviseSink, pdwConnection *uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IDataObjectImpl) DUnadvise(dwConnection uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IDataObjectImpl) EnumDAdvise(ppenumAdvise **win32.IEnumSTATDATA) win32.HRESULT {
	return win32.E_NOTIMPL
}

//
type IDataObjectComObj struct {
	com.IUnknownComObj
	impl win32.IDataObjectInterface
}

func (this *IDataObjectComObj) GetData(pformatetcIn *win32.FORMATETC, pmedium *win32.STGMEDIUM) uintptr {
	return uintptr(this.impl.GetData(pformatetcIn, pmedium))
}

func (this *IDataObjectComObj) GetDataHere(pformatetc *win32.FORMATETC, pmedium *win32.STGMEDIUM) uintptr {
	return uintptr(this.impl.GetDataHere(pformatetc, pmedium))
}

func (this *IDataObjectComObj) QueryGetData(pformatetc *win32.FORMATETC) uintptr {
	return uintptr(this.impl.QueryGetData(pformatetc))
}

func (this *IDataObjectComObj) GetCanonicalFormatEtc(pformatectIn *win32.FORMATETC, pformatetcOut *win32.FORMATETC) uintptr {
	return uintptr(this.impl.GetCanonicalFormatEtc(pformatectIn, pformatetcOut))
}

func (this *IDataObjectComObj) SetData(pformatetc *win32.FORMATETC, pmedium *win32.STGMEDIUM, fRelease win32.BOOL) uintptr {
	return uintptr(this.impl.SetData(pformatetc, pmedium, fRelease))
}

func (this *IDataObjectComObj) EnumFormatEtc(dwDirection uint32, ppenumFormatEtc **win32.IEnumFORMATETC) uintptr {
	return uintptr(this.impl.EnumFormatEtc(dwDirection, ppenumFormatEtc))
}

func (this *IDataObjectComObj) DAdvise(pformatetc *win32.FORMATETC, advf uint32, pAdvSink *win32.IAdviseSink, pdwConnection *uint32) uintptr {
	return uintptr(this.impl.DAdvise(pformatetc, advf, pAdvSink, pdwConnection))
}

func (this *IDataObjectComObj) DUnadvise(dwConnection uint32) uintptr {
	return uintptr(this.impl.DUnadvise(dwConnection))
}

func (this *IDataObjectComObj) EnumDAdvise(ppenumAdvise **win32.IEnumSTATDATA) uintptr {
	return uintptr(this.impl.EnumDAdvise(ppenumAdvise))
}

var _pIDataObjectVtbl *win32.IDataObjectVtbl

func (this *IDataObjectComObj) BuildVtbl(lock bool) *win32.IDataObjectVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIDataObjectVtbl != nil {
		return _pIDataObjectVtbl
	}
	_pIDataObjectVtbl = &win32.IDataObjectVtbl{
		IUnknownVtbl:          *this.IUnknownComObj.BuildVtbl(false),
		GetData:               syscall.NewCallback((*IDataObjectComObj).GetData),
		GetDataHere:           syscall.NewCallback((*IDataObjectComObj).GetDataHere),
		QueryGetData:          syscall.NewCallback((*IDataObjectComObj).QueryGetData),
		GetCanonicalFormatEtc: syscall.NewCallback((*IDataObjectComObj).GetCanonicalFormatEtc),
		SetData:               syscall.NewCallback((*IDataObjectComObj).SetData),
		EnumFormatEtc:         syscall.NewCallback((*IDataObjectComObj).EnumFormatEtc),
		DAdvise:               syscall.NewCallback((*IDataObjectComObj).DAdvise),
		DUnadvise:             syscall.NewCallback((*IDataObjectComObj).DUnadvise),
		EnumDAdvise:           syscall.NewCallback((*IDataObjectComObj).EnumDAdvise),
	}
	return _pIDataObjectVtbl
}

func NewIDataObjectComObj(impl win32.IDataObjectInterface) *IDataObjectComObj {
	comObj := com.NewComObj[IDataObjectComObj](impl)
	return comObj
}
