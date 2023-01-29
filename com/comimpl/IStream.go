package comimpl

import (
	"io"
	"syscall"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
)

//
type IStreamImpl struct {
	ISequentialStreamImpl
}

func (this *IStreamImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IStream {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.ISequentialStreamImpl.QueryInterface(riid, ppvObject)
}

func (this *IStreamImpl) Seek(dlibMove int64, dwOrigin win32.STREAM_SEEK, plibNewPosition *uint64) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IStreamImpl) SetSize(libNewSize uint64) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IStreamImpl) CopyTo(pstm *win32.IStream, cb uint64, pcbRead *uint64, pcbWritten *uint64) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IStreamImpl) Commit(grfCommitFlags win32.STGC) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IStreamImpl) Revert() win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IStreamImpl) LockRegion(libOffset uint64, cb uint64, dwLockType win32.LOCKTYPE) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IStreamImpl) UnlockRegion(libOffset uint64, cb uint64, dwLockType uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IStreamImpl) Stat(pstatstg *win32.STATSTG, grfStatFlag win32.STATFLAG) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IStreamImpl) Clone(ppstm **win32.IStream) win32.HRESULT {
	return win32.E_NOTIMPL
}

//
type IStreamComObj struct {
	ISequentialStreamComObj
}

func (this *IStreamComObj) GetIUnknownComObj() *com.IUnknownComObj {
	return &this.IUnknownComObj
}

func (this *IStreamComObj) impl() win32.IStreamInterface {
	return this.Impl().(win32.IStreamInterface)
}

func (this *IStreamComObj) IStream() *win32.IStream {
	return (*win32.IStream)(unsafe.Pointer(this))
}

func (this *IStreamComObj) Seek(dlibMove int64, dwOrigin win32.STREAM_SEEK, plibNewPosition *uint64) uintptr {
	return (uintptr)(this.impl().Seek(dlibMove, dwOrigin, plibNewPosition))
}

func (this *IStreamComObj) SetSize(libNewSize uint64) uintptr {
	return (uintptr)(this.impl().SetSize(libNewSize))
}

func (this *IStreamComObj) CopyTo(pstm *win32.IStream, cb uint64, pcbRead *uint64, pcbWritten *uint64) uintptr {
	return (uintptr)(this.impl().CopyTo(pstm, cb, pcbRead, pcbWritten))
}

func (this *IStreamComObj) Commit(grfCommitFlags win32.STGC) uintptr {
	return (uintptr)(this.impl().Commit(grfCommitFlags))
}

func (this *IStreamComObj) Revert() uintptr {
	return (uintptr)(this.impl().Revert())
}

func (this *IStreamComObj) LockRegion(libOffset uint64, cb uint64, dwLockType win32.LOCKTYPE) uintptr {
	return (uintptr)(this.impl().LockRegion(libOffset, cb, dwLockType))
}

func (this *IStreamComObj) UnlockRegion(libOffset uint64, cb uint64, dwLockType uint32) uintptr {
	return (uintptr)(this.impl().UnlockRegion(libOffset, cb, dwLockType))
}

func (this *IStreamComObj) Stat(pstatstg *win32.STATSTG, grfStatFlag win32.STATFLAG) uintptr {
	return (uintptr)(this.impl().Stat(pstatstg, grfStatFlag))
}

func (this *IStreamComObj) Clone(ppstm **win32.IStream) uintptr {
	return (uintptr)(this.impl().Clone(ppstm))
}

var _IStreamVtbl *win32.IStreamVtbl

func (this *IStreamComObj) BuildVtbl(lock bool) *win32.IStreamVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _IStreamVtbl != nil {
		return _IStreamVtbl
	}
	_IStreamVtbl = (*win32.IStreamVtbl)(com.Malloc(unsafe.Sizeof(*_IStreamVtbl)))
	*_IStreamVtbl = win32.IStreamVtbl{
		ISequentialStreamVtbl: *this.ISequentialStreamComObj.BuildVtbl(false),
		Seek:                  syscall.NewCallback((*IStreamComObj).Seek),
		SetSize:               syscall.NewCallback((*IStreamComObj).SetSize),
		CopyTo:                syscall.NewCallback((*IStreamComObj).CopyTo),
		Commit:                syscall.NewCallback((*IStreamComObj).Commit),
		Revert:                syscall.NewCallback((*IStreamComObj).Revert),
		LockRegion:            syscall.NewCallback((*IStreamComObj).LockRegion),
		UnlockRegion:          syscall.NewCallback((*IStreamComObj).UnlockRegion),
		Stat:                  syscall.NewCallback((*IStreamComObj).Stat),
		Clone:                 syscall.NewCallback((*IStreamComObj).Clone),
	}
	return _IStreamVtbl
}

func (this *IStreamComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

func NewIStreamComObj(impl win32.IStreamInterface) *IStreamComObj {
	comObj := com.NewComObj[IStreamComObj](impl)
	return comObj
}

func NewIStream(impl win32.IStreamInterface) *win32.IStream {
	return NewIStreamComObj(impl).IStream()
}

func NewReaderWriterIStream(reader io.Reader, writer io.Writer) *win32.IStream {
	impl := NewReaderWriterIStreamImpl(reader, writer)
	return NewIStreamComObj(impl).IStream()
}

//
type ReaderWriterIStreamImpl struct {
	IStreamImpl

	reader io.Reader
	writer io.Writer
}

func NewReaderWriterIStreamImpl(reader io.Reader, writer io.Writer) *ReaderWriterIStreamImpl {
	obj := &ReaderWriterIStreamImpl{reader: reader, writer: writer}
	return obj
}

func (this *ReaderWriterIStreamImpl) Read(pv unsafe.Pointer, cb uint32, pcbRead *uint32) win32.HRESULT {
	if this.reader == nil {
		return win32.E_NOTIMPL
	}
	bts := unsafe.Slice((*byte)(pv), cb)
	cbRead, err := this.reader.Read(bts)
	if err != nil {
		return win32.E_UNEXPECTED
	} else {
		*pcbRead = uint32(cbRead)
	}
	return win32.S_OK
}

func (this *ReaderWriterIStreamImpl) Write(pv unsafe.Pointer, cb uint32, pcbWritten *uint32) win32.HRESULT {
	if this.writer == nil {
		return win32.E_NOTIMPL
	}
	bts := unsafe.Slice((*byte)(pv), cb)
	this.writer.Write(bts)
	*pcbWritten = cb
	return win32.S_OK
}

func (this *ReaderWriterIStreamImpl) Seek(dlibMove int64,
	dwOrigin win32.STREAM_SEEK, plibNewPosition *uint64) win32.HRESULT {
	seeker, ok := this.reader.(io.ReadSeeker)
	if !ok {
		return win32.E_NOTIMPL
	}
	newPos, err := seeker.Seek(dlibMove, int(dwOrigin))
	if err != nil {
		return win32.E_FAIL
	}
	if plibNewPosition != nil {
		*plibNewPosition = uint64(newPos)
	}
	return win32.S_OK
}

func (this *ReaderWriterIStreamImpl) Commit(grfCommitFlags win32.STGC) win32.HRESULT {
	type Flusher interface {
		Flush() error
	}
	if flusher, ok := this.writer.(Flusher); ok {
		flusher.Flush()
	}
	return win32.S_OK
}
