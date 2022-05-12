package main

import (
	"syscall"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
)

type IFileDialogEventsInterfaceImpl struct {
	com.IUnknownImpl
}

func (this *IFileDialogEventsInterfaceImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == win32.IID_IFileDialogEvents {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	if *riid == win32.IID_IFileDialogControlEvents {
		*(*unsafe.Pointer)(ppvObject) = this.ComObject.(*IFileDialogEventsComObj).controlEventsComObj.Pointer()
		this.AddRef()
		return win32.S_OK
	}
	return this.IUnknownImpl.QueryInterface(riid, ppvObject)
}

// IFileDialogEvents
func (this *IFileDialogEventsInterfaceImpl) OnFileOk(pfd *win32.IFileDialog) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IFileDialogEventsInterfaceImpl) OnFolderChanging(pfd *win32.IFileDialog, psiFolder *win32.IShellItem) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IFileDialogEventsInterfaceImpl) OnFolderChange(pfd *win32.IFileDialog) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IFileDialogEventsInterfaceImpl) OnSelectionChange(pfd *win32.IFileDialog) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IFileDialogEventsInterfaceImpl) OnShareViolation(pfd *win32.IFileDialog, psi *win32.IShellItem, pResponse *win32.FDE_SHAREVIOLATION_RESPONSE) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IFileDialogEventsInterfaceImpl) OnTypeChange(pfd *win32.IFileDialog) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IFileDialogEventsInterfaceImpl) OnOverwrite(pfd *win32.IFileDialog, psi *win32.IShellItem, pResponse *win32.FDE_OVERWRITE_RESPONSE) win32.HRESULT {
	return win32.E_NOTIMPL
}

// IFileDialogControlEvents
func (this *IFileDialogEventsInterfaceImpl) OnItemSelected(pfdc *win32.IFileDialogCustomize, dwIDCtl uint32, dwIDItem uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IFileDialogEventsInterfaceImpl) OnButtonClicked(pfdc *win32.IFileDialogCustomize, dwIDCtl uint32) win32.HRESULT {
	win32.MessageBox(win32.GetForegroundWindow(), win32.StrToPwstr("My Button clicked!"),
		win32.StrToPwstr("IFileDialogControlEvents.OnButtonClicked"), win32.MB_ICONINFORMATION)
	return win32.S_OK
}

func (this *IFileDialogEventsInterfaceImpl) OnCheckButtonToggled(pfdc *win32.IFileDialogCustomize, dwIDCtl uint32, bChecked win32.BOOL) win32.HRESULT {
	return win32.E_NOTIMPL
}

func (this *IFileDialogEventsInterfaceImpl) OnControlActivating(pfdc *win32.IFileDialogCustomize, dwIDCtl uint32) win32.HRESULT {
	return win32.E_NOTIMPL
}

//
type IFileDialogEventsComObj struct {
	com.IUnknownComObj
	//
	controlEventsComObj IFileDialogControlEventsComObj
}

func (this *IFileDialogEventsComObj) Initialize() {
	this.controlEventsComObj.LpVtbl = this.controlEventsComObj.GetVtbl()
	this.controlEventsComObj.Parent = &this.IUnknownComObj
}

func (this *IFileDialogEventsComObj) impl() win32.IFileDialogEventsInterface {
	return this.Impl().(win32.IFileDialogEventsInterface)
}

func (this *IFileDialogEventsComObj) IFileDialogEvents() *win32.IFileDialogEvents {
	return (*win32.IFileDialogEvents)(unsafe.Pointer(this))
}

func (this *IFileDialogEventsComObj) OnFileOk(pfd *win32.IFileDialog) uintptr {
	return (uintptr)(this.impl().OnFileOk(pfd))
}

func (this *IFileDialogEventsComObj) OnFolderChanging(pfd *win32.IFileDialog, psiFolder *win32.IShellItem) uintptr {
	return (uintptr)(this.impl().OnFolderChanging(pfd, psiFolder))
}

func (this *IFileDialogEventsComObj) OnFolderChange(pfd *win32.IFileDialog) uintptr {
	return (uintptr)(this.impl().OnFolderChange(pfd))
}

func (this *IFileDialogEventsComObj) OnSelectionChange(pfd *win32.IFileDialog) uintptr {
	return (uintptr)(this.impl().OnSelectionChange(pfd))
}

func (this *IFileDialogEventsComObj) OnShareViolation(pfd *win32.IFileDialog, psi *win32.IShellItem, pResponse *win32.FDE_SHAREVIOLATION_RESPONSE) uintptr {
	return (uintptr)(this.impl().OnShareViolation(pfd, psi, pResponse))
}

func (this *IFileDialogEventsComObj) OnTypeChange(pfd *win32.IFileDialog) uintptr {
	return (uintptr)(this.impl().OnTypeChange(pfd))
}

func (this *IFileDialogEventsComObj) OnOverwrite(pfd *win32.IFileDialog, psi *win32.IShellItem, pResponse *win32.FDE_OVERWRITE_RESPONSE) uintptr {
	return (uintptr)(this.impl().OnOverwrite(pfd, psi, pResponse))
}

//
var _pIFileDialogEventsVtbl *win32.IFileDialogEventsVtbl

func (this *IFileDialogEventsComObj) BuildVtbl(lock bool) *win32.IFileDialogEventsVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIFileDialogEventsVtbl != nil {
		return _pIFileDialogEventsVtbl
	}
	_pIFileDialogEventsVtbl = (*win32.IFileDialogEventsVtbl)(
		com.Malloc(unsafe.Sizeof(*_pIFileDialogEventsVtbl)))

	*_pIFileDialogEventsVtbl = win32.IFileDialogEventsVtbl{
		IUnknownVtbl:      *this.IUnknownComObj.BuildVtbl(false),
		OnFileOk:          syscall.NewCallback((*IFileDialogEventsComObj).OnFileOk),
		OnFolderChanging:  syscall.NewCallback((*IFileDialogEventsComObj).OnFolderChanging),
		OnFolderChange:    syscall.NewCallback((*IFileDialogEventsComObj).OnFolderChange),
		OnSelectionChange: syscall.NewCallback((*IFileDialogEventsComObj).OnSelectionChange),
		OnShareViolation:  syscall.NewCallback((*IFileDialogEventsComObj).OnShareViolation),
		OnTypeChange:      syscall.NewCallback((*IFileDialogEventsComObj).OnTypeChange),
		OnOverwrite:       syscall.NewCallback((*IFileDialogEventsComObj).OnOverwrite),
	}
	return _pIFileDialogEventsVtbl
}

func (this *IFileDialogEventsComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

//
type IFileDialogControlEventsComObj struct {
	com.SubComObj
}

func (this *IFileDialogControlEventsComObj) IFileDialogControlEvents() *win32.IFileDialogControlEvents {
	return (*win32.IFileDialogControlEvents)(unsafe.Pointer(this))
}

func (this *IFileDialogControlEventsComObj) impl() win32.IFileDialogControlEventsInterface {
	pParent := (*IFileDialogEventsComObj)(unsafe.Pointer(this.Parent))
	return pParent.Impl().(win32.IFileDialogControlEventsInterface)
}

func (this *IFileDialogControlEventsComObj) OnItemSelected(pfdc *win32.IFileDialogCustomize, dwIDCtl uint32, dwIDItem uint32) uintptr {
	return uintptr(this.impl().OnItemSelected(pfdc, dwIDCtl, dwIDItem))
}

func (this *IFileDialogControlEventsComObj) OnButtonClicked(pfdc *win32.IFileDialogCustomize, dwIDCtl uint32) uintptr {
	return (uintptr)(this.impl().OnButtonClicked(pfdc, dwIDCtl))
}

func (this *IFileDialogControlEventsComObj) OnCheckButtonToggled(pfdc *win32.IFileDialogCustomize, dwIDCtl uint32, bChecked win32.BOOL) uintptr {
	return (uintptr)(this.impl().OnCheckButtonToggled(pfdc, dwIDCtl, bChecked))
}

func (this *IFileDialogControlEventsComObj) OnControlActivating(pfdc *win32.IFileDialogCustomize, dwIDCtl uint32) uintptr {
	return (uintptr)(this.impl().OnControlActivating(pfdc, dwIDCtl))
}

//
var _pIFileDialogControlEventsVtbl *win32.IFileDialogControlEventsVtbl

func (this *IFileDialogControlEventsComObj) BuildVtbl(lock bool) *win32.IFileDialogControlEventsVtbl {
	if lock {
		com.MuVtbl.Lock()
		defer com.MuVtbl.Unlock()
	}
	if _pIFileDialogControlEventsVtbl != nil {
		return _pIFileDialogControlEventsVtbl
	}
	_pIFileDialogControlEventsVtbl = (*win32.IFileDialogControlEventsVtbl)(
		com.Malloc(unsafe.Sizeof(*_pIFileDialogControlEventsVtbl)))

	*_pIFileDialogControlEventsVtbl = win32.IFileDialogControlEventsVtbl{
		IUnknownVtbl:         *this.SubComObj.BuildVtbl(false),
		OnItemSelected:       syscall.NewCallback((*IFileDialogControlEventsComObj).OnItemSelected),
		OnButtonClicked:      syscall.NewCallback((*IFileDialogControlEventsComObj).OnButtonClicked),
		OnCheckButtonToggled: syscall.NewCallback((*IFileDialogControlEventsComObj).OnCheckButtonToggled),
		OnControlActivating:  syscall.NewCallback((*IFileDialogControlEventsComObj).OnControlActivating),
	}
	return _pIFileDialogControlEventsVtbl
}

func (this *IFileDialogControlEventsComObj) GetVtbl() *win32.IUnknownVtbl {
	return &this.BuildVtbl(true).IUnknownVtbl
}

/* https://docs.microsoft.com/en-us/windows/win32/shell/common-file-dialog#basic-usage */
//
func main() {

	com.Initialize()

	var pfd *win32.IFileDialog

	var CLSID_FileOpenDialog = syscall.GUID{0xdc1c5a9c, 0xe88a, 0x4dde,
		[8]byte{0xa5, 0xa1, 0x60, 0xf8, 0x2a, 0x20, 0xae, 0xf7}}

	hr := win32.CoCreateInstance(&CLSID_FileOpenDialog, nil, win32.CLSCTX_INPROC_SERVER,
		&win32.IID_IFileDialog, unsafe.Pointer(&pfd))
	win32.ASSERT_SUCCEEDED(hr)

	//
	var pfdc *win32.IFileDialogCustomize
	hr = pfd.QueryInterface(&win32.IID_IFileDialogCustomize, unsafe.Pointer(&pfdc))
	if win32.SUCCEEDED(hr) {
		pfdc.AddPushButton(1001, win32.StrToPwstr("My Button"))
		pfdc.MakeProminent(1001)
		pfdc.Release()
	}

	//
	var pfde *win32.IFileDialogEvents
	pfde = com.NewComObj[IFileDialogEventsComObj](&IFileDialogEventsInterfaceImpl{}).IFileDialogEvents()
	var cookie uint32
	hr = pfd.Advise(pfde, &cookie)
	win32.ASSERT_SUCCEEDED(hr)
	pfde.Release()

	//
	var fos win32.FILEOPENDIALOGOPTIONS
	hr = pfd.GetOptions(&fos)
	win32.ASSERT_SUCCEEDED(hr)

	hr = pfd.SetOptions(fos | win32.FOS_FORCEFILESYSTEM)
	win32.ASSERT_SUCCEEDED(hr)

	fileTypes := []win32.COMDLG_FILTERSPEC{
		{PszName: win32.StrToPwstr("Word Document (*.doc)"), PszSpec: win32.StrToPwstr("*.doc")},
		{PszName: win32.StrToPwstr("Web Page (*.htm; *.html)"), PszSpec: win32.StrToPwstr("*.htm;*.html")},
		{PszName: win32.StrToPwstr("Text Document (*.txt)"), PszSpec: win32.StrToPwstr(".txt")},
		{PszName: win32.StrToPwstr("All Documents (*.*)"), PszSpec: win32.StrToPwstr("*.*")},
	}

	hr = pfd.SetFileTypes(uint32(len(fileTypes)), &fileTypes[0])
	win32.ASSERT_SUCCEEDED(hr)

	const INDEX_WORDDOC = 1
	hr = pfd.SetFileTypeIndex(INDEX_WORDDOC)
	win32.ASSERT_SUCCEEDED(hr)

	hr = pfd.SetDefaultExtension(win32.StrToPwstr("doc;docx"))
	win32.ASSERT_SUCCEEDED(hr)

	hr = pfd.Show(0)
	if hr == win32.S_OK {
		var psiResult *win32.IShellItem
		hr = pfd.GetResult(&psiResult)
		win32.ASSERT_SUCCEEDED(hr)

		var pszFilePath win32.PWSTR
		hr = psiResult.GetDisplayName(win32.SIGDN_FILESYSPATH, &pszFilePath)
		win32.ASSERT_SUCCEEDED(hr)

		//
		win32.MessageBox(0, pszFilePath, win32.StrToPwstr("File selected"), win32.MB_ICONINFORMATION)

		win32.CoTaskMemFree(unsafe.Pointer(pszFilePath))
		psiResult.Release()
	} else if hr == com.HresultFromWin32(win32.ERROR_CANCELLED) {
		println("User canceled")
	} else {
		println(win32.HRESULT_ToString(hr))
	}
	//
	pfd.Unadvise(cookie)
	pfd.Release()
}
