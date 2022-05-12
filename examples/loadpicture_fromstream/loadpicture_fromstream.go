package main

import (
	"log"
	"os"
	"unsafe"

	"github.com/zzl/go-com/com/comimpl"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/win32"
)

func main() {

	ole.Initialize()

	imgPath := `C:\Windows\Web\Wallpaper\Windows\img0.jpg`
	f, _ := os.Open(imgPath)
	pStream := comimpl.NewReaderWriterIStream(f, nil)
	var pPicture *win32.IPicture
	hr := win32.OleLoadPicture(pStream, 0, win32.FALSE,
		&win32.IID_IPicture, unsafe.Pointer(&pPicture))

	if win32.FAILED(hr) {
		log.Fatal(win32.HRESULT_ToString(hr))
	}

	var width, height int32

	pPicture.Get_Width(&width)
	pPicture.Get_Height(&height)

	width, height = ole.HiMetricToPixel(width, height)

	println("Width:", width)
	println("Height:", height)

	pStream.Release()
	pPicture.Release()
	println("done.")

	ole.Uninitialize()

}
