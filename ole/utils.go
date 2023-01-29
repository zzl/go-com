package ole

import (
	"unsafe"

	"github.com/zzl/go-win32api/v2/win32"
)

func ProcessInvokeArgs(pDispParams *win32.DISPPARAMS, formalArgc int) ([]*Variant, []int) {
	if formalArgc == 0 {
		return nil, nil
	}
	result := make([]*Variant, formalArgc)
	srcArgIndexes := make([]int, formalArgc)
	argc := int(pDispParams.CArgs)
	argv := unsafe.Slice(pDispParams.Rgvarg, argc)
	namedArgc := int(pDispParams.CNamedArgs)
	if namedArgc == 1 && *pDispParams.RgdispidNamedArgs == win32.DISPID_PROPERTYPUT {
		namedArgc = 0
	}
	if namedArgc != 0 {
		namedArgIdxs := unsafe.Slice(pDispParams.RgdispidNamedArgs, namedArgc)
		for n, idx := range namedArgIdxs {
			if int(idx) < formalArgc {
				result[idx] = (*Variant)(&argv[n])
				srcArgIndexes[idx] = n
			}
		}
	}
	var resultIdx = 0
	for n := argc - 1; n >= namedArgc; n-- {
		if resultIdx == formalArgc {
			break
		}
		result[resultIdx] = (*Variant)(&argv[n])
		srcArgIndexes[resultIdx] = n
		resultIdx += 1
	}
	for n := resultIdx; n < formalArgc; n++ {
		if result[n] == nil {
			result[n] = (*Variant)(NewVariantScode(win32.DISP_E_PARAMNOTFOUND))
		}
	}
	return result, srcArgIndexes
}

func HiMetricToPixel(hiX, hiY int32) (int32, int32) {
	hdcScreen := win32.GetDC(0)
	nPixelsPerInchX := win32.GetDeviceCaps(hdcScreen, win32.LOGPIXELSX)
	nPixelsPerInchY := win32.GetDeviceCaps(hdcScreen, win32.LOGPIXELSY)
	win32.ReleaseDC(0, hdcScreen)
	return win32.MulDiv(nPixelsPerInchX, hiX, 2540),
		win32.MulDiv(nPixelsPerInchY, hiY, 2540)
}

type OleClientConstraint interface {
	GetOleClient() *OleClient
}

func As[TTo OleClientConstraint, TFrom OleClientConstraint](from TFrom) TTo {
	return *(*TTo)(unsafe.Pointer(&from))
}
