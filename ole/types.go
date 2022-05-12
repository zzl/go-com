package ole

import (
	"time"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
)

type NamedArg struct {
	Name  string
	Value interface{}
}

type NamedArgs map[string]interface{}

type Date float64

func (me Date) ToGoTime() time.Time {
	return time.UnixMilli(int64((me - 25569) * 24 * 3600 * 1000))
}

func NewOleDateFromGoTime(t time.Time) Date {
	return Date(float64(t.UnixMilli())/(24*3600*1000) + 25569)
}

type Currency win32.CY

type Decimal win32.DECIMAL

type DispatchClass struct {
	OleClient
}

func NewDispatchClass(pDisp *win32.IDispatch, scoped bool) *DispatchClass {
	p := &DispatchClass{OleClient{pDisp}}
	if scoped {
		com.AddScopedComPtr(&p.IUnknown)
	}
	return p
}
