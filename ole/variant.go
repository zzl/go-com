package ole

import (
	"fmt"
	"math"
	"reflect"
	"strconv"
	"strings"
	"time"
	"unsafe"

	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
)

type Variant win32.VARIANT

func NewVariant(value interface{}) *Variant {
	switch val := value.(type) {
	case *win32.VARIANT:
		return (*Variant)(val).Copy()
	case string:
		return NewVariantString(val)
	case bool:
		return NewVariantBool(val)
	case *win32.IDispatch:
		return NewVariantDispatch(val)
	case *win32.IUnknown:
		v := &Variant{}
		v.Vt = uint16(win32.VT_UNKNOWN)
		*v.PunkVal() = val
		return v
	case nil:
		return NewVariantScode(win32.DISP_E_PARAMNOTFOUND)
	case int:
		v := &Variant{}
		if val >= math.MinInt32 && val <= math.MaxInt32 {
			v.Vt = uint16(win32.VT_INT)
			*v.IntVal() = int32(val)
		} else {
			v.Vt = uint16(win32.VT_I8)
			*v.LlVal() = int64(val)
		}
		return v
	case uint:
		v := &Variant{}
		if val <= math.MaxUint32 {
			v.Vt = uint16(win32.VT_UINT)
			*v.UintVal() = uint32(val)
		} else {
			v.Vt = uint16(win32.VT_UI8)
			*v.UllVal() = uint64(val)
		}
		return v
	case int8:
		v := &Variant{}
		v.Vt = uint16(win32.VT_I1)
		//*v.CVal() = win32.CHAR(val)
		*v.CVal() = val
		return v
	case uint8:
		v := &Variant{}
		v.Vt = uint16(win32.VT_I1)
		*v.BVal() = val
		return v
	case int16:
		v := &Variant{}
		v.Vt = uint16(win32.VT_I2)
		*v.IVal() = val
		return v
	case uint16:
		v := &Variant{}
		v.Vt = uint16(win32.VT_UI2)
		*v.UiVal() = val
		return v
	case int32:
		v := &Variant{}
		v.Vt = uint16(win32.VT_I4)
		*v.LVal() = val
		return v
	case uint32:
		v := &Variant{}
		v.Vt = uint16(win32.VT_UI4)
		*v.UlVal() = val
		return v
	case int64:
		v := &Variant{}
		v.Vt = uint16(win32.VT_I8)
		*v.LlVal() = val
		return v
	case uint64:
		v := &Variant{}
		v.Vt = uint16(win32.VT_UI8)
		*v.UllVal() = val
		return v
	case time.Time:
		v := &Variant{}
		v.Vt = uint16(win32.VT_DATE)
		*v.Date() = float64(NewOleDateFromGoTime(val))
		return v
	default:
		return nil
	}
}

func (this *Variant) Copy() *Variant {
	var v2 win32.VARIANT
	win32.VariantCopy((*win32.VARIANT)(this), &v2)
	return (*Variant)(&v2)
}

func (this *Variant) Clear() {
	win32.VariantClear((*win32.VARIANT)(this))
}

//
func (this *Variant) ChangeType(vt win32.VARENUM) (*Variant, error) {
	var v Variant
	hr := win32.VariantChangeType((*win32.VARIANT)(&v),
		(*win32.VARIANT)(this), 0, uint16(vt))
	return &v, com.NewErrorOrNil(hr)
}

func (this *Variant) ValueOfType(typ reflect.Type) (value interface{}, err error) {
	kind := typ.Kind()
	switch kind {
	case reflect.Bool:
		return this.ToBool()
	case reflect.Int:
		return this.ToInt()
	case reflect.Int8:
		return this.ToInt8()
	case reflect.Int16:
		return this.ToInt16()
	case reflect.Int32:
		return this.ToInt32()
	case reflect.Int64:
		return this.ToInt64()
	case reflect.Uint:
		return this.ToUInt()
	case reflect.Uint8:
		return this.ToUint8()
	case reflect.Uint16:
		return this.ToUint16()
	case reflect.Uint32:
		return this.ToUint32()
	case reflect.Uint64:
		return this.ToUint64()
	case reflect.Uintptr:
		v, e := this.ToUint64()
		return uintptr(v), e
	case reflect.Float32:
		return this.ToFloat32()
	case reflect.Float64:
		return this.ToFloat64()
	case reflect.String:
		return this.ToString()
	}
	return nil, fmt.Errorf("unsupported type")
}

func (this *Variant) ToInt8() (int8, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return v, nil
	case int16:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case int32:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case int64:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case int:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case uint8:
		return int8(v), nil
	case uint16:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case uint32:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case uint64:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case uint:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case float32:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case float64:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(v), nil
	case *uint16:
		n, err := strconv.Atoi(win32.PwstrToStr(v))
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		if n > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int8(n), nil
	case bool:
		if v {
			return -1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

func (this *Variant) ToUint8() (uint8, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return uint8(v), nil
	case int16:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case int32:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case int64:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case int:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case uint8:
		return v, nil
	case uint16:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case uint32:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case uint64:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case uint:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case float32:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case float64:
		if v > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(v), nil
	case *uint16:
		n, err := strconv.Atoi(win32.PwstrToStr(v))
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		if n > 255 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint8(n), nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

func (this *Variant) ToByte() (byte, error) {
	return this.ToUint8()
}

//
func (this *Variant) ToInt16() (int16, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return int16(v), nil
	case int16:
		return v, nil
	case int32:
		if v > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(v), nil
	case int64:
		if v > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(v), nil
	case int:
		if v > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(v), nil
	case uint8:
		return int16(v), nil
	case uint16:
		return int16(v), nil
	case uint32:
		if v > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(v), nil
	case uint64:
		if v > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(v), nil
	case uint:
		if v > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(v), nil
	case float32:
		if v > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(v), nil
	case float64:
		if v > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(v), nil
	case *uint16:
		n, err := strconv.Atoi(win32.PwstrToStr(v))
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		if n > math.MaxInt16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int16(n), nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

func (this *Variant) ToUint16() (uint16, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return uint16(v), nil
	case int16:
		return uint16(v), nil
	case int32:
		if v > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(v), nil
	case int64:
		if v > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(v), nil
	case int:
		if v > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(v), nil
	case uint8:
		return uint16(v), nil
	case uint16:
		return v, nil
	case uint32:
		if v > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(v), nil
	case uint64:
		if v > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(v), nil
	case uint:
		if v > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(v), nil
	case float32:
		if v > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(v), nil
	case float64:
		if v > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(v), nil
	case *uint16:
		n, err := strconv.Atoi(win32.PwstrToStr(v))
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		if n > math.MaxUint16 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint16(n), nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

//
func (this *Variant) ToInt32() (int32, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return int32(v), nil
	case int16:
		return int32(v), nil
	case int32:
		return v, nil
	case int64:
		if v > math.MaxInt32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int32(v), nil
	case int:
		if v > math.MaxInt32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int32(v), nil
	case uint8:
		return int32(v), nil
	case uint16:
		return int32(v), nil
	case uint32:
		return int32(v), nil
	case uint64:
		if v > math.MaxInt32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int32(v), nil
	case uint:
		if v > math.MaxInt32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int32(v), nil
	case float32:
		if v > math.MaxInt32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int32(v), nil
	case float64:
		if v > math.MaxInt32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int32(v), nil
	case *uint16:
		n, err := strconv.Atoi(win32.PwstrToStr(v))
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		if n > math.MaxInt32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return int32(n), nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

func (this *Variant) ToUint32() (uint32, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return uint32(v), nil
	case int16:
		return uint32(v), nil
	case int32:
		return uint32(v), nil
	case int64:
		if v > math.MaxUint32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint32(v), nil
	case int:
		return uint32(v), nil
	case uint8:
		return uint32(v), nil
	case uint16:
		return uint32(v), nil
	case uint32:
		return v, nil
	case uint64:
		if v > math.MaxUint32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint32(v), nil
	case uint:
		if v > math.MaxUint32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint32(v), nil
	case float32:
		if v > math.MaxUint32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint32(v), nil
	case float64:
		if v > math.MaxUint32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint32(v), nil
	case *uint16:
		n, err := strconv.Atoi(win32.PwstrToStr(v))
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		if n > math.MaxUint32 {
			return 0, com.NewError(win32.DISP_E_OVERFLOW)
		}
		return uint32(n), nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

//
func (this *Variant) ToInt64() (int64, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return int64(v), nil
	case int16:
		return int64(v), nil
	case int32:
		return int64(v), nil
	case int64:
		return v, nil
	case int:
		return int64(v), nil
	case uint8:
		return int64(v), nil
	case uint16:
		return int64(v), nil
	case uint32:
		return int64(v), nil
	case uint64:
		return int64(v), nil
	case uint:
		return int64(v), nil
	case float32:
		return int64(v), nil
	case float64:
		return int64(v), nil
	case *uint16:
		n, err := strconv.Atoi(win32.PwstrToStr(v))
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		return int64(n), nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

func (this *Variant) ToInt() (int, error) {
	ret, err := this.ToInt64()
	return int(ret), err
}

func (this *Variant) ToUint64() (uint64, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return uint64(v), nil
	case int16:
		return uint64(v), nil
	case int32:
		return uint64(v), nil
	case int64:
		return uint64(v), nil
	case int:
		return uint64(v), nil
	case uint8:
		return uint64(v), nil
	case uint16:
		return uint64(v), nil
	case uint32:
		return uint64(v), nil
	case uint64:
		return v, nil
	case uint:
		return uint64(v), nil
	case float32:
		return uint64(v), nil
	case float64:
		return uint64(v), nil
	case *uint16:
		n, err := strconv.Atoi(win32.PwstrToStr(v))
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		return uint64(n), nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

func (this *Variant) ToUInt() (uint, error) {
	ret, err := this.ToUint64()
	return uint(ret), err
}

//
func (this *Variant) ToFloat32() (float32, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return float32(v), nil
	case int16:
		return float32(v), nil
	case int32:
		return float32(v), nil
	case int64:
		return float32(v), nil
	case int:
		return float32(v), nil
	case uint8:
		return float32(v), nil
	case uint16:
		return float32(v), nil
	case uint32:
		return float32(v), nil
	case uint64:
		return float32(v), nil
	case uint:
		return float32(v), nil
	case float32:
		return v, nil
	case float64:
		return float32(v), nil
	case *uint16:
		f, err := strconv.ParseFloat(win32.PwstrToStr(v), 32)
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		return float32(f), nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
}

//
func (this *Variant) ToFloat64() (float64, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return float64(v), nil
	case int16:
		return float64(v), nil
	case int32:
		return float64(v), nil
	case int64:
		return float64(v), nil
	case int:
		return float64(v), nil
	case uint8:
		return float64(v), nil
	case uint16:
		return float64(v), nil
	case uint32:
		return float64(v), nil
	case uint64:
		return float64(v), nil
	case uint:
		return float64(v), nil
	case float32:
		return float64(v), nil
	case float64:
		return v, nil
	case *uint16:
		f, err := strconv.ParseFloat(win32.PwstrToStr(v), 64)
		if err != nil {
			return 0, com.NewError(win32.DISP_E_TYPEMISMATCH)
		}
		return f, nil
	case bool:
		if v {
			return 1, nil
		} else {
			return 0, nil
		}
	}
	return 0, nil
}

//
func (this *Variant) ToTime() (time.Time, error) {
	v, vt := this, uint16(win32.VT_DATE)
	if v.Vt != vt {
		if v.Vt == vt|uint16(win32.VT_BYREF) {
			return Date(*v.PdateVal()).ToGoTime(), nil
		}
		v = &Variant{}
		hr := win32.VariantChangeType((*win32.VARIANT)(v), (*win32.VARIANT)(this), 0, vt)
		if win32.FAILED(hr) {
			return time.Time{}, com.NewError(hr)
		}
	}
	return Date(v.DblValVal()).ToGoTime(), nil
}

func (this *Variant) ToCurrency() (Currency, error) {
	v, vt := this, uint16(win32.VT_CY)
	if v.Vt != vt {
		if v.Vt == vt|uint16(win32.VT_BYREF) {
			return Currency(*v.PcyValVal()), nil
		}
		v = &Variant{}
		hr := win32.VariantChangeType((*win32.VARIANT)(v), (*win32.VARIANT)(this), 0, vt)
		if win32.FAILED(hr) {
			return Currency{}, com.NewError(hr)
		}
	}
	return Currency(v.CyValVal()), nil
}

func (this *Variant) ToString() (string, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return strconv.Itoa(int(v)), nil
	case int16:
		return strconv.Itoa(int(v)), nil
	case int32:
		return strconv.Itoa(int(v)), nil
	case int64:
		return strconv.Itoa(int(v)), nil
	case int:
		return strconv.Itoa(v), nil
	case uint8:
		return strconv.Itoa(int(v)), nil
	case uint16:
		return strconv.Itoa(int(v)), nil
	case uint32:
		return strconv.Itoa(int(v)), nil
	case uint64:
		return strconv.Itoa(int(v)), nil
	case uint:
		return strconv.Itoa(int(v)), nil
	case float32:
		return strconv.FormatFloat(float64(v), 'f', 7, 32), nil
	case float64:
		return strconv.FormatFloat(v, 'f', 15, 64), nil
	case *uint16:
		return win32.PwstrToStr(v), nil
	case bool:
		if v {
			return "True", nil
		} else {
			return "False", nil
		}
	}
	return "", nil
}

func (this *Variant) ToIDispatch() (*win32.IDispatch, error) {
	v, vt := this, uint16(win32.VT_DISPATCH)
	if v.Vt != vt {
		if v.Vt == vt|uint16(win32.VT_BYREF) {
			return *v.PpdispValVal(), nil
		}
		v = &Variant{}
		hr := win32.VariantChangeType((*win32.VARIANT)(v), (*win32.VARIANT)(this), 0, vt)
		//defer v.Clear()
		if win32.FAILED(hr) {
			return nil, com.NewError(hr)
		}
	}
	return v.PdispValVal(), nil
}

func (this *Variant) ToIUnknonw() (*win32.IUnknown, error) {
	v, vt := this, uint16(win32.VT_UNKNOWN)
	if v.Vt != vt {
		if v.Vt == vt|uint16(win32.VT_BYREF) {
			return *v.PpunkValVal(), nil
		}
		v = &Variant{}
		hr := win32.VariantChangeType((*win32.VARIANT)(v), (*win32.VARIANT)(this), 0, vt)
		//defer v.Clear()
		if win32.FAILED(hr) {
			return nil, com.NewError(hr)
		}
	}
	return v.PunkValVal(), nil
}

//
func (this *Variant) ToHresult() (win32.HRESULT, error) {
	v, vt := this, uint16(win32.VT_ERROR)
	if v.Vt != vt {
		if v.Vt == vt|uint16(win32.VT_BYREF) {
			return *v.PscodeVal(), nil
		}
		v = &Variant{}
		hr := win32.VariantChangeType((*win32.VARIANT)(v), (*win32.VARIANT)(this), 0, vt)
		if win32.FAILED(hr) {
			return 0, com.NewError(hr)
		}
	}
	return v.ScodeVal(), nil
}

//
func (this *Variant) ToBool() (bool, error) {
	vObj := this.Value()
	switch v := vObj.(type) {
	case int8:
		return v != 0, nil
	case int16:
		return v != 0, nil
	case int32:
		return v != 0, nil
	case int64:
		return v != 0, nil
	case int:
		return v != 0, nil
	case uint8:
		return v != 0, nil
	case uint16:
		return v != 0, nil
	case uint32:
		return v != 0, nil
	case uint64:
		return v != 0, nil
	case uint:
		return v != 0, nil
	case float32:
		return v != 0, nil
	case float64:
		return v != 0, nil
	case *uint16:
		s := strings.ToLower(win32.PwstrToStr(v))
		return s == "true" || s == "1", nil
	case bool:
		return v, nil
	}
	return false, nil
}

func (this *Variant) ToDecimal() (Decimal, error) {
	v, vt := this, uint16(win32.VT_DECIMAL)
	if v.Vt != vt {
		if v.Vt == vt|uint16(win32.VT_BYREF) {
			return Decimal(*v.PdecValVal()), nil
		}
		v = &Variant{}
		hr := win32.VariantChangeType((*win32.VARIANT)(v), (*win32.VARIANT)(this), 0, vt)
		if win32.FAILED(hr) {
			return Decimal{}, com.NewError(hr)
		}
	}
	return Decimal(v.DecValVal()), nil
}

func (this *Variant) ToArray() (*win32.SAFEARRAY, error) {
	v, vt := this, uint16(win32.VT_ARRAY)
	if v.Vt != vt {
		if v.Vt == vt|uint16(win32.VT_BYREF) {
			return *v.PparrayVal(), nil
		}
		v = &Variant{}
		hr := win32.VariantChangeType((*win32.VARIANT)(v), (*win32.VARIANT)(this), 0, vt)
		if win32.FAILED(hr) {
			return nil, com.NewError(hr)
		}
	}
	return v.ParrayVal(), nil
}

//
func (this *Variant) ToInt8Ref() *int8 {
	if this.Vt != uint16(win32.VT_I1|win32.VT_BYREF) {
		return nil
	}
	return (*int8)(unsafe.Pointer(*this.PcVal()))
}

func (this *Variant) ToUint8Ref() *uint8 {
	if this.Vt != uint16(win32.VT_UI1|win32.VT_BYREF) {
		return nil
	}
	return *this.PbVal()
}

func (this *Variant) ToByteRef() *byte {
	return this.ToUint8Ref()
}

//
func (this *Variant) ToInt16Ref() *int16 {
	if this.Vt != uint16(win32.VT_I2|win32.VT_BYREF) {
		return nil
	}
	return *this.PiVal()
}

func (this *Variant) ToUint16Ref() *uint16 {
	if this.Vt != uint16(win32.VT_UI2|win32.VT_BYREF) {
		return nil
	}
	return *this.PuiVal()
}

//
func (this *Variant) ToInt32Ref() *int32 {
	if this.Vt != uint16(win32.VT_I4|win32.VT_BYREF) ||
		this.Vt != uint16(win32.VT_INT|win32.VT_BYREF) {
		return nil
	}
	return *this.PlVal()
}

func (this *Variant) ToUint32Ref() *uint32 {
	if this.Vt != uint16(win32.VT_UI4|win32.VT_BYREF) ||
		this.Vt != uint16(win32.VT_UINT|win32.VT_BYREF) {
		return nil
	}
	return *this.PulVal()
}

//
func (this *Variant) ToInt64Ref() *int64 {
	if this.Vt != uint16(win32.VT_I8|win32.VT_BYREF) {
		return nil
	}
	return *this.PllVal()
}

func (this *Variant) ToUint64Ref() *uint64 {
	if this.Vt != uint16(win32.VT_UI8|win32.VT_BYREF) {
		return nil
	}
	return *this.PullVal()
}

//
func (this *Variant) ToFloat32Ref() *float32 {
	if this.Vt != uint16(win32.VT_R4|win32.VT_BYREF) {
		return nil
	}
	return *this.PfltVal()
}

//
func (this *Variant) ToFloat64Ref() *float64 {
	if this.Vt != uint16(win32.VT_R8|win32.VT_BYREF) {
		return nil
	}
	return *this.PdblVal()
}

func (v *Variant) Value() interface{} {
	if v == nil {
		return nil
	}
	vt := win32.VARENUM(v.Vt)
	if vt&win32.VT_BYREF != 0 {
		vt &^= win32.VT_BYREF
		switch vt {
		case win32.VT_I2:
			return *v.PiValVal()
		case win32.VT_I4:
			return *v.PlValVal()
		case win32.VT_R4:
			return *v.PfltValVal()
		case win32.VT_R8:
			return *v.PdblValVal()
		case win32.VT_CY:
			return *v.PcyValVal()
		case win32.VT_DATE:
			return *v.PdateVal()
		case win32.VT_BSTR:
			return *v.PbstrValVal()
		case win32.VT_DISPATCH:
			return *v.PpdispValVal()
		case win32.VT_ERROR:
			return *v.PscodeVal()
		case win32.VT_BOOL:
			return *v.PboolValVal()
		case win32.VT_UNKNOWN:
			return *v.PpunkValVal()
		case win32.VT_DECIMAL:
			return *v.PdecValVal()
		case win32.VT_I1:
			return *v.PcValVal()
		case win32.VT_UI1:
			return *v.PbValVal()
		case win32.VT_UI2:
			return *v.PuiValVal()
		case win32.VT_UI4:
			return *v.PulValVal()
		case win32.VT_I8:
			return *v.PllValVal()
		case win32.VT_UI8:
			return *v.PullValVal()
		case win32.VT_INT:
			return *v.PintValVal()
		case win32.VT_UINT:
			return *v.PuintValVal()
		case win32.VT_HRESULT:
			return *v.PscodeVal()
		case win32.VT_SAFEARRAY:
			return *v.PparrayVal()
		case win32.VT_ARRAY:
			return *v.PparrayVal()
		case win32.VT_VARIANT:
			return (*Variant)(v.PvarValVal()).Value()
		}
	} else {
		switch vt {
		case win32.VT_I2:
			return v.IValVal()
		case win32.VT_I4:
			return v.LValVal()
		case win32.VT_R4:
			return v.FltValVal()
		case win32.VT_R8:
			return v.DblValVal()
		case win32.VT_CY:
			return v.CyValVal()
		case win32.VT_DATE:
			return v.DateVal()
		case win32.VT_BSTR:
			return v.BstrValVal()
		case win32.VT_DISPATCH:
			return v.PdispValVal()
		case win32.VT_ERROR:
			return v.ScodeVal()
		case win32.VT_BOOL:
			return v.BoolValVal()
		case win32.VT_UNKNOWN:
			return v.PunkValVal()
		case win32.VT_DECIMAL:
			return v.DecValVal()
		case win32.VT_I1:
			return v.CValVal()
		case win32.VT_UI1:
			return v.BValVal()
		case win32.VT_UI2:
			return v.UiValVal()
		case win32.VT_UI4:
			return v.UlValVal()
		case win32.VT_I8:
			return v.LlValVal()
		case win32.VT_UI8:
			return v.UllValVal()
		case win32.VT_INT:
			return v.IntValVal()
		case win32.VT_UINT:
			return v.UintValVal()
		case win32.VT_HRESULT:
			return v.ScodeVal()
		case win32.VT_SAFEARRAY:
			return v.ParrayVal()
		case win32.VT_ARRAY:
			return v.ParrayVal()
		}
	}
	return nil
}

//
func (this *Variant) ToPointer() unsafe.Pointer {
	v := this
	if v == nil {
		return nil
	}
	vt := win32.VARENUM(v.Vt)
	if vt&win32.VT_BYREF != 0 {
		vt &^= win32.VT_BYREF
		switch vt {
		case win32.VT_I2:
			return unsafe.Pointer(v.PiValVal())
		case win32.VT_I4:
			return unsafe.Pointer(v.PlValVal())
		case win32.VT_R4:
			return unsafe.Pointer(v.PfltValVal())
		case win32.VT_R8:
			return unsafe.Pointer(v.PdblValVal())
		case win32.VT_CY:
			return unsafe.Pointer(v.PcyValVal())
		case win32.VT_DATE:
			return unsafe.Pointer(v.PdateVal())
		case win32.VT_BSTR:
			return unsafe.Pointer(v.PbstrValVal())
		case win32.VT_DISPATCH:
			return unsafe.Pointer(v.PpdispValVal())
		case win32.VT_ERROR:
			return unsafe.Pointer(v.PscodeVal())
		case win32.VT_BOOL:
			return unsafe.Pointer(v.PboolValVal())
		case win32.VT_UNKNOWN:
			return unsafe.Pointer(v.PpunkValVal())
		case win32.VT_DECIMAL:
			return unsafe.Pointer(v.PdecValVal())
		case win32.VT_I1:
			return unsafe.Pointer(v.PcValVal())
		case win32.VT_UI1:
			return unsafe.Pointer(v.PbValVal())
		case win32.VT_UI2:
			return unsafe.Pointer(v.PuiValVal())
		case win32.VT_UI4:
			return unsafe.Pointer(v.PulValVal())
		case win32.VT_I8:
			return unsafe.Pointer(v.PllValVal())
		case win32.VT_UI8:
			return unsafe.Pointer(v.PullValVal())
		case win32.VT_INT:
			return unsafe.Pointer(v.PintValVal())
		case win32.VT_UINT:
			return unsafe.Pointer(v.PuintValVal())
		case win32.VT_HRESULT:
			return unsafe.Pointer(v.PscodeVal())
		case win32.VT_SAFEARRAY:
			return unsafe.Pointer(v.PparrayVal())
		case win32.VT_ARRAY:
			return unsafe.Pointer(v.PparrayVal())
		case win32.VT_VARIANT:
			return unsafe.Pointer(v.PvarValVal())
		}
	} else {
		switch vt {
		case win32.VT_I2:
			ret := v.IValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_I4:
			ret := v.LValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_R4:
			ret := v.FltValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_R8:
			ret := v.DblValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_CY:
			ret := v.CyValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_DATE:
			ret := v.DateVal()
			return unsafe.Pointer(&ret)
		case win32.VT_BSTR:
			ret := v.BstrValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_DISPATCH:
			ret := v.PdispValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_ERROR:
			ret := v.ScodeVal()
			return unsafe.Pointer(&ret)
		case win32.VT_BOOL:
			ret := v.BoolValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_UNKNOWN:
			ret := v.PunkValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_DECIMAL:
			ret := v.DecValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_I1:
			ret := v.CValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_UI1:
			ret := v.BValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_UI2:
			ret := v.UiValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_UI4:
			ret := v.UlValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_I8:
			ret := v.LlValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_UI8:
			ret := v.UllValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_INT:
			ret := v.IntValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_UINT:
			ret := v.UintValVal()
			return unsafe.Pointer(&ret)
		case win32.VT_HRESULT:
			ret := v.ScodeVal()
			return unsafe.Pointer(&ret)
		case win32.VT_SAFEARRAY:
			ret := v.ParrayVal()
			return unsafe.Pointer(&ret)
		case win32.VT_ARRAY:
			ret := v.ParrayVal()
			return unsafe.Pointer(&ret)
		}
	}
	return nil
}

//
type VariantBool struct { //24
	Vt    win32.VARENUM      //4@0
	_pad1 int32              //4@4
	Value win32.VARIANT_BOOL //2@8
	_pad2 [6]byte            //6@10
	_pad3 int64              //8@16
}

func NewVariantBool(b bool) *Variant {
	return (*Variant)(unsafe.Pointer(&VariantBool{
		//Vt: VT_BOOL, Value: ^(*(*VARIANT_BOOL)(unsafe.Pointer(&b)) - 1)}))
		Vt: win32.VT_BOOL, Value: win32.VARIANT_BOOL(-(*(*int8)(unsafe.Pointer(&b))))}))
}

type VariantString struct {
	Vt    win32.VARENUM //4@0
	_pad1 int32         //4@4
	Value win32.BSTR    //8@8
	_pad2 int64         //8@16
}

func NewVariantString(s string) *Variant {
	return (*Variant)(unsafe.Pointer(&VariantString{
		Vt: win32.VT_BSTR, Value: win32.StrToBstr(s)}))
}

//
type VariantScode struct { //24
	Vt    win32.VARENUM //4@0
	_pad1 int32         //4@4
	Value win32.HRESULT //4@8
	_pad2 [4]byte       //4@12
	_pad3 int64         //8@16
}

func NewVariantScode(hr win32.HRESULT) *Variant {
	return (*Variant)(unsafe.Pointer(&VariantScode{
		Vt: win32.VT_ERROR, Value: hr}))
}

//
type VariantDispatch struct { //24
	Vt    win32.VARENUM    //4@0
	_pad1 int32            //4@4
	Value *win32.IDispatch //8@8
	_pad3 int64            //8@16
}

func NewVariantDispatch(pDisp *win32.IDispatch) *Variant {
	return (*Variant)(unsafe.Pointer(&VariantDispatch{
		Vt: win32.VT_DISPATCH, Value: pDisp}))
}

func Var(value *win32.VARIANT) *Variant {
	return (*Variant)(value)
}
