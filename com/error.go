package com

import (
	"github.com/zzl/go-win32api/win32"
	"syscall"
)

type Error win32.HRESULT

type ErrorInfo struct {
	Error       Error
	Source      string
	IID         syscall.GUID
	Description string
}

func (this *ErrorInfo) String() string {
	if this == nil || this.Error == 0 {
		return ""
	}
	s := this.Error.Error()
	if this.Description != "" {
		s += " -- " + this.Description
	}
	return s
}

func NewError(hr win32.HRESULT) Error {
	return Error(hr)
}

const OK = Error(win32.S_OK)
const FAIL = Error(win32.E_FAIL)

func NewErrorOrNil(hr win32.HRESULT) error {
	if win32.SUCCEEDED(hr) {
		return nil
	}
	return Error(hr)
}

func (me Error) Error() string {
	return win32.HRESULT_ToString(win32.HRESULT(me))
}

func (me Error) FAILED() bool {
	return me < 0
}

func (me Error) HRESULT() win32.HRESULT {
	return win32.HRESULT(me)
}

//
func SetLastError(err Error) {
	var info ErrorInfo
	info.Error = err

	var pEi *win32.IErrorInfo
	hr := win32.GetErrorInfo(0, &pEi)
	if win32.SUCCEEDED(hr) {
		var bs BStr
		pEi.GetDescription(bs.PBSTR())
		info.Description = bs.ToStringAndFree()
		pEi.GetGUID(&info.IID)
		pEi.GetSource(bs.PBSTR())
		info.Source = bs.ToStringAndFree()
	}
	context := GetContext()
	context.LastError = &info
}

func GetLastErrorInfo() *ErrorInfo {
	context := GetContext()
	return context.LastError
}
