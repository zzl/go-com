package com

import "github.com/zzl/go-win32api/win32"

type Error win32.HRESULT

func NewError(hr win32.HRESULT) Error {
	return Error(hr)
}

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
