:ukraine:

# GO-COM

GO-COM is a golang library that helps working with [COM](https://docs.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal) easier.

## What go-com provides

* COM data type wrappers, such as HRESULT, BSTR, VARIANT, and SAFEARRAY. 
* COM related helper functions
* COM resource lifecycle management, through the Scope type
* A pattern to implement COM interfaces, and some pre-built implementations.
* IDispatch driver, through the OleClient type
* Some general purpose IDispatch implementations.

## How do I use this library?

If you're working with standard COM interfaces that Windows defined, 
their definitions would probably already be included in the go-win32api library.
In this case, use them as is, 
and utilize the data type wrappers, helper functions, 
lifecycle management mechanism in this library. 

When a com interface require you to provide 
an event-listener/callback interface implementation, 
you can follow the pattern established in this library.  

If you're working with non-standard COM interfaces, hopefully you can find a 
typelib(tlb) that describes the interfaces. With the tlb in place, 
you can use the [GO-TlbImp](https://github.com/zzl/go-tlbimp)
tool to generate interface definitions and 
event-listener/callback implementations.

## Example projects
* [go-word-automation](https://github.com/zzl/go-word-automation)
* [go-excel-automation](https://github.com/zzl/go-excel-automation)
