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
* [go-webview2](https://github.com/zzl/go-webview2)
* [go-wmi](https://github.com/zzl/go-wmi)

## About Scopes
There are several kinds of COM resources that need to be explicitly freed
once you're done with them(if the ownership is on your side).

Taking care of these resource management tasks is tedious and error prone. 

In c++, you can use resource wrapper objects that free resources in
their destructors, which are automatically called when the wrapper 
objects go out of scope.  

In golang, there's no such language construct. We have to invent our own wheel.
That's where Scopes come into the scene in GO-COM. 

Resources to be freed are added into a scope, when the scope is left, 
the resources in it are freed automatically. 
You might think it's just a bit better than freeing each resource individually, 
if at all. That's OK. Scope usage is not mandatory in GO-COM. 

However, in GO-TlbImp generated codes, Scope is required to support method
call chaining, which is essential for a fluent API.

The resource types that could be added to Scopes includes: 
* COM interface pointer
* BSTR
* VARAINT
* SAFEARRAY

In most cases, add this line at the beginning of a function would be enough 
to introduce scope into the function:

```defer com.NewScope().Leave()```

If there are many resources created inside a loop, creating a new scope 
in the loop body might be a good idea, to avoid too many resources accumulated
waiting for free.
```
for {
    localScope := com.NewScope()
    // code that create resources ...
    localScope.Leave()
}
```
Or use a provided helper funciton:
```
for {
    com.WithScope(function() {
        // code that create resources ...
    }
}
```
The most common way to add a resource into the scope is:
```
com.AddToScope(aResourceObject)
```
This will add the resource into the closest scope.
