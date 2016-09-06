# Node.js COM interop
### Creates runtime proxy classes for a type library.

## Installation

```
npm install cominterop
```

**Only compatible with Windows (which goes to most typelibs too anyway).**

## Usage

```
let cominterop = require( 'cominterop' );
let lib = cominterop.load( 'path/to/typelib.dll' );

// Normal constructors.
let obj = new lib.MyClass();

// propget/propput methods mapped to JavaScript getters/setters.
let value = obj.Value;

// Objects have a hidden .Async property which exposes a promise interface.
obj.Async.GetItem()
    .then( item => {
        console.log( item.Value );
    } );
```

## Caveats

- Pointer return values won't work maintain identity: `obj.Member !== obj.Member`,
  when `Member` is non-primitive.
- Support for several data types missing. `SAFE_ARRAY` the biggest one.
- Only getters supported for indexed properties: `arr[ 0 ]`.
- Uses `IDispatch` for method invocation.
- My current test libraries are limited to [M-Files API](https://www.m-files.com/api/documentation/latest/index.html).
  Other libraries may be completely incompatible without me knowing about it.

## Future plans

- Track pointers and return same references for equal pointer values.
- Add compatibility option to `load()` for mangling the member names for lower
  case. The upper case method names will confuse linters that expect these to
  be constructors.

## Performance

Preprocessing the type library should help a bit with the major performance
issues of going with raw IDispatch.

However there are still a lot of dynamic lookups happening - especially in
parameter resolution. And the method invocations are done over IDispatch (with
all the VARIANTing) instead of using direct COM vtables and other internals.
