
#include "common.h"
#include "CollectionInfo.h"
#include "InteropType.h"
#include "InteropInstance.h"
#include "MethodInfo.h"
#include "TypeLib.h"
#include <memory>

#include <iostream>

namespace {
	void DoInvokeAsync( uv_work_t* req );
	void DoInvokeAfter( uv_work_t* req, int status );
	v8::Local< v8::Value > GetInvokeResult( HRESULT hr, MethodInfo* methodInfo, VARIANT& result, EXCEPINFO& exception );
}

struct InvokeBaton
{
	uv_work_t request;

	// Make sure the target and callee don't go out of scope.
	Nan::Persistent< v8::Object > target;
	Nan::Persistent< v8::Function > callee;
	std::unique_ptr< std::vector< CComVariant > > pargs;

	// Method info reference is held by the callee.
	// The pointer should stay alive as long as the callee stays alive.
	InteropInstance* obj;
	MethodInfo* methodInfo;

	VARIANT result;
	EXCEPINFO exception;
	HRESULT hr;

	Nan::Persistent< v8::Promise::Resolver > resolver;
};

InteropType::InteropType( const CComPtr< ITypeInfo >& typeInfo, TYPEATTR* typeattr, const TypeLib* typeLib )
	: typeInfo( typeInfo ), typeattr( typeattr ), hasInit( false ), typeLib( typeLib )
{
	// Get the type name.
	CComBSTR bstrName;
	typeInfo->GetDocumentation( MEMBERID_NIL, OUT &bstrName, nullptr, nullptr, nullptr );
	name = Nan::New( ToUTF8( bstrName ).c_str() ).ToLocalChecked();

	// Init the constructor template here. We'll need this when we are initing other classes.
	constructorTemplate = Nan::New< v8::FunctionTemplate >( New, Nan::New< v8::External >( this ) );
	constructorTemplate->SetClassName( name );
	constructorTemplate->InstanceTemplate()->SetInternalFieldCount( 1 );

	asyncConstructorTemplate = Nan::New< v8::FunctionTemplate >( NewAsync, Nan::New< v8::External >( this ) );
	asyncConstructorTemplate->SetClassName( Nan::New( ( ToUTF8( bstrName ) + "$Async" ).c_str() ).ToLocalChecked() );
	asyncConstructorTemplate->InstanceTemplate()->SetInternalFieldCount( 1 );
}

void InteropType::Init( const TypeLib* typeLib )
{
	if( hasInit ) return;
	hasInit = true;
	this->typeLib = typeLib;

	// Check if this is a coclass that implements a type.
	bool inherited = false;
	for( WORD i = 0; i < typeattr->cImplTypes; ++i )
	{
		// Ensure this is the default implemented type that we're processing.
		INT implTypeFlags;
		typeInfo->GetImplTypeFlags( i, &implTypeFlags );
		if( ( implTypeFlags & IMPLTYPEFLAG_FRESTRICTED ) != 0 ||
			( implTypeFlags & IMPLTYPEFLAG_FDEFAULT ) == 0 )
			continue;

		// Only one interface is supported in JavaScript.
		_ASSERTE( !inherited );

		// Get the implemented type info.
		HREFTYPE implRef;
		typeInfo->GetRefTypeOfImplType( i, &implRef );
		CComPtr< ITypeInfo > implTypeInfo;
		typeInfo->GetRefTypeInfo( implRef, &implTypeInfo );

		// Inherit from the type template.
		std::shared_ptr< InteropType > implType = typeLib->GetInteropType( implTypeInfo );
		implType->Init( typeLib );
		implType->AddSubclass( this );

		// Set the prototype path.
		constructorTemplate->Inherit( implType->constructorTemplate );
		asyncConstructorTemplate->Inherit( implType->asyncConstructorTemplate );
		inherited = true;
	}

	// Create JS function templates for all COM functions.
	for( WORD i = 0; i < typeattr->cFuncs; ++i )
	{
		// Generate the method info.
		std::unique_ptr< MethodInfo > methodInfo( new MethodInfo( typeInfo, typeattr->guid, i, typeLib ) );
		if( methodInfo->funcdesc->wFuncFlags & FUNCFLAG_FRESTRICTED )
			continue;

		// Grab the important bits from the methodInfo before we release it to v8::External.
		CComBSTR bstrFuncName;
		typeInfo->GetDocumentation( methodInfo->funcdesc->memid, OUT &bstrFuncName, nullptr, nullptr, nullptr );
		INVOKEKIND invkind = methodInfo->funcdesc->invkind;

		// If the type has 'Add' method it can be used to construct
		// the type from an array.
		if( wcscmp( bstrFuncName, L"Add" ) == 0 &&
			methodInfo->funcdesc->cParams == 2 &&
			methodInfo->funcdesc->lprgelemdescParam[ 0 ].tdesc.vt == VT_I4 &&
			methodInfo->funcdesc->lprgelemdescParam[ 0 ].paramdesc.wParamFlags & PARAMFLAG_FIN )
		{
			// Suitable Add-method was found.
			// Create the collection info.
			// We need to create a new MethodInfo as the current 'methodInfo' will be released for
			// v8::External later.
			this->collectionInfo.reset(
					new CollectionInfo( new MethodInfo( typeInfo, typeattr->guid, i, typeLib ) ) );
		}

		// Create the member function templates.
		//
		// We want to pass in methodInfo as external data so we need to do the
		// function template definition ourselves. Nan::SetPrototypeMethod doesn't
		// support the data parameter.
		v8::Local< v8::Value > methodLocal = Nan::New< v8::External >( methodInfo.release() );
		v8::Local< v8::FunctionTemplate > funcTemplate = Nan::New< v8::FunctionTemplate >(
				Invoke, methodLocal, Nan::New< v8::Signature >( constructorTemplate ) );
		v8::Local< v8::FunctionTemplate > asyncFuncTemplate = Nan::New< v8::FunctionTemplate >(
				InvokeAsync, methodLocal, Nan::New< v8::Signature >( asyncConstructorTemplate ) );

		// Figure out whether this is a property getter/setter.
		std::string name = ToUTF8( bstrFuncName );
		if( invkind == INVOKE_PROPERTYGET ) {

			// Getter
			funcTemplate->Set( Nan::New( "getter" ).ToLocalChecked(), Nan::New( name.c_str() ).ToLocalChecked() );
			asyncFuncTemplate->Set( Nan::New( "getter" ).ToLocalChecked(), Nan::New( name.c_str() ).ToLocalChecked() );
			name = "get_" + name;

		} else if( invkind == INVOKE_PROPERTYPUT ) {

			// Setter
			funcTemplate->Set( Nan::New( "setter" ).ToLocalChecked(), Nan::New( name.c_str() ).ToLocalChecked() );
			asyncFuncTemplate->Set( Nan::New( "setter" ).ToLocalChecked(), Nan::New( name.c_str() ).ToLocalChecked() );
			name = "put_" + name;
		}

		// Name the function.
		v8::Local< v8::String > funcName = Nan::New( name.c_str() ).ToLocalChecked();
		funcTemplate->SetClassName( funcName );
		asyncFuncTemplate->SetClassName( funcName );

		// Finally set the member function on the prototypes.
		constructorTemplate->PrototypeTemplate()->Set( funcName, funcTemplate );
		asyncConstructorTemplate->PrototypeTemplate()->Set( funcName, asyncFuncTemplate );
	}

	// Store the constructors.
	constructor.Reset( constructorTemplate->GetFunction() );
	asyncConstructor.Reset( asyncConstructorTemplate->GetFunction() );
}

/**
 * Destructor
 */
InteropType::~InteropType()
{
	if( typeattr )
	{
		typeInfo->ReleaseTypeAttr( typeattr );
		typeattr = nullptr;
	}
}

/**
 * Create instance of the item.
 */
CComPtr< IDispatch > InteropType::CreateInstance()
{
	// Try to create the instance normally.
	void* voidPtr;
	if( !SUCCEEDED( typeInfo->CreateInstance( nullptr, IID_IDispatch, OUT &voidPtr ) ) )
		JsException::ThrowCantCreate( typeInfo );

	// Return as CComPtr so we can remain ignorant of reference counting. :p
	CComPtr< IDispatch > idisp;
	idisp.Attach( static_cast< IDispatch* >( voidPtr ) );
	return idisp;
}

/**
 * Node constructor method
 */
NAN_METHOD( InteropType::New )
{
	// Unwrap the method data.
	v8::Local< v8::External > external = v8::Local< v8::External >::Cast( info.Data() );
	InteropType* interopType = reinterpret_cast< InteropType* >( external->Value() );

	// New can be either called with the external pointer when wrapping existing
	// instances or with no parameters when creating instances.
	CComPtr< IDispatch > ptr;
	if( info.Length() > 0 && info[ 0 ]->IsExternal() )
	{
		// External pointer found. Wrap that one.
		v8::Local< v8::External > externalPtrToWrap = v8::Local< v8::External >::Cast( info[ 0 ] );
		ptr = reinterpret_cast< IDispatch* >( externalPtrToWrap->Value() );
	}
	else
	{
		// No pointer. Create a new one.
		void* voidPtr;
		if( !SUCCEEDED( interopType->typeInfo->CreateInstance( nullptr, IID_IDispatch, OUT &voidPtr ) ) )
		{
			Nan::ThrowTypeError( "Could not create COM instance. Ensure the type is a coclass." );
			return;
		}

		ptr = reinterpret_cast< IDispatch* >( voidPtr );
	}

	// Wrap the pointer.
	InteropInstance* obj = new InteropInstance( ptr );
	obj->Wrap( info.This() );

	// Create the read-only hidden async member.
	v8::Local< v8::Function > asyncCtorLocal = Nan::New( interopType->asyncConstructor );
	v8::Local< v8::Value > argv[ 1 ] = { Nan::New< v8::External >( ptr.p ) };
	info.This()->DefineOwnProperty(
		Nan::GetCurrentContext(),
		Nan::New( "Async" ).ToLocalChecked(),
		asyncCtorLocal->NewInstance( 1, argv ),
		static_cast< v8::PropertyAttribute >(
				v8::PropertyAttribute::DontDelete |
				v8::PropertyAttribute::DontEnum |
				v8::PropertyAttribute::ReadOnly ) );

	info.GetReturnValue().Set( info.This() );
}

/**
 * Creates the async JavaScript interface for the object.
 */
NAN_METHOD( InteropType::NewAsync )
{
	// Unwrap the method data.
	v8::Local< v8::External > external = v8::Local< v8::External >::Cast( info.Data() );
	InteropType* interopType = reinterpret_cast< InteropType* >( external->Value() );

	// Async interfaces can only be created for an existing object.
	CComPtr< IDispatch > ptr;
	if( info.Length() > 0 && info[ 0 ]->IsExternal() )
	{
		v8::Local< v8::External > externalPtrToWrap = v8::Local< v8::External >::Cast( info[ 0 ] );
		ptr = reinterpret_cast< IDispatch* >( externalPtrToWrap->Value() );
	}
	else
	{
		return Nan::ThrowTypeError( "Async interfaces are available behind .Async member on objects." );
	}

	// Wrap the pointer.
	InteropInstance* obj = new InteropInstance( ptr );
	obj->Wrap( info.This() );
	info.GetReturnValue().Set( info.This() );
}

/**
 * Invokes a method synchronously
 */
NAN_METHOD( InteropType::Invoke )
{
	try
	{
		// Delegate
		InvokeSyncOrAsync( false, info );
	}
	catch( JsException ex )
	{
		Nan::ThrowError( ex.GetError() );
	}
}

/**
 * Invokes a method asynchronously
 */
NAN_METHOD( InteropType::InvokeAsync )
{
	try
	{
		// Delegate
		InvokeSyncOrAsync( true, info );
	}
	catch( JsException ex )
	{
		Nan::ThrowError( ex.GetError() );
	}
}

/**
 * Does the method invocation.
 */
void InteropType::InvokeSyncOrAsync( bool isAsync, Nan::NAN_METHOD_ARGS_TYPE info )
{
	// Unwrap the bound method info.
	v8::Local< v8::External > externalData = v8::Local< v8::External >::Cast( info.Data() );
	MethodInfo* methodInfo = reinterpret_cast< MethodInfo* >( externalData->Value() );

	// Gather the parameters.
	std::unique_ptr< std::vector< CComVariant > > pargs( new std::vector< CComVariant >() );
	pargs->resize( methodInfo->funcdesc->cParams );
	for( int i = 0; i < methodInfo->funcdesc->cParams; ++i )
	{
		try
		{
			ELEMDESC& elem = methodInfo->funcdesc->lprgelemdescParam[ i ];

			// Parameters are stored in reverse order.
			unsigned int rgvarg_i = methodInfo->funcdesc->cParams - i - 1;

			// Check whether there exists a JS parameter for the current COM parameter.
			if( info.Length() <= i )
			{
				// No parameter exists. Try to use a default.
				if( elem.paramdesc.pparamdescex )
				{
					// Default value.
					( *pargs )[ rgvarg_i ] = elem.paramdesc.pparamdescex->varDefaultValue;
				}
				else if( elem.paramdesc.wParamFlags & PARAMFLAG_FOPT )
				{
					// Optional value. Leave empty.
					( *pargs )[ rgvarg_i ].ChangeType( elem.tdesc.vt );
				}
				else
				{
					// Try to init from as the last resort.
					InitVariant( methodInfo->typeInfo, elem.tdesc, Nan::Undefined(), OUT ( *pargs )[ rgvarg_i ] );
				}
			}
			else
			{
				// We have JS parameter. Convert it.
				InitVariant( methodInfo->typeInfo, elem.tdesc, info[ i ], OUT ( *pargs )[ rgvarg_i ] );
			}

		}
		catch( JsException ex )
		{
			JsException::ThrowParameter( i, ex );
		}

	}  // end for

	// Check for sync vs async call.
	InteropInstance* obj = InteropInstance::Unwrap< InteropInstance >( info.This() );
	if( !isAsync )
	{
		// Synchronous call.
		VARIANT result;
		EXCEPINFO exception;
		HRESULT hr = methodInfo->Invoke( obj->instance, *pargs, OUT &result, OUT &exception );

		info.GetReturnValue().Set( methodInfo->GetInvokeResult( hr, result, exception ) );
	}
	else
	{
		// Asynchronous call.

		// Create the baton passed to the other thread.
		std::unique_ptr< InvokeBaton > baton( new InvokeBaton() );
		baton->request.data = baton.get();

		// Set the data.
		baton->target.Reset( info.This() );
		baton->callee.Reset( info.Callee() );
		baton->pargs.swap( pargs );
		baton->methodInfo = methodInfo;
		baton->obj = obj;

		// Create the promise.
		auto resolver = v8::Promise::Resolver::New( info.GetIsolate() );
		baton->resolver.Reset( resolver );

		// Queue the invocation and release the baton.
		// libuv should take care of releasing it now.
		uv_queue_work( uv_default_loop(), &baton.get()->request, DoInvokeAsync, DoInvokeAfter );
		baton.release();

		// Return the promise.
		info.GetReturnValue().Set( resolver->GetPromise() );
	}
}

namespace
{

	/**
	 * Asynchronous invoke callback.
	 *
	 * Executed in a different thread. No access to v8 internals.
	 */
	void DoInvokeAsync( uv_work_t* req )
	{
		InvokeBaton* baton = static_cast< InvokeBaton* >( req->data );
		baton->hr = baton->methodInfo->Invoke( baton->obj->instance, *( baton->pargs ), OUT &baton->result, OUT &baton->exception );
	}

	/**
	 * Asynchronou result callback.
	 *
	 * Executed back in v8-thread.
	 */
	void DoInvokeAfter( uv_work_t* req, int status )
	{
		v8::HandleScope scope( v8::Isolate::GetCurrent() );

		// Fetch the Promise from the baton.
		InvokeBaton* baton = static_cast< InvokeBaton* >( req->data );
		auto resolver = Nan::New( baton->resolver );

		try
		{
			// Resolve the promise.
			// GetInvokeResult will throw exception if the invoke failed.
			resolver->Resolve( baton->methodInfo->GetInvokeResult( baton->hr, baton->result, baton->exception ) );
		}
		catch( JsException ex )
		{
			// Reject the promise on exception.
			resolver->Reject( ex.GetError() );
		}
	}

}
