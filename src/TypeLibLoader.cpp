#include "TypeLibLoader.h"

#include <atlbase.h>
#include "utils.h"

#include "TypeLib.h"

void TypeLibLoader::Load( const Nan::FunctionCallbackInfo< v8::Value >& info )
{
	if( info.Length() < 1 ) {
		Nan::ThrowTypeError( "Missing library path" );
		return;
	}

	// Convert the path argument to a native string.
	v8::String::Utf8Value utf8( info[ 0 ]->ToString() );
	std::wstring str = FromUTF8( *utf8 );

	HINSTANCE module = CoLoadLibrary( const_cast< wchar_t* >( str.c_str() ), false );

	// Try to load the type library.
	CComPtr< ITypeLib > typeLib;
	if( !SUCCEEDED( LoadTypeLib( str.c_str(), OUT &typeLib ) ) ) {
		Nan::ThrowTypeError( "Could not load type library." );
		return;
	}

	v8::Local< v8::Value > argv[ 1 ] = { Nan::New< v8::External >( typeLib.p ) };
	v8::Local< v8::Function > cons = Nan::New< v8::Function >( TypeLib::constructor );
	info.GetReturnValue().Set( cons->NewInstance( Nan::GetCurrentContext(), 1, argv ).ToLocalChecked() );
}

void TypeLibLoader::Init( v8::Local< v8::Object > exports )
{
	Nan::HandleScope scope;

	v8::Local< v8::FunctionTemplate > load = Nan::New< v8::FunctionTemplate >( Load );
	exports->Set( Nan::New( "load" ).ToLocalChecked(), load->GetFunction() );
}
