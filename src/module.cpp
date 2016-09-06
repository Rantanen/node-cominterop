
#include <nan.h>

#include "TypeLibLoader.h"
#include "TypeLib.h"

NAN_METHOD( Assert )
{
	_ASSERTE( false );
}

void InitAll( v8::Local< v8::Object > exports ) {

	CoInitialize( nullptr );

	TypeLibLoader::Init( exports );
	TypeLib::Init( exports );

#ifdef DEBUG
	Nan::HandleScope scope;
	v8::Local< v8::FunctionTemplate > assert = Nan::New< v8::FunctionTemplate >( Assert );
	exports->Set( Nan::New( "assert" ).ToLocalChecked(), assert->GetFunction() );
#endif
}

NODE_MODULE( Interop, InitAll );