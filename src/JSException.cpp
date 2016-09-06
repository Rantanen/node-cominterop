#include "JsException.h"

#include <string>
#include <comdef.h>
#include <sstream>
#include <iomanip>


JsException::~JsException()
{
}

void JsException::Throw( HRESULT hr )
{
	_com_error err( hr );

	std::stringstream ss;
	ss << err.ErrorMessage() << " (0x" << std::nouppercase << std::setfill( '0' ) << std::setw( 8 ) << std::hex << hr << ")";

	throw JsException( Nan::TypeError( ss.str().c_str() ) );
}

void JsException::Throw( EXCEPINFO ex )
{
	std::stringstream ss;
	ss << ToUTF8( ex.bstrDescription );

	throw JsException( Nan::TypeError( ss.str().c_str() ) );
}

void JsException::ThrowParameter( int i, JsException& ex )
{
	v8::Local< v8::Object > errObject = ex.GetError().As< v8::Object >();
	v8::Local< v8::Value > errMsg = errObject->Get( Nan::New< v8::String >( "message" ).ToLocalChecked() );
	v8::String::Utf8Value errStr( errMsg->ToString() );

	std::stringstream ss;
	ss << "Invalid argument " << ( i + 1 ) << ": " << *errStr;

	throw JsException( Nan::TypeError( ss.str().c_str() ) );
}

void JsException::ThrowCantCreate( const CComPtr< ITypeInfo >& expected )
{
	CComBSTR bstrName;
	expected->GetDocumentation( MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr );

	std::stringstream ss;
	ss << "Can't create " << ToUTF8( bstrName );

	throw JsException( Nan::TypeError( ss.str().c_str() ) );
}

void JsException::ThrowMismatch( const CComPtr< ITypeInfo >& expected, v8::Local< v8::Value > actual )
{
	CComBSTR bstrName;
	expected->GetDocumentation( MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr );
	v8::String::Utf8Value actualStr( actual->ToString() );

	std::stringstream ss;
	ss << "Type mismatch. Expected: " << ToUTF8( bstrName ) << ", received: " << *actualStr << ".";

	throw JsException( Nan::TypeError( ss.str().c_str() ) );
}

void JsException::ThrowWin32()
{
	std::string msg;

	DWORD errorID = ::GetLastError();
	if( errorID != 0 )
	{

		LPSTR msgBuffer = nullptr;
		size_t size = FormatMessageA(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
					FORMAT_MESSAGE_FROM_SYSTEM |
					FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL, errorID, MAKELANGID( LANG_NEUTRAL, SUBLANG_DEFAULT ),
				OUT ( LPSTR )&msgBuffer, 0, NULL );

		msg = std::string( msgBuffer, size );
		LocalFree( msgBuffer );
	}
	
	throw JsException( Nan::TypeError( msg.c_str() ) );
}
