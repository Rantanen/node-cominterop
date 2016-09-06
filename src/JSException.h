#pragma once

#include "common.h"
#include <exception>

class JsException : public std::exception
{
public:
	JsException( v8::Local< v8::Value > error )
		: error( error )
	{
	}

	~JsException();

	static void Throw( HRESULT hr );
	static void Throw( EXCEPINFO hr );
	static void Throw( const char* msg ) { throw JsException( Nan::TypeError( msg ) ); }
	static void ThrowParameter( int i, JsException& ex );
	static void ThrowMismatch( const CComPtr< ITypeInfo >& expected, v8::Local< v8::Value > actual );
	static void ThrowCantCreate( const CComPtr< ITypeInfo >& expected );
	static void ThrowWin32();

	v8::Local< v8::Value > GetError() { return error; }

private:
	v8::Local< v8::Value > error;
};

