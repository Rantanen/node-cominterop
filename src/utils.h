#pragma once

#define ATLASSERT( expr )
#include <atlbase.h>

#include <string>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#include <Windows.h>

#include <nan.h>
#include "JsException.h"

#define VERIFY( hr ) { \
		HRESULT __hr = ( hr ); \
		_ASSERTE( SUCCEEDED( hr ) ); \
		if( ! SUCCEEDED( __hr ) ) \
			JsException::Throw( __hr ); \
	}

inline std::string ToUTF8( const wchar_t* sz )
{
	int requiredLength = WideCharToMultiByte( CP_UTF8, 0, sz, -1, nullptr, 0, nullptr, nullptr );

	std::vector< char > vecData;
	vecData.reserve( requiredLength );
	WideCharToMultiByte( CP_UTF8, 0, sz, -1, vecData.data(), requiredLength, nullptr, nullptr );
	std::string out( vecData.data() );

	return out;
}

inline std::wstring FromUTF8( const char* sz )
{
	int requiredLength = MultiByteToWideChar( CP_UTF8, 0, sz, -1, nullptr, 0 );

	std::vector< wchar_t > vecData;
	vecData.reserve( requiredLength );
	MultiByteToWideChar( CP_UTF8, 0, sz, -1, vecData.data(), requiredLength );
	std::wstring out( vecData.data() );

	return out;
}

void InitVariant( ITypeInfo* typeInfo, const TYPEDESC& typedesc, v8::Local< v8::Value > value, OUT CComVariant& variant );
void InitVariantPtr( ITypeInfo* typeInfo, const TYPEDESC& typedesc, v8::Local< v8::Value > value, OUT CComVariant& variant );
void InitVariantEnum( ITypeInfo* typeInfo, v8::Local< v8::Value > value, OUT CComVariant& variant );
void InitVariantDispatch( ITypeInfo* typeInfo, v8::Local< v8::Value > value, OUT CComVariant& variant );

CComPtr< IDispatch > GetDispatch( ITypeInfo* typeInfo, v8::Local< v8::Value > value );

v8::Local< v8::Value > VariantToValue( ITypeInfo* typeInfo, const TYPEDESC& typedesc, const CComVariant& variant, v8::Isolate* isolate );
v8::Local< v8::Value > PtrVariantToValue( ITypeInfo* typeInfo, const TYPEDESC& typedesc, const CComVariant& variant, v8::Isolate* isolate );
v8::Local< v8::Value > UserVariantToValue( ITypeInfo* typeInfo, const TYPEDESC& typedesc, const CComVariant& variant, v8::Isolate* isolate );

bool operator < ( const GUID &guid1, const GUID &guid2 );

void Unwrap( v8::Local< v8::Value > input, OUT IDispatch** output );
void Unwrap( v8::Local< v8::Value > input, OUT IUnknown** output );
