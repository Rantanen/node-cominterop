#pragma once

#include "utils.h"
#include <nan.h>

class InteropType;

class InteropInstance : public Nan::ObjectWrap
{
public:

	InteropInstance( const CComPtr< IDispatch >& ptr );
	~InteropInstance();

	CComPtr< IDispatch > instance;

	inline void Wrap( v8::Local< v8::Object > handle ) { Nan::ObjectWrap::Wrap( handle ); }

	/*
	inline static InteropInstance* Unwrap( v8::Local< v8::Value > handle )
	{
		auto extHandle = handle.As< v8::External >();
		return reinterpret_cast< InteropInstance* >( extHandle->Value() );
	}
	*/

	friend InteropType;
};

