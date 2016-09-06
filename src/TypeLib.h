#pragma once

#include "utils.h"

#include <nan.h>
#include <map>
#include <string>
#include <memory>

#include "InteropType.h"

class TypeLib : public Nan::ObjectWrap
{
public:
	TypeLib( const CComPtr< ITypeLib >& typeLib );
	~TypeLib();

	static NAN_METHOD( New );
	static void Init( v8::Local< v8::Object > exports );

	static Nan::Persistent< v8::Function > constructor;
	CComPtr< ITypeLib > typeLib;

	static std::shared_ptr< InteropType > GetInteropType( ITypeInfo* typeInfo );
	static std::shared_ptr< InteropType > GetInteropType( TYPEATTR* typeattr );
	static std::shared_ptr< InteropType > GetInteropType( GUID guid );

private:

	static std::map< GUID, std::shared_ptr< InteropType > > types;
};

