#pragma once

#include "utils.h"
#include <nan.h>

#include <vector>

class TypeLib;
class InteropInstance;
class MethodInfo;
class CollectionInfo;

class InteropType
{
public:
	InteropType( const CComPtr< ITypeInfo >& typeInfo, TYPEATTR* typeattr, const TypeLib* typeLib );
	~InteropType();

	void Init( const TypeLib* typeLib );

	CComPtr< IDispatch > CreateInstance();

	static NAN_METHOD( New );
	static NAN_METHOD( NewAsync );
	static NAN_METHOD( Invoke );
	static NAN_METHOD( InvokeAsync );
	static NAN_INDEX_GETTER( GetIndex );

	static void InvokeSyncOrAsync( bool async, Nan::NAN_METHOD_ARGS_TYPE info );

	void AddSubclass( InteropType* subclass ) { subclasses.push_back( subclass ); }
	InteropType* GetCoclass() {
		if( typeattr->wTypeFlags & TYPEFLAG_FCANCREATE )
			return this;

		for( auto&& subclass : subclasses )
		{
			auto* createable = subclass->GetCoclass();
			if( createable )
				return createable;
		}

		return nullptr;
	}

	v8::Local< v8::String > name;
	v8::Local< v8::FunctionTemplate > constructorTemplate;
	v8::Local< v8::FunctionTemplate > asyncConstructorTemplate;

	Nan::Persistent< v8::Function > constructor;
	Nan::Persistent< v8::Function > asyncConstructor;

	CComPtr< ITypeInfo > typeInfo;
	std::unique_ptr< CollectionInfo > collectionInfo;

private:
	const TypeLib* typeLib;

	TYPEATTR* typeattr;
	std::vector< FUNCDESC* > funcDescs;

	bool hasInit;
	CComPtr< ITypeInfo > itemInfo;
	std::vector< InteropType* > subclasses;
};

