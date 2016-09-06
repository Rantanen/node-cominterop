#include "TypeLib.h"

#include "common.h"
#include "InteropType.h"

#include <memory>

#include <iostream>

Nan::Persistent< v8::Function > TypeLib::constructor;
std::map< GUID, std::shared_ptr< InteropType > > TypeLib::types;

TypeLib::TypeLib( const CComPtr< ITypeLib >& typeLib )
	: typeLib( typeLib )
{
}


TypeLib::~TypeLib()
{
	std::cout << "TypeLib dtor" << std::endl;
}

NAN_METHOD( TypeLib::New )
{
	if( info.Length() != 1 || !info[ 0 ]->IsExternal() ) {
		Nan::ThrowTypeError( "Use Interop.load() to create the TypeLib." );
		return;
	}

	v8::Local< v8::External > external = v8::Local< v8::External >::Cast( info[ 0 ] );
	CComPtr< ITypeLib > typeLib = reinterpret_cast< ITypeLib* >( external->Value() );

	TypeLib* obj = new TypeLib( typeLib );
	obj->Wrap( info.This() );

	for( ULONG ul = 0; ul < typeLib->GetTypeInfoCount(); ul++ )
	{
		CComPtr< ITypeInfo > typeInfo;
		typeLib->GetTypeInfo( ul, OUT &typeInfo );

		CComBSTR bstrName;
		if( !SUCCEEDED( typeInfo->GetDocumentation( MEMBERID_NIL, OUT &bstrName, nullptr, nullptr, nullptr ) ) )
		{
			Nan::ThrowTypeError( "Could not load type information" );
			return;
		}

		TYPEATTR* typeattr;
		if( !SUCCEEDED( typeInfo->GetTypeAttr( &typeattr ) ) )
		{
			Nan::ThrowTypeError( "Could not load type attributes" );
			return;
		}

		std::shared_ptr< InteropType > ptr( new InteropType( typeInfo, typeattr, obj ) );
		obj->types[ typeattr->guid ] = ptr;
	}

	for( auto it = obj->types.begin(); it != obj->types.end(); ++it ) {
		it->second->Init( obj );
		info.This()->Set( it->second->name, it->second->constructor.Get( info.GetIsolate() ) );
	}

	info.GetReturnValue().Set( info.This() );
}

std::shared_ptr< InteropType > TypeLib::GetInteropType( ITypeInfo* typeInfo )
{
	TypeInfoPtr< TYPEATTR > typeattr( typeInfo );
	typeInfo->GetTypeAttr( typeattr.Out() );

	return GetInteropType( typeattr.get() );
}

std::shared_ptr< InteropType > TypeLib::GetInteropType( TYPEATTR* typeattr )
{
	auto it = types.find( typeattr->guid );
	if( it == types.end() )
		return nullptr;
	return it->second;
}

std::shared_ptr< InteropType > TypeLib::GetInteropType( GUID guid )
{
	auto it = types.find( guid );
	if( it == types.end() )
		return nullptr;
	return it->second;
}

void TypeLib::Init( v8::Local< v8::Object > exports )
{
	Nan::HandleScope scope;

	v8::Local< v8::FunctionTemplate > ctorTemplate = Nan::New< v8::FunctionTemplate >( New );
	ctorTemplate->SetClassName( Nan::New( "TypeLib" ).ToLocalChecked() );
	ctorTemplate->InstanceTemplate()->SetInternalFieldCount( 1 );

	v8::Local< v8::Function > ctor = ctorTemplate->GetFunction();
	constructor.Reset( ctor );
	exports->Set( Nan::New( "TypeLib" ).ToLocalChecked(), ctor );

}
