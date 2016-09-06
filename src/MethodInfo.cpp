#include "MethodInfo.h"
#include "utils.h"

MethodInfo::MethodInfo( const CComPtr< ITypeInfo >& typeInfo, IID interfaceID, UINT index, const TypeLib* typeLib )
	: typeInfo( typeInfo ), iid( interfaceID ), typeLib( typeLib )
{
	VERIFY( typeInfo->GetFuncDesc( index, &funcdesc ) );
}


MethodInfo::~MethodInfo()
{
	if( funcdesc != nullptr )
	{
		typeInfo->ReleaseFuncDesc( funcdesc );
	}
}

HRESULT MethodInfo::Invoke( IDispatch* obj, std::vector< CComVariant >& args, OUT VARIANT* presult, OUT EXCEPINFO* pexcepInfo )
{
	unsigned int argErr;
	DISPPARAMS params;

	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = nullptr;
	params.cArgs = static_cast< int >( args.size() );
	params.rgvarg = args.data();

	WORD wFlags = 0;
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	switch( funcdesc->invkind )
	{
	case INVOKE_FUNC:
		wFlags |= DISPATCH_METHOD;
		break;
	case INVOKE_PROPERTYGET:
		wFlags |= DISPATCH_PROPERTYGET;
		break;
	case INVOKE_PROPERTYPUT:
		wFlags |= DISPATCH_PROPERTYPUT;
		params.cNamedArgs = 1;
		params.rgdispidNamedArgs = &dispidNamed;
		break;
	case INVOKE_PROPERTYPUTREF:
		wFlags |= DISPATCH_PROPERTYPUTREF;
		params.cNamedArgs = 1;
		params.rgdispidNamedArgs = &dispidNamed;
		break;
	}

	return typeInfo->Invoke( obj, funcdesc->memid, wFlags, &params, OUT presult, OUT pexcepInfo, OUT &argErr );
}

v8::Local< v8::Value > MethodInfo::GetInvokeResult( HRESULT hr, VARIANT& result, EXCEPINFO& exception )
{

	if( hr == DISP_E_EXCEPTION )
		JsException::Throw( exception );
	
	VERIFY( hr );

	return VariantToValue(
		typeInfo, funcdesc->elemdescFunc.tdesc,
		result, nullptr );
}
