#pragma once

#include "utils.h"

class TypeLib;

class MethodInfo
{
public:
	MethodInfo(
			const CComPtr< ITypeInfo >& typeInfo,
			IID iid,
			UINT index,
			const TypeLib* typeLib );
	~MethodInfo();

	CComPtr< ITypeInfo > typeInfo;
	IID iid;
	FUNCDESC* funcdesc;
	const TypeLib* typeLib;

	HRESULT Invoke( IDispatch* obj, std::vector< CComVariant >& args, OUT VARIANT* presult, OUT EXCEPINFO* pexcepInfo );
	v8::Local< v8::Value > GetInvokeResult( HRESULT hr, VARIANT& result, EXCEPINFO& exception );
};

