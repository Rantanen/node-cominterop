#pragma once

#define ATLASSERT( expr )
#include <atlbase.h>

#include <memory>

namespace std
{
	template<>
	class default_delete< TYPEATTR >
	{
	public:
		default_delete( ITypeInfo* typeInfo ) : typeInfo( typeInfo ) {}

		void operator()( TYPEATTR *ptr )
		{
			typeInfo->ReleaseTypeAttr( ptr );
		}

	private:
		CComPtr< ITypeInfo > typeInfo;
	};
}
