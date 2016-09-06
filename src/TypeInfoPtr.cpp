#include "TypeInfoPtr.h"

#include "common.h"

template<>
TypeInfoPtr< TYPEATTR >::TypeInfoPtr( ITypeInfo* typeInfo )
	: typeInfo( typeInfo ), ptr( nullptr )
{
	VERIFY( typeInfo->GetTypeAttr( OUT &ptr ) );
}

template<>
void TypeInfoPtr< TYPEATTR >::Destroy()
{
	if( ptr != nullptr )
	{
		typeInfo->ReleaseTypeAttr( ptr );
		ptr = nullptr;
	}
}
