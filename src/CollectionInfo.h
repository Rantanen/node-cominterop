#pragma once

#include "MethodInfo.h"
#include <memory>

/**
 * Extra information for collection types.
 */
class CollectionInfo
{
public:

	/**
	 * Constructor. Takes ownership of method info.
	 */
	CollectionInfo( MethodInfo* methodInfo ) : addMethod( methodInfo )
	{
		_ASSERTE( methodInfo->funcdesc->cParams == 2 );
	}

	/**
	 * Destructor.
	 */
	~CollectionInfo() {}

	/**
	 * Returns the collection item TYPEDESC.
	 */
	TYPEDESC ItemDesc() {
		return addMethod->funcdesc->lprgelemdescParam[ 1 ].tdesc;
	}

	/**
	 * Adds an item to the collection.
	 */
	void AddItem( CComPtr< IDispatch > collection, CComVariant item );

private:
	std::unique_ptr< MethodInfo > addMethod;
};

