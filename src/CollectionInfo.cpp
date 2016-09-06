#include "CollectionInfo.h"

/**
 * Adds an item to the target collection.
 */
void CollectionInfo::AddItem( CComPtr< IDispatch > collection, CComVariant item )
{
	// Create parameters.
	std::vector< CComVariant > args;
	args.resize( 2 );
	args[ 0 ].ChangeType( VT_I4 );
	args[ 0 ].lVal = -1;
	args[ 1 ] = item;

	// Invoke the add.
	VARIANT result;
	EXCEPINFO exception;
	HRESULT hr = addMethod->Invoke( collection, args, OUT &result, OUT &exception );

	// There should be no result, but process the HR anyway in case of an exception.
	addMethod->GetInvokeResult( hr, result, exception );
}
