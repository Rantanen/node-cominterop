#include "InteropInstance.h"



InteropInstance::InteropInstance( const CComPtr< IDispatch >& ptr )
	: instance( ptr )
{
}


InteropInstance::~InteropInstance()
{
}
