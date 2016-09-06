#pragma once

#include "common.h"

template< class T >
class TypeInfoPtr
{
public:
	TypeInfoPtr( ITypeInfo* typeInfo );
	~TypeInfoPtr() { Destroy(); }

	T** Out() {
		Destroy();
		return &ptr;
	}
	T* operator->() { return ptr; }
	T* get() { return ptr; }

private:
	void Destroy();

	CComPtr< ITypeInfo > typeInfo;
	T* ptr;
};
