#pragma once

#include "utils.h"

#include <nan.h>

class TypeLibLoader
{
public:
	static void Init( v8::Local< v8::Object > exports );
	static void Load( const Nan::FunctionCallbackInfo< v8::Value >& info );

private:
	static Nan::Persistent< v8::FunctionTemplate > constructorTemplate;

	TypeLibLoader() {};
	~TypeLibLoader() {};
};

