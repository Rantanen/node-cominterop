
#include "utils.h"
#include "CollectionInfo.h"
#include "InteropInstance.h"
#include "TypeInfoPtr.h"
#include "TypeLib.h"

bool operator < (const GUID &guid1, const GUID &guid2) {
    if(guid1.Data1!=guid2.Data1) {
        return guid1.Data1 < guid2.Data1;
    }
    if(guid1.Data2!=guid2.Data2) {
        return guid1.Data2 < guid2.Data2;
    }
    if(guid1.Data3!=guid2.Data3) {
        return guid1.Data3 < guid2.Data3;
    }
    for(int i=0;i<8;i++) {
        if(guid1.Data4[i]!=guid2.Data4[i]) {
            return guid1.Data4[i] < guid2.Data4[i];
        }
    }
    return false;
}

void InitVariant( ITypeInfo* typeInfo, const TYPEDESC& typedesc, v8::Local< v8::Value > value, OUT CComVariant& variant )
{
	variant.vt = typedesc.vt;
	switch( typedesc.vt )
	{
	case VT_I1:  //signed char
		variant.cVal = static_cast< char >( value->Int32Value() );
		break;
	case VT_UI1:  //unsigned char
		variant.bVal = static_cast< unsigned char >( value->Int32Value() );
		break;
	case VT_I2:  //2 byte signed int
		variant.iVal = static_cast< SHORT >( value->Int32Value() );
		break;
	case VT_UI2:  //unsigned short
		variant.uiVal = static_cast< USHORT >( value->Int32Value() );
		break;
	case VT_I4:  //4 byte signed int
		variant.lVal = static_cast< LONG >( value->Int32Value() );
		break;
	case VT_UI4:  //ULONG
		variant.ulVal = static_cast< ULONG >( value->Int32Value() );
		break;
	case VT_I8:  //signed 64-bit int
		variant.llVal = static_cast< LONGLONG >( value->IntegerValue() );
		break;
	case VT_UI8:  //unsigned 64-bit int
		variant.ullVal = static_cast< ULONGLONG >( value->IntegerValue() );
		break;
	case VT_INT:  //signed machine int
		variant.intVal = static_cast< INT >( value->Int32Value() );
		break;
	case VT_UINT:  //unsigned machine int
		variant.uintVal = static_cast< UINT >( value->Int32Value() );
		break;
	case VT_R4:  //4 byte real
		variant.fltVal = static_cast< float >( value->NumberValue() );
		break;
	case VT_R8:  //8 byte real
		variant.dblVal = static_cast< double >( value->NumberValue() );
		break;
	case VT_DATE:  //date
	{
		auto date = v8::Local< v8::Date >::Cast( value );
		double dateValue = date->ValueOf();

		// JS measures milliseconds. COM measures days. Convert between these two.
		dateValue /= 1000 * 3600 * 24;

		// JS uses Jan 1, 1970 as epoch. COM uses Dec 30, 1899.
		dateValue -= 70 * 365 + 20;
		variant.date = dateValue;
		break;
	}
	case VT_BSTR:  //OLE Automation string
	{
		v8::String::Utf8Value utf8( value->ToString() );
		variant.bstrVal = SysAllocString( FromUTF8( *utf8 ).c_str() );
		break;
	}
	case VT_DISPATCH:  //IDispatch *
	{
		Unwrap( value, OUT &variant.pdispVal );
		break;
	}
	case VT_BOOL:  //True=-1, False=0
		variant.boolVal = value->ToBoolean()->BooleanValue() ? VARIANT_TRUE : VARIANT_FALSE;
		break;
	case VT_UNKNOWN:  //IUnknown *
		Unwrap( value, &variant.punkVal );
		break;

	case VT_INT_PTR:  //signed machine register size width
	case VT_VARIANT:  //VARIANT *
	case VT_ERROR:  //SCODE
	case VT_DECIMAL:  //16 byte fixed point
	case VT_UINT_PTR:  //unsigned machine register size width
	case VT_VOID:  //C style void
	case VT_HRESULT:  //Standard return type
	case VT_LPSTR:  //null terminated string
	case VT_LPWSTR:  //wide null terminated string
	case VT_CY:  //currency
		_ASSERTE( false );

	case VT_SAFEARRAY:  //(use VT_ARRAY in VARIANT)
	case VT_CARRAY:  //C style array
		_ASSERTE( false );
		break;

	case VT_USERDEFINED:  //user defined type
		InitVariantPtr( typeInfo, typedesc, value, OUT variant );
		break;

	case VT_PTR:  //pointer type
		InitVariantPtr( typeInfo, *typedesc.lptdesc, value, OUT variant );
		break;

	default:
		_ASSERTE( false );
	}
}

void InitVariantPtr( ITypeInfo* typeInfo, const TYPEDESC& typedesc, v8::Local< v8::Value > value, OUT CComVariant& variant )
{
	switch( typedesc.vt )
	{
	case VT_USERDEFINED:
	{
		CComPtr< ITypeInfo > hrefType;
		typeInfo->GetRefTypeInfo( typedesc.hreftype, &hrefType );

		TypeInfoPtr< TYPEATTR > typeattr( hrefType );
		hrefType->GetTypeAttr( typeattr.Out() );

		if( typeattr->typekind == TKIND_ENUM )
			InitVariantEnum( hrefType, value, OUT variant );
		else
			InitVariantDispatch( hrefType, value, OUT variant );

		break;
	}

	default:
		_ASSERTE( false );
	}
}

void InitVariantEnum( ITypeInfo* typeInfo, v8::Local< v8::Value > value, OUT CComVariant& variant )
{
	int numValue;
	if( value->IsNumber() )
		numValue = value->Int32Value();
	else if( value->IsUndefined() )
		numValue = 0;
	else
		JsException::ThrowMismatch( typeInfo, value );

	variant.vt = VT_I4;
	variant.lVal = numValue;
}

void InitVariantDispatch( ITypeInfo* typeInfo, v8::Local< v8::Value > value, OUT CComVariant& variant )
{
	if( !value->IsObject() ) {
		JsException::ThrowMismatch( typeInfo, value );
	}

	variant.vt = VT_DISPATCH;

	CComPtr< IDispatch > dispatch = GetDispatch( typeInfo, value );
	VERIFY( dispatch.CopyTo( OUT &variant.pdispVal ) );
}

CComPtr< IDispatch > GetDispatch( ITypeInfo* typeInfo, v8::Local< v8::Value > value )
{
	if( value.IsEmpty() || value->IsNull() || value->IsUndefined() )
		return nullptr;

	auto obj = value.As< v8::Object >();
	if( obj->InternalFieldCount() == 1 ) {

		auto interop = InteropInstance::Unwrap< InteropInstance >( obj );
		return interop->instance;
	}

	TypeInfoPtr< TYPEATTR > typeattr( typeInfo );
	auto arrayType = TypeLib::GetInteropType( typeattr.get() );
	if( arrayType->collectionInfo && obj->IsArray() )
	{
		auto arrayCoclass = arrayType->GetCoclass();
		CComPtr< IDispatch > arr = arrayCoclass->CreateInstance();

		auto jsarray = obj.As< v8::Array >();
		for( uint32_t i = 0; i < jsarray->Length(); i++ )
		{
			auto jsitem = jsarray->Get( i );
			CComVariant item;
			InitVariant( arrayCoclass->typeInfo, arrayType->collectionInfo->ItemDesc(), jsitem, OUT item );
			arrayType->collectionInfo->AddItem( arr, item );
		}

		return arr;
	}

	JsException::ThrowMismatch( typeInfo, value );
	return nullptr;
}

v8::Local< v8::Value > VariantToValue( ITypeInfo* typeInfo, const TYPEDESC& typedesc, const CComVariant& variant, v8::Isolate* isolate )
{
	switch( typedesc.vt )
	{
	case VT_I1:  //signed char
		return Nan::New< v8::Number >( variant.cVal );
	case VT_UI1:  //unsigned char
		return Nan::New< v8::Number >( variant.bVal );
	case VT_I2:  //2 byte signed int
		return Nan::New< v8::Number >( variant.iVal );
	case VT_UI2:  //unsigned short
		return Nan::New< v8::Number >( variant.uiVal );
	case VT_I4:  //4 byte signed int
		return Nan::New< v8::Number >( variant.lVal );
	case VT_UI4:  //ULONG
		return Nan::New< v8::Number >( variant.ulVal );
	case VT_I8:  //signed 64-bit int
		return Nan::New< v8::Number >(
				static_cast< double >( variant.ullVal ) );
	case VT_UI8:  //unsigned 64-bit int
		return Nan::New< v8::Number >(
				static_cast< double >( variant.ullVal ) );
	case VT_INT:  //signed machine int
		return Nan::New< v8::Number >( variant.intVal );
	case VT_UINT:  //unsigned machine int
		return Nan::New< v8::Number >( variant.uintVal );
	case VT_R4:  //4 byte real
		return Nan::New< v8::Number >( variant.fltVal );
	case VT_R8:  //8 byte real
		return Nan::New< v8::Number >( variant.dblVal );
	case VT_DATE:  //date
	{
		DATE dateValue = variant.date;

		// JS measures milliseconds. COM measures days. Convert between these two.
		dateValue *= 1000 * 3600 * 24;

		// JS uses Jan 1, 1970 as epoch. COM uses Dec 30, 1899.
		dateValue += 70 * 365 + 20;

		return Nan::New< v8::Date >( dateValue ).ToLocalChecked();
		break;
	}
	case VT_BSTR:  //OLE Automation string
		return Nan::New( ToUTF8( variant.bstrVal ).c_str() ).ToLocalChecked();
	case VT_BOOL:  //True=-1, False=0
		return Nan::New( variant.boolVal != VARIANT_FALSE );
	case VT_DISPATCH:  //IDispatch *
	{
		CComPtr< IDispatch > idisp = variant.pdispVal;
		//CComPtr< IUnknown > iunk;
		//idisp->QueryInterface< IUnknown >( &iunk );
		InteropInstance* instance = new InteropInstance( idisp );
		auto obj = Nan::New< v8::Object >();
		instance->Wrap( obj );
		return obj;
	}
	case VT_UNKNOWN:  //IUnknown *
	{
		CComPtr< IUnknown > iunk = variant.punkVal;
		CComPtr< IDispatch > idisp;
		iunk->QueryInterface< IDispatch >( &idisp );
		InteropInstance* instance = new InteropInstance( idisp );
		auto obj = Nan::New< v8::Object >();
		instance->Wrap( obj );
		return obj;
	}

	case VT_INT_PTR:  //signed machine register size width
	case VT_VARIANT:  //VARIANT *
	case VT_ERROR:  //SCODE
	case VT_DECIMAL:  //16 byte fixed point
	case VT_UINT_PTR:  //unsigned machine register size width
	case VT_HRESULT:  //Standard return type
	case VT_LPSTR:  //null terminated string
	case VT_LPWSTR:  //wide null terminated string
	case VT_CY:  //currency
		_ASSERTE( false );
		return Nan::Undefined();

	case VT_VOID:  //C style void
		return Nan::Undefined();

	case VT_SAFEARRAY:  //(use VT_ARRAY in VARIANT)
	case VT_CARRAY:  //C style array
		_ASSERTE( false );
		return Nan::Undefined();

	case VT_USERDEFINED:  //user defined type
		return UserVariantToValue( typeInfo, typedesc, variant, isolate );

	case VT_PTR:  //pointer type
		return PtrVariantToValue( typeInfo, *( typedesc.lptdesc ), variant, isolate );

	default:
		_ASSERTE( false );
		return Nan::Undefined();
	}
}

v8::Local< v8::Value > UserVariantToValue( ITypeInfo* typeInfo, const TYPEDESC& typedesc, const CComVariant& variant, v8::Isolate* isolate )
{
	_ASSERTE( typedesc.vt == VT_USERDEFINED );
	_ASSERTE( typedesc.hreftype != 0 );

	CComPtr< ITypeInfo > hrefInfo;
	typeInfo->GetRefTypeInfo( typedesc.hreftype, OUT &hrefInfo );

	TypeInfoPtr< TYPEATTR > typeattr( typeInfo );
	hrefInfo->GetTypeAttr( typeattr.Out() );

	GUID typeGuid = typeattr->guid;
	
	// Get the name.
	CComBSTR bstrName;
	hrefInfo->GetDocumentation( MEMBERID_NIL, OUT &bstrName, nullptr, nullptr, nullptr );

	switch( typeattr->typekind )
	{
	case TKIND_ENUM:
		// TODO: Add enum types.
		return Nan::New< v8::Number >( variant.lVal );

	case TKIND_COCLASS:
		_ASSERTE( false );
		return Nan::Undefined();

	case TKIND_INTERFACE:
		_ASSERTE( false );
		return Nan::Undefined();

	case TKIND_DISPATCH:
	{
		std::shared_ptr< InteropType > type = TypeLib::GetInteropType( typeattr.get() );

		v8::Local< v8::Value > argv[ 1 ] = { Nan::New< v8::External >( variant.pdispVal ) };
		v8::Local< v8::Function > cons = Nan::New( type->constructor );
		return cons->NewInstance( Nan::GetCurrentContext(), 1, argv ).ToLocalChecked();
	}

	default:
		_ASSERTE( false );
		return Nan::Undefined();
	}
}

v8::Local< v8::Value > PtrVariantToValue( ITypeInfo* typeInfo, const TYPEDESC& ptrdesc, const CComVariant& variant, v8::Isolate* isolate )
{
	switch( ptrdesc.vt )
	{
	case VT_I1:  //signed char
		return Nan::New< v8::Number >( *variant.pcVal );
	case VT_UI1:  //unsigned char
		return Nan::New< v8::Number >( *variant.pbVal );
	case VT_I2:  //2 byte signed int
		return Nan::New< v8::Number >( *variant.piVal );
	case VT_UI2:  //unsigned short
		return Nan::New< v8::Number >( *variant.puiVal );
	case VT_I4:  //4 byte signed int
		return Nan::New< v8::Number >( *variant.plVal );
	case VT_UI4:  //ULONG
		return Nan::New< v8::Number >( *variant.pulVal );
	case VT_I8:  //signed 64-bit int
		return Nan::New< v8::Number >(
				static_cast< double >( *variant.pullVal ) );
	case VT_UI8:  //unsigned 64-bit int
		return Nan::New< v8::Number >(
				static_cast< double >( *variant.pullVal ) );
	case VT_INT:  //signed machine int
		return Nan::New< v8::Number >( *variant.pintVal );
	case VT_UINT:  //unsigned machine int
		return Nan::New< v8::Number >( *variant.puintVal );
	case VT_R4:  //4 byte real
		return Nan::New< v8::Number >( *variant.pfltVal );
	case VT_R8:  //8 byte real
		return Nan::New< v8::Number >( *variant.pdblVal );
	case VT_DATE:  //date
	{
		DATE dateValue = *variant.pdate;

		// JS measures milliseconds. COM measures days. Convert between these two.
		dateValue *= 1000 * 3600 * 24;

		// JS uses Jan 1, 1970 as epoch. COM uses Dec 30, 1899.
		dateValue += 70 * 365 + 20;

		return Nan::New< v8::Date >( dateValue ).ToLocalChecked();
		break;
	}
	case VT_BSTR:  //OLE Automation string
		return Nan::New( ToUTF8( *variant.pbstrVal ).c_str() ).ToLocalChecked();
	case VT_BOOL:  //True=-1, False=0
		return Nan::New( variant.boolVal != VARIANT_FALSE );
	case VT_DISPATCH:  //IDispatch *
	{
		CComPtr< IDispatch > idisp = *variant.ppdispVal;
		InteropInstance* instance = new InteropInstance( idisp );
		auto obj = Nan::New< v8::Object >();
		instance->Wrap( obj );
		return obj;
	}
	case VT_UNKNOWN:  //IUnknown *
	{
		CComPtr< IUnknown > iunk = *variant.ppunkVal;
		CComPtr< IDispatch > idisp;
		iunk->QueryInterface< IDispatch >( &idisp );
		InteropInstance* instance = new InteropInstance( idisp );
		auto obj = Nan::New< v8::Object >();
		instance->Wrap( obj );
		return obj;
	}

	case VT_INT_PTR:  //signed machine register size width
	case VT_VARIANT:  //VARIANT *
	case VT_ERROR:  //SCODE
	case VT_DECIMAL:  //16 byte fixed point
	case VT_UINT_PTR:  //unsigned machine register size width
	case VT_VOID:  //C style void
	case VT_HRESULT:  //Standard return type
	case VT_LPSTR:  //null terminated string
	case VT_LPWSTR:  //wide null terminated string
	case VT_CY:  //currency
		_ASSERTE( false );
		return Nan::Undefined();

	case VT_SAFEARRAY:  //(use VT_ARRAY in VARIANT)
	case VT_CARRAY:  //C style array
		_ASSERTE( false );
		return Nan::Undefined();

	case VT_USERDEFINED:  //user defined type
		return UserVariantToValue( typeInfo, ptrdesc, variant, isolate );

	case VT_PTR:  //pointer type
		_ASSERTE( false );
		return Nan::Undefined();

	default:
		_ASSERTE( false );
		return Nan::Undefined();
	}
}

void Unwrap( v8::Local< v8::Value > input, OUT IUnknown** output )
{
	_ASSERTE( input->IsObject() );
	InteropInstance* instance = Nan::ObjectWrap::Unwrap< InteropInstance >( Nan::To < v8::Object >( input ).ToLocalChecked() );
	instance->instance->QueryInterface< IUnknown >( OUT output );
}

void Unwrap( v8::Local< v8::Value > input, OUT IDispatch** output )
{
	_ASSERTE( input->IsObject() );
	InteropInstance* instance = Nan::ObjectWrap::Unwrap< InteropInstance >( Nan::To < v8::Object >( input ).ToLocalChecked() );
	instance->instance.CopyTo( output );
}
