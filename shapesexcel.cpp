// Machine generated IDispatch wrapper class(es) created with ClassWizard

#include "stdafx.h"
#include "shapesexcel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



/////////////////////////////////////////////////////////////////////////////
// Shapesexcel properties

/////////////////////////////////////////////////////////////////////////////
// Shapesexcel operations

LPDISPATCH Shapesexcel::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long Shapesexcel::GetCreator()
{
	long result;
	InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH Shapesexcel::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long Shapesexcel::GetCount()
{
	long result;
	InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH Shapesexcel::Item(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xaa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

LPDISPATCH Shapesexcel::_Default(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

LPUNKNOWN Shapesexcel::Get_NewEnum()
{
	LPUNKNOWN result;
	InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
	return result;
}

LPDISPATCH Shapesexcel::AddCallout(long Type, float Left, float Top, float Width, float Height)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4;
	InvokeHelper(0x6b1, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Type, Left, Top, Width, Height);
	return result;
}

LPDISPATCH Shapesexcel::AddConnector(long Type, float BeginX, float BeginY, float EndX, float EndY)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4;
	InvokeHelper(0x6b2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Type, BeginX, BeginY, EndX, EndY);
	return result;
}

LPDISPATCH Shapesexcel::AddCurve(const VARIANT& SafeArrayOfPoints)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x6b7, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&SafeArrayOfPoints);
	return result;
}

LPDISPATCH Shapesexcel::AddLabel(long Orientation, float Left, float Top, float Width, float Height)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4;
	InvokeHelper(0x6b9, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Orientation, Left, Top, Width, Height);
	return result;
}

LPDISPATCH Shapesexcel::AddLine(float BeginX, float BeginY, float EndX, float EndY)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_R4 VTS_R4 VTS_R4 VTS_R4;
	InvokeHelper(0x6ba, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		BeginX, BeginY, EndX, EndY);
	return result;
}

LPDISPATCH Shapesexcel::AddPicture(LPCTSTR Filename, long LinkToFile, long SaveWithDocument, float Left, float Top, float Width, float Height)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_I4 VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4;
	InvokeHelper(0x6bb, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height);
	return result;
}

LPDISPATCH Shapesexcel::AddPolyline(const VARIANT& SafeArrayOfPoints)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x6be, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&SafeArrayOfPoints);
	return result;
}

LPDISPATCH Shapesexcel::AddShape(long Type, float Left, float Top, float Width, float Height)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4;
	InvokeHelper(0x6bf, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Type, Left, Top, Width, Height);
	return result;
}

LPDISPATCH Shapesexcel::AddTextEffect(long PresetTextEffect, LPCTSTR Text, LPCTSTR FontName, float FontSize, long FontBold, long FontItalic, float Left, float Top)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_BSTR VTS_BSTR VTS_R4 VTS_I4 VTS_I4 VTS_R4 VTS_R4;
	InvokeHelper(0x6c0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top);
	return result;
}

LPDISPATCH Shapesexcel::AddTextbox(long Orientation, float Left, float Top, float Width, float Height)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4;
	InvokeHelper(0x6c6, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Orientation, Left, Top, Width, Height);
	return result;
}

LPDISPATCH Shapesexcel::BuildFreeform(long EditingType, float X1, float Y1)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_R4 VTS_R4;
	InvokeHelper(0x6c7, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		EditingType, X1, Y1);
	return result;
}

LPDISPATCH Shapesexcel::GetRange(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

void Shapesexcel::SelectAll()
{
	InvokeHelper(0x6c9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH Shapesexcel::AddFormControl(long Type, long Left, long Top, long Width, long Height)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0x6ca, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Type, Left, Top, Width, Height);
	return result;
}

LPDISPATCH Shapesexcel::AddOLEObject(const VARIANT& ClassType, const VARIANT& Filename, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel, const VARIANT& Left, 
		const VARIANT& Top, const VARIANT& Width, const VARIANT& Height)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x6cb, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&ClassType, &Filename, &Link, &DisplayAsIcon, &IconFileName, &IconIndex, &IconLabel, &Left, &Top, &Width, &Height);
	return result;
}

LPDISPATCH Shapesexcel::AddDiagram(long Type, float Left, float Top, float Width, float Height)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4;
	InvokeHelper(0x880, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Type, Left, Top, Width, Height);
	return result;
}
