// Machine generated IDispatch wrapper class(es) created with ClassWizard
/////////////////////////////////////////////////////////////////////////////
// Shapesexcel wrapper class

class Shapesexcel : public COleDispatchDriver
{
public:
	Shapesexcel() {}		// Calls COleDispatchDriver default constructor
	Shapesexcel(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Shapesexcel(const Shapesexcel& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Item(const VARIANT& Index);
	LPDISPATCH _Default(const VARIANT& Index);
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH AddCallout(long Type, float Left, float Top, float Width, float Height);
	LPDISPATCH AddConnector(long Type, float BeginX, float BeginY, float EndX, float EndY);
	LPDISPATCH AddCurve(const VARIANT& SafeArrayOfPoints);
	LPDISPATCH AddLabel(long Orientation, float Left, float Top, float Width, float Height);
	LPDISPATCH AddLine(float BeginX, float BeginY, float EndX, float EndY);
	LPDISPATCH AddPicture(LPCTSTR Filename, long LinkToFile, long SaveWithDocument, float Left, float Top, float Width, float Height);
	LPDISPATCH AddPolyline(const VARIANT& SafeArrayOfPoints);
	LPDISPATCH AddShape(long Type, float Left, float Top, float Width, float Height);
	LPDISPATCH AddTextEffect(long PresetTextEffect, LPCTSTR Text, LPCTSTR FontName, float FontSize, long FontBold, long FontItalic, float Left, float Top);
	LPDISPATCH AddTextbox(long Orientation, float Left, float Top, float Width, float Height);
	LPDISPATCH BuildFreeform(long EditingType, float X1, float Y1);
	LPDISPATCH GetRange(const VARIANT& Index);
	void SelectAll();
	LPDISPATCH AddFormControl(long Type, long Left, long Top, long Width, long Height);
	LPDISPATCH AddOLEObject(const VARIANT& ClassType, const VARIANT& Filename, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel, const VARIANT& Left, 
		const VARIANT& Top, const VARIANT& Width, const VARIANT& Height);
	LPDISPATCH AddDiagram(long Type, float Left, float Top, float Width, float Height);
};
