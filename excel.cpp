// Machine generated IDispatch wrapper class(es) created with ClassWizard

#include "stdafx.h"
#include "excel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



/////////////////////////////////////////////////////////////////////////////
// _Worksheet properties

/////////////////////////////////////////////////////////////////////////////
// _Worksheet operations

LPDISPATCH _Worksheet::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Worksheet::GetCreator()
{
	long result;
	InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Worksheet::Activate()
{
	InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Worksheet::Copy(const VARIANT& Before, const VARIANT& After)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x227, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Before, &After);
}

void _Worksheet::Delete()
{
	InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CString _Worksheet::GetCodeName()
{
	CString result;
	InvokeHelper(0x55d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString _Worksheet::Get_CodeName()
{
	CString result;
	InvokeHelper(0x80010000, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Worksheet::Set_CodeName(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x80010000, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

long _Worksheet::GetIndex()
{
	long result;
	InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Worksheet::Move(const VARIANT& Before, const VARIANT& After)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x27d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Before, &After);
}

CString _Worksheet::GetName()
{
	CString result;
	InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetName(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x6e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

LPDISPATCH _Worksheet::GetNext()
{
	LPDISPATCH result;
	InvokeHelper(0x1f6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetPageSetup()
{
	LPDISPATCH result;
	InvokeHelper(0x3e6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetPrevious()
{
	LPDISPATCH result;
	InvokeHelper(0x1f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Worksheet::PrintPreview(const VARIANT& EnableChanges)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x119, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &EnableChanges);
}

BOOL _Worksheet::GetProtectContents()
{
	BOOL result;
	InvokeHelper(0x124, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Worksheet::GetProtectDrawingObjects()
{
	BOOL result;
	InvokeHelper(0x125, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Worksheet::GetProtectionMode()
{
	BOOL result;
	InvokeHelper(0x487, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Worksheet::GetProtectScenarios()
{
	BOOL result;
	InvokeHelper(0x126, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::Select(const VARIANT& Replace)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xeb, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Replace);
}

void _Worksheet::Unprotect(const VARIANT& Password)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x11d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Password);
}

long _Worksheet::GetVisible()
{
	long result;
	InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetVisible(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH _Worksheet::GetShapes()
{
	LPDISPATCH result;
	InvokeHelper(0x561, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Worksheet::GetTransitionExpEval()
{
	BOOL result;
	InvokeHelper(0x191, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetTransitionExpEval(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x191, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Worksheet::GetAutoFilterMode()
{
	BOOL result;
	InvokeHelper(0x318, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetAutoFilterMode(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x318, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Worksheet::SetBackgroundPicture(LPCTSTR Filename)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x4a4, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Filename);
}

void _Worksheet::Calculate()
{
	InvokeHelper(0x117, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

BOOL _Worksheet::GetEnableCalculation()
{
	BOOL result;
	InvokeHelper(0x590, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetEnableCalculation(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x590, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Worksheet::GetCells()
{
	LPDISPATCH result;
	InvokeHelper(0xee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::ChartObjects(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x424, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

void _Worksheet::CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& SpellLang)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x1f9, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &SpellLang);
}

LPDISPATCH _Worksheet::GetCircularReference()
{
	LPDISPATCH result;
	InvokeHelper(0x42d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Worksheet::ClearArrows()
{
	InvokeHelper(0x3ca, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH _Worksheet::GetColumns()
{
	LPDISPATCH result;
	InvokeHelper(0xf1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Worksheet::GetConsolidationFunction()
{
	long result;
	InvokeHelper(0x315, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

VARIANT _Worksheet::GetConsolidationOptions()
{
	VARIANT result;
	InvokeHelper(0x316, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

VARIANT _Worksheet::GetConsolidationSources()
{
	VARIANT result;
	InvokeHelper(0x317, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

BOOL _Worksheet::GetEnableAutoFilter()
{
	BOOL result;
	InvokeHelper(0x484, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetEnableAutoFilter(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x484, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Worksheet::GetEnableSelection()
{
	long result;
	InvokeHelper(0x591, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetEnableSelection(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x591, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Worksheet::GetEnableOutlining()
{
	BOOL result;
	InvokeHelper(0x485, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetEnableOutlining(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x485, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Worksheet::GetEnablePivotTable()
{
	BOOL result;
	InvokeHelper(0x486, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetEnablePivotTable(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x486, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

VARIANT _Worksheet::Evaluate(const VARIANT& Name)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x1, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&Name);
	return result;
}

VARIANT _Worksheet::_Evaluate(const VARIANT& Name)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xfffffffb, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&Name);
	return result;
}

BOOL _Worksheet::GetFilterMode()
{
	BOOL result;
	InvokeHelper(0x320, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::ResetAllPageBreaks()
{
	InvokeHelper(0x592, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH _Worksheet::GetNames()
{
	LPDISPATCH result;
	InvokeHelper(0x1ba, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::OLEObjects(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x31f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

LPDISPATCH _Worksheet::GetOutline()
{
	LPDISPATCH result;
	InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Worksheet::Paste(const VARIANT& Destination, const VARIANT& Link)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0xd3, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Destination, &Link);
}

LPDISPATCH _Worksheet::PivotTables(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x2b2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

LPDISPATCH _Worksheet::PivotTableWizard(const VARIANT& SourceType, const VARIANT& SourceData, const VARIANT& TableDestination, const VARIANT& TableName, const VARIANT& RowGrand, const VARIANT& ColumnGrand, const VARIANT& SaveData, 
		const VARIANT& HasAutoFormat, const VARIANT& AutoPage, const VARIANT& Reserved, const VARIANT& BackgroundQuery, const VARIANT& OptimizeCache, const VARIANT& PageFieldOrder, const VARIANT& PageFieldWrapCount, const VARIANT& ReadData, 
		const VARIANT& Connection)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x2ac, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&SourceType, &SourceData, &TableDestination, &TableName, &RowGrand, &ColumnGrand, &SaveData, &HasAutoFormat, &AutoPage, &Reserved, &BackgroundQuery, &OptimizeCache, &PageFieldOrder, &PageFieldWrapCount, &ReadData, &Connection);
	return result;
}

LPDISPATCH _Worksheet::GetRange(const VARIANT& Cell1, const VARIANT& Cell2)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&Cell1, &Cell2);
	return result;
}

LPDISPATCH _Worksheet::GetRows()
{
	LPDISPATCH result;
	InvokeHelper(0x102, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::Scenarios(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x38c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

CString _Worksheet::GetScrollArea()
{
	CString result;
	InvokeHelper(0x599, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetScrollArea(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x599, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

void _Worksheet::ShowAllData()
{
	InvokeHelper(0x31a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Worksheet::ShowDataForm()
{
	InvokeHelper(0x199, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

double _Worksheet::GetStandardHeight()
{
	double result;
	InvokeHelper(0x197, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

double _Worksheet::GetStandardWidth()
{
	double result;
	InvokeHelper(0x198, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetStandardWidth(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x198, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

BOOL _Worksheet::GetTransitionFormEntry()
{
	BOOL result;
	InvokeHelper(0x192, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetTransitionFormEntry(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x192, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Worksheet::GetType()
{
	long result;
	InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetUsedRange()
{
	LPDISPATCH result;
	InvokeHelper(0x19c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetHPageBreaks()
{
	LPDISPATCH result;
	InvokeHelper(0x58a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetVPageBreaks()
{
	LPDISPATCH result;
	InvokeHelper(0x58b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetQueryTables()
{
	LPDISPATCH result;
	InvokeHelper(0x59a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Worksheet::GetDisplayPageBreaks()
{
	BOOL result;
	InvokeHelper(0x59b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetDisplayPageBreaks(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x59b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Worksheet::GetComments()
{
	LPDISPATCH result;
	InvokeHelper(0x23f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetHyperlinks()
{
	LPDISPATCH result;
	InvokeHelper(0x571, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Worksheet::ClearCircles()
{
	InvokeHelper(0x59c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Worksheet::CircleInvalid()
{
	InvokeHelper(0x59d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH _Worksheet::GetAutoFilter()
{
	LPDISPATCH result;
	InvokeHelper(0x319, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Worksheet::GetDisplayRightToLeft()
{
	BOOL result;
	InvokeHelper(0x6ee, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Worksheet::SetDisplayRightToLeft(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x6ee, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Worksheet::GetScripts()
{
	LPDISPATCH result;
	InvokeHelper(0x718, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Worksheet::PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate, const VARIANT& PrToFileName)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x6ec, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName);
}

LPDISPATCH _Worksheet::GetTab()
{
	LPDISPATCH result;
	InvokeHelper(0x411, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetMailEnvelope()
{
	LPDISPATCH result;
	InvokeHelper(0x7e5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Worksheet::SaveAs(LPCTSTR Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& AddToMru, const VARIANT& TextCodepage, 
		const VARIANT& TextVisualLayout, const VARIANT& Local)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x785, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &AddToMru, &TextCodepage, &TextVisualLayout, &Local);
}

LPDISPATCH _Worksheet::GetCustomProperties()
{
	LPDISPATCH result;
	InvokeHelper(0x7ee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetSmartTags()
{
	LPDISPATCH result;
	InvokeHelper(0x7e0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::GetProtection()
{
	LPDISPATCH result;
	InvokeHelper(0xb0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Worksheet::PasteSpecial(const VARIANT& Format, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel, const VARIANT& NoHTMLFormatting)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x788, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Format, &Link, &DisplayAsIcon, &IconFileName, &IconIndex, &IconLabel, &NoHTMLFormatting);
}

void _Worksheet::Protect(const VARIANT& Password, const VARIANT& DrawingObjects, const VARIANT& Contents, const VARIANT& Scenarios, const VARIANT& UserInterfaceOnly, const VARIANT& AllowFormattingCells, const VARIANT& AllowFormattingColumns, 
		const VARIANT& AllowFormattingRows, const VARIANT& AllowInsertingColumns, const VARIANT& AllowInsertingRows, const VARIANT& AllowInsertingHyperlinks, const VARIANT& AllowDeletingColumns, const VARIANT& AllowDeletingRows, 
		const VARIANT& AllowSorting, const VARIANT& AllowFiltering, const VARIANT& AllowUsingPivotTables)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x7ed, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Password, &DrawingObjects, &Contents, &Scenarios, &UserInterfaceOnly, &AllowFormattingCells, &AllowFormattingColumns, &AllowFormattingRows, &AllowInsertingColumns, &AllowInsertingRows, &AllowInsertingHyperlinks, &AllowDeletingColumns, 
		&AllowDeletingRows, &AllowSorting, &AllowFiltering, &AllowUsingPivotTables);
}

LPDISPATCH _Worksheet::GetListObjects()
{
	LPDISPATCH result;
	InvokeHelper(0x8d3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Worksheet::XmlDataQuery(LPCTSTR XPath, const VARIANT& SelectionNamespaces, const VARIANT& Map)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x8d4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		XPath, &SelectionNamespaces, &Map);
	return result;
}

LPDISPATCH _Worksheet::XmlMapQuery(LPCTSTR XPath, const VARIANT& SelectionNamespaces, const VARIANT& Map)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x8d7, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		XPath, &SelectionNamespaces, &Map);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// _Workbook properties

/////////////////////////////////////////////////////////////////////////////
// _Workbook operations

LPDISPATCH _Workbook::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Workbook::GetCreator()
{
	long result;
	InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetAcceptLabelsInFormulas()
{
	BOOL result;
	InvokeHelper(0x5a1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetAcceptLabelsInFormulas(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5a1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Workbook::Activate()
{
	InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH _Workbook::GetActiveChart()
{
	LPDISPATCH result;
	InvokeHelper(0xb7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetActiveSheet()
{
	LPDISPATCH result;
	InvokeHelper(0x133, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Workbook::GetAutoUpdateFrequency()
{
	long result;
	InvokeHelper(0x5a2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Workbook::SetAutoUpdateFrequency(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x5a2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Workbook::GetAutoUpdateSaveChanges()
{
	BOOL result;
	InvokeHelper(0x5a3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetAutoUpdateSaveChanges(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5a3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Workbook::GetChangeHistoryDuration()
{
	long result;
	InvokeHelper(0x5a4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Workbook::SetChangeHistoryDuration(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x5a4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH _Workbook::GetBuiltinDocumentProperties()
{
	LPDISPATCH result;
	InvokeHelper(0x498, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::ChangeFileAccess(long Mode, const VARIANT& WritePassword, const VARIANT& Notify)
{
	static BYTE parms[] =
		VTS_I4 VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x3dd, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Mode, &WritePassword, &Notify);
}

void _Workbook::ChangeLink(LPCTSTR Name, LPCTSTR NewName, long Type)
{
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR VTS_I4;
	InvokeHelper(0x322, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Name, NewName, Type);
}

LPDISPATCH _Workbook::GetCharts()
{
	LPDISPATCH result;
	InvokeHelper(0x79, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::Close(const VARIANT& SaveChanges, const VARIANT& Filename, const VARIANT& RouteWorkbook)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x115, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &SaveChanges, &Filename, &RouteWorkbook);
}

CString _Workbook::GetCodeName()
{
	CString result;
	InvokeHelper(0x55d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString _Workbook::Get_CodeName()
{
	CString result;
	InvokeHelper(0x80010000, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Workbook::Set_CodeName(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x80010000, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

VARIANT _Workbook::GetColors(const VARIANT& Index)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x11e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		&Index);
	return result;
}

void _Workbook::SetColors(const VARIANT& Index, const VARIANT& newValue)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x11e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 &Index, &newValue);
}

LPDISPATCH _Workbook::GetCommandBars()
{
	LPDISPATCH result;
	InvokeHelper(0x59f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Workbook::GetConflictResolution()
{
	long result;
	InvokeHelper(0x497, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Workbook::SetConflictResolution(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x497, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH _Workbook::GetContainer()
{
	LPDISPATCH result;
	InvokeHelper(0x4a6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetCreateBackup()
{
	BOOL result;
	InvokeHelper(0x11f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetCustomDocumentProperties()
{
	LPDISPATCH result;
	InvokeHelper(0x499, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetDate1904()
{
	BOOL result;
	InvokeHelper(0x193, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetDate1904(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x193, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Workbook::DeleteNumberFormat(LPCTSTR NumberFormat)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x18d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 NumberFormat);
}

long _Workbook::GetDisplayDrawingObjects()
{
	long result;
	InvokeHelper(0x194, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Workbook::SetDisplayDrawingObjects(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x194, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Workbook::ExclusiveAccess()
{
	BOOL result;
	InvokeHelper(0x490, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
	return result;
}

long _Workbook::GetFileFormat()
{
	long result;
	InvokeHelper(0x120, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Workbook::ForwardMailer()
{
	InvokeHelper(0x3cd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CString _Workbook::GetFullName()
{
	CString result;
	InvokeHelper(0x121, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetHasPassword()
{
	BOOL result;
	InvokeHelper(0x122, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetHasRoutingSlip()
{
	BOOL result;
	InvokeHelper(0x3b6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetHasRoutingSlip(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3b6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Workbook::GetIsAddin()
{
	BOOL result;
	InvokeHelper(0x5a5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetIsAddin(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5a5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

VARIANT _Workbook::LinkInfo(LPCTSTR Name, long LinkInfo, const VARIANT& Type, const VARIANT& EditionRef)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_BSTR VTS_I4 VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x327, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		Name, LinkInfo, &Type, &EditionRef);
	return result;
}

VARIANT _Workbook::LinkSources(const VARIANT& Type)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x328, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&Type);
	return result;
}

LPDISPATCH _Workbook::GetMailer()
{
	LPDISPATCH result;
	InvokeHelper(0x3d3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::MergeWorkbook(const VARIANT& Filename)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x5a6, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Filename);
}

BOOL _Workbook::GetMultiUserEditing()
{
	BOOL result;
	InvokeHelper(0x491, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

CString _Workbook::GetName()
{
	CString result;
	InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetNames()
{
	LPDISPATCH result;
	InvokeHelper(0x1ba, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::NewWindow()
{
	LPDISPATCH result;
	InvokeHelper(0x118, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::OpenLinks(LPCTSTR Name, const VARIANT& ReadOnly, const VARIANT& Type)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x323, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Name, &ReadOnly, &Type);
}

CString _Workbook::GetPath()
{
	CString result;
	InvokeHelper(0x123, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetPersonalViewListSettings()
{
	BOOL result;
	InvokeHelper(0x5a7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetPersonalViewListSettings(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5a7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Workbook::GetPersonalViewPrintSettings()
{
	BOOL result;
	InvokeHelper(0x5a8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetPersonalViewPrintSettings(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5a8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Workbook::PivotCaches()
{
	LPDISPATCH result;
	InvokeHelper(0x5a9, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::Post(const VARIANT& DestName)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x48e, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &DestName);
}

BOOL _Workbook::GetPrecisionAsDisplayed()
{
	BOOL result;
	InvokeHelper(0x195, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetPrecisionAsDisplayed(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x195, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Workbook::PrintPreview(const VARIANT& EnableChanges)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x119, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &EnableChanges);
}

void _Workbook::ProtectSharing(const VARIANT& Filename, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& SharingPassword)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x5aa, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Filename, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &SharingPassword);
}

BOOL _Workbook::GetProtectStructure()
{
	BOOL result;
	InvokeHelper(0x24c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetProtectWindows()
{
	BOOL result;
	InvokeHelper(0x127, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetReadOnly()
{
	BOOL result;
	InvokeHelper(0x128, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::RefreshAll()
{
	InvokeHelper(0x5ac, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Workbook::Reply()
{
	InvokeHelper(0x3d1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Workbook::ReplyAll()
{
	InvokeHelper(0x3d2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Workbook::RemoveUser(long Index)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x5ad, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Index);
}

long _Workbook::GetRevisionNumber()
{
	long result;
	InvokeHelper(0x494, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Workbook::Route()
{
	InvokeHelper(0x3b2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

BOOL _Workbook::GetRouted()
{
	BOOL result;
	InvokeHelper(0x3b7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetRoutingSlip()
{
	LPDISPATCH result;
	InvokeHelper(0x3b5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::RunAutoMacros(long Which)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x27a, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Which);
}

void _Workbook::Save()
{
	InvokeHelper(0x11b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Workbook::SaveCopyAs(const VARIANT& Filename)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xaf, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Filename);
}

BOOL _Workbook::GetSaved()
{
	BOOL result;
	InvokeHelper(0x12a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetSaved(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x12a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Workbook::GetSaveLinkValues()
{
	BOOL result;
	InvokeHelper(0x196, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetSaveLinkValues(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x196, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Workbook::SendMail(const VARIANT& Recipients, const VARIANT& Subject, const VARIANT& ReturnReceipt)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x3b3, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Recipients, &Subject, &ReturnReceipt);
}

void _Workbook::SendMailer(const VARIANT& FileFormat, long Priority)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_I4;
	InvokeHelper(0x3d4, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &FileFormat, Priority);
}

void _Workbook::SetLinkOnData(LPCTSTR Name, const VARIANT& Procedure)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT;
	InvokeHelper(0x329, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Name, &Procedure);
}

LPDISPATCH _Workbook::GetSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x1e5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetShowConflictHistory()
{
	BOOL result;
	InvokeHelper(0x493, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetShowConflictHistory(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x493, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Workbook::GetStyles()
{
	LPDISPATCH result;
	InvokeHelper(0x1ed, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::Unprotect(const VARIANT& Password)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x11d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Password);
}

void _Workbook::UnprotectSharing(const VARIANT& SharingPassword)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x5af, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &SharingPassword);
}

void _Workbook::UpdateFromFile()
{
	InvokeHelper(0x3e3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Workbook::UpdateLink(const VARIANT& Name, const VARIANT& Type)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x324, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Name, &Type);
}

BOOL _Workbook::GetUpdateRemoteReferences()
{
	BOOL result;
	InvokeHelper(0x19b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetUpdateRemoteReferences(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x19b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

VARIANT _Workbook::GetUserStatus()
{
	VARIANT result;
	InvokeHelper(0x495, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetCustomViews()
{
	LPDISPATCH result;
	InvokeHelper(0x5b0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetWindows()
{
	LPDISPATCH result;
	InvokeHelper(0x1ae, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetWorksheets()
{
	LPDISPATCH result;
	InvokeHelper(0x1ee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetWriteReserved()
{
	BOOL result;
	InvokeHelper(0x12b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

CString _Workbook::GetWriteReservedBy()
{
	CString result;
	InvokeHelper(0x12c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetExcel4IntlMacroSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x245, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetExcel4MacroSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x243, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetTemplateRemoveExtData()
{
	BOOL result;
	InvokeHelper(0x5b1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetTemplateRemoveExtData(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5b1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Workbook::HighlightChangesOptions(const VARIANT& When, const VARIANT& Who, const VARIANT& Where)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x5b2, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &When, &Who, &Where);
}

BOOL _Workbook::GetHighlightChangesOnScreen()
{
	BOOL result;
	InvokeHelper(0x5b5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetHighlightChangesOnScreen(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5b5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Workbook::GetKeepChangeHistory()
{
	BOOL result;
	InvokeHelper(0x5b6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetKeepChangeHistory(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5b6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Workbook::GetListChangesOnNewSheet()
{
	BOOL result;
	InvokeHelper(0x5b7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetListChangesOnNewSheet(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5b7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Workbook::PurgeChangeHistoryNow(long Days, const VARIANT& SharingPassword)
{
	static BYTE parms[] =
		VTS_I4 VTS_VARIANT;
	InvokeHelper(0x5b8, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Days, &SharingPassword);
}

void _Workbook::AcceptAllChanges(const VARIANT& When, const VARIANT& Who, const VARIANT& Where)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x5ba, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &When, &Who, &Where);
}

void _Workbook::RejectAllChanges(const VARIANT& When, const VARIANT& Who, const VARIANT& Where)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x5bb, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &When, &Who, &Where);
}

void _Workbook::ResetColors()
{
	InvokeHelper(0x5bc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH _Workbook::GetVBProject()
{
	LPDISPATCH result;
	InvokeHelper(0x5bd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::FollowHyperlink(LPCTSTR Address, const VARIANT& SubAddress, const VARIANT& NewWindow, const VARIANT& AddHistory, const VARIANT& ExtraInfo, const VARIANT& Method, const VARIANT& HeaderInfo)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x5be, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Address, &SubAddress, &NewWindow, &AddHistory, &ExtraInfo, &Method, &HeaderInfo);
}

void _Workbook::AddToFavorites()
{
	InvokeHelper(0x5c4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

BOOL _Workbook::GetIsInplace()
{
	BOOL result;
	InvokeHelper(0x6e9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate, const VARIANT& PrToFileName)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x6ec, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName);
}

void _Workbook::WebPagePreview()
{
	InvokeHelper(0x71a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH _Workbook::GetPublishObjects()
{
	LPDISPATCH result;
	InvokeHelper(0x71b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetWebOptions()
{
	LPDISPATCH result;
	InvokeHelper(0x71c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::ReloadAs(long Encoding)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x71d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Encoding);
}

LPDISPATCH _Workbook::GetHTMLProject()
{
	LPDISPATCH result;
	InvokeHelper(0x71f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetEnvelopeVisible()
{
	BOOL result;
	InvokeHelper(0x720, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetEnvelopeVisible(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x720, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Workbook::GetCalculationVersion()
{
	long result;
	InvokeHelper(0x70e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetVBASigned()
{
	BOOL result;
	InvokeHelper(0x724, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetShowPivotTableFieldList()
{
	BOOL result;
	InvokeHelper(0x7fe, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetShowPivotTableFieldList(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x7fe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Workbook::GetUpdateLinks()
{
	long result;
	InvokeHelper(0x360, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Workbook::SetUpdateLinks(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x360, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

void _Workbook::BreakLink(LPCTSTR Name, long Type)
{
	static BYTE parms[] =
		VTS_BSTR VTS_I4;
	InvokeHelper(0x7ff, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Name, Type);
}

void _Workbook::SaveAs(const VARIANT& Filename, const VARIANT& FileFormat, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, long AccessMode, const VARIANT& ConflictResolution, 
		const VARIANT& AddToMru, const VARIANT& TextCodepage, const VARIANT& TextVisualLayout, const VARIANT& Local)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x785, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, AccessMode, &ConflictResolution, &AddToMru, &TextCodepage, &TextVisualLayout, &Local);
}

BOOL _Workbook::GetEnableAutoRecover()
{
	BOOL result;
	InvokeHelper(0x801, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetEnableAutoRecover(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x801, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Workbook::GetRemovePersonalInformation()
{
	BOOL result;
	InvokeHelper(0x802, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetRemovePersonalInformation(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x802, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

CString _Workbook::GetFullNameURLEncoded()
{
	CString result;
	InvokeHelper(0x787, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Workbook::CheckIn(const VARIANT& SaveChanges, const VARIANT& Comments, const VARIANT& MakePublic)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x803, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &SaveChanges, &Comments, &MakePublic);
}

BOOL _Workbook::CanCheckIn()
{
	BOOL result;
	InvokeHelper(0x805, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SendForReview(const VARIANT& Recipients, const VARIANT& Subject, const VARIANT& ShowMessage, const VARIANT& IncludeAttachment)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x806, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Recipients, &Subject, &ShowMessage, &IncludeAttachment);
}

void _Workbook::ReplyWithChanges(const VARIANT& ShowMessage)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x809, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &ShowMessage);
}

void _Workbook::EndReview()
{
	InvokeHelper(0x80a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

CString _Workbook::GetPassword()
{
	CString result;
	InvokeHelper(0x1ad, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Workbook::SetPassword(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x1ad, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString _Workbook::GetWritePassword()
{
	CString result;
	InvokeHelper(0x468, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Workbook::SetWritePassword(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x468, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString _Workbook::GetPasswordEncryptionProvider()
{
	CString result;
	InvokeHelper(0x80b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString _Workbook::GetPasswordEncryptionAlgorithm()
{
	CString result;
	InvokeHelper(0x80c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

long _Workbook::GetPasswordEncryptionKeyLength()
{
	long result;
	InvokeHelper(0x80d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Workbook::SetPasswordEncryptionOptions(const VARIANT& PasswordEncryptionProvider, const VARIANT& PasswordEncryptionAlgorithm, const VARIANT& PasswordEncryptionKeyLength, const VARIANT& PasswordEncryptionFileProperties)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x80e, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &PasswordEncryptionProvider, &PasswordEncryptionAlgorithm, &PasswordEncryptionKeyLength, &PasswordEncryptionFileProperties);
}

BOOL _Workbook::GetPasswordEncryptionFileProperties()
{
	BOOL result;
	InvokeHelper(0x80f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetReadOnlyRecommended()
{
	BOOL result;
	InvokeHelper(0x7d5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetReadOnlyRecommended(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x7d5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Workbook::Protect(const VARIANT& Password, const VARIANT& Structure, const VARIANT& Windows)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x7ed, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Password, &Structure, &Windows);
}

LPDISPATCH _Workbook::GetSmartTagOptions()
{
	LPDISPATCH result;
	InvokeHelper(0x810, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::RecheckSmartTags()
{
	InvokeHelper(0x811, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH _Workbook::GetPermission()
{
	LPDISPATCH result;
	InvokeHelper(0x8d8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetSharedWorkspace()
{
	LPDISPATCH result;
	InvokeHelper(0x8d9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetSync()
{
	LPDISPATCH result;
	InvokeHelper(0x8da, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Workbook::SendFaxOverInternet(const VARIANT& Recipients, const VARIANT& Subject, const VARIANT& ShowMessage)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x8db, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Recipients, &Subject, &ShowMessage);
}

LPDISPATCH _Workbook::GetXmlNamespaces()
{
	LPDISPATCH result;
	InvokeHelper(0x8dc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetXmlMaps()
{
	LPDISPATCH result;
	InvokeHelper(0x8dd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Workbook::XmlImport(LPCTSTR Url, LPDISPATCH* ImportMap, const VARIANT& Overwrite, const VARIANT& Destination)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_PDISPATCH VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x8de, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Url, ImportMap, &Overwrite, &Destination);
	return result;
}

LPDISPATCH _Workbook::GetSmartDocument()
{
	LPDISPATCH result;
	InvokeHelper(0x8e1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Workbook::GetDocumentLibraryVersions()
{
	LPDISPATCH result;
	InvokeHelper(0x8e2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Workbook::GetInactiveListBorderVisible()
{
	BOOL result;
	InvokeHelper(0x8e3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetInactiveListBorderVisible(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x8e3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Workbook::GetDisplayInkComments()
{
	BOOL result;
	InvokeHelper(0x8e4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Workbook::SetDisplayInkComments(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x8e4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Workbook::XmlImportXml(LPCTSTR Data, LPDISPATCH* ImportMap, const VARIANT& Overwrite, const VARIANT& Destination)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_PDISPATCH VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x8e5, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Data, ImportMap, &Overwrite, &Destination);
	return result;
}

void _Workbook::SaveAsXMLData(LPCTSTR Filename, LPDISPATCH Map)
{
	static BYTE parms[] =
		VTS_BSTR VTS_DISPATCH;
	InvokeHelper(0x8e6, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Filename, Map);
}

void _Workbook::ToggleFormsDesign()
{
	InvokeHelper(0x8e7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// Workbooks properties

/////////////////////////////////////////////////////////////////////////////
// Workbooks operations

LPDISPATCH Workbooks::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long Workbooks::GetCreator()
{
	long result;
	InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH Workbooks::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH Workbooks::Add(const VARIANT& Template)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xb5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Template);
	return result;
}

void Workbooks::Close()
{
	InvokeHelper(0x115, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

long Workbooks::GetCount()
{
	long result;
	InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH Workbooks::GetItem(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

LPUNKNOWN Workbooks::Get_NewEnum()
{
	LPUNKNOWN result;
	InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
	return result;
}

LPDISPATCH Workbooks::Get_Default(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

LPDISPATCH Workbooks::Open(LPCTSTR Filename, const VARIANT& UpdateLinks, const VARIANT& ReadOnly, const VARIANT& Format, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& IgnoreReadOnlyRecommended, const VARIANT& Origin, 
		const VARIANT& Delimiter, const VARIANT& Editable, const VARIANT& Notify, const VARIANT& Converter, const VARIANT& AddToMru, const VARIANT& Local, const VARIANT& CorruptLoad)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x783, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Filename, &UpdateLinks, &ReadOnly, &Format, &Password, &WriteResPassword, &IgnoreReadOnlyRecommended, &Origin, &Delimiter, &Editable, &Notify, &Converter, &AddToMru, &Local, &CorruptLoad);
	return result;
}

void Workbooks::OpenText(LPCTSTR Filename, const VARIANT& Origin, const VARIANT& StartRow, const VARIANT& DataType, long TextQualifier, const VARIANT& ConsecutiveDelimiter, const VARIANT& Tab, const VARIANT& Semicolon, const VARIANT& Comma, 
		const VARIANT& Space, const VARIANT& Other, const VARIANT& OtherChar, const VARIANT& FieldInfo, const VARIANT& TextVisualLayout, const VARIANT& DecimalSeparator, const VARIANT& ThousandsSeparator, const VARIANT& TrailingMinusNumbers, 
		const VARIANT& Local)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x784, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Filename, &Origin, &StartRow, &DataType, TextQualifier, &ConsecutiveDelimiter, &Tab, &Semicolon, &Comma, &Space, &Other, &OtherChar, &FieldInfo, &TextVisualLayout, &DecimalSeparator, &ThousandsSeparator, &TrailingMinusNumbers, &Local);
}

LPDISPATCH Workbooks::OpenDatabase(LPCTSTR Filename, const VARIANT& CommandText, const VARIANT& CommandType, const VARIANT& BackgroundQuery, const VARIANT& ImportDataAs)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x813, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Filename, &CommandText, &CommandType, &BackgroundQuery, &ImportDataAs);
	return result;
}

void Workbooks::CheckOut(LPCTSTR Filename)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x815, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Filename);
}

BOOL Workbooks::CanCheckOut(LPCTSTR Filename)
{
	BOOL result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x816, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms,
		Filename);
	return result;
}

LPDISPATCH Workbooks::OpenXML(LPCTSTR Filename, const VARIANT& Stylesheets, const VARIANT& LoadOption)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x8e8, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Filename, &Stylesheets, &LoadOption);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// Worksheets properties

/////////////////////////////////////////////////////////////////////////////
// Worksheets operations

LPDISPATCH Worksheets::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long Worksheets::GetCreator()
{
	long result;
	InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH Worksheets::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH Worksheets::Add(const VARIANT& Before, const VARIANT& After, const VARIANT& Count, const VARIANT& Type)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0xb5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Before, &After, &Count, &Type);
	return result;
}

void Worksheets::Copy(const VARIANT& Before, const VARIANT& After)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x227, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Before, &After);
}

long Worksheets::GetCount()
{
	long result;
	InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void Worksheets::Delete()
{
	InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void Worksheets::FillAcrossSheets(LPDISPATCH Range, long Type)
{
	static BYTE parms[] =
		VTS_DISPATCH VTS_I4;
	InvokeHelper(0x1d5, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Range, Type);
}

LPDISPATCH Worksheets::GetItem(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

void Worksheets::Move(const VARIANT& Before, const VARIANT& After)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x27d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Before, &After);
}

LPUNKNOWN Worksheets::Get_NewEnum()
{
	LPUNKNOWN result;
	InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
	return result;
}

void Worksheets::PrintPreview(const VARIANT& EnableChanges)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x119, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &EnableChanges);
}

void Worksheets::Select(const VARIANT& Replace)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xeb, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Replace);
}

LPDISPATCH Worksheets::GetHPageBreaks()
{
	LPDISPATCH result;
	InvokeHelper(0x58a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH Worksheets::GetVPageBreaks()
{
	LPDISPATCH result;
	InvokeHelper(0x58b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

VARIANT Worksheets::GetVisible()
{
	VARIANT result;
	InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

void Worksheets::SetVisible(const VARIANT& newValue)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 &newValue);
}

LPDISPATCH Worksheets::Get_Default(const VARIANT& Index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&Index);
	return result;
}

void Worksheets::PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate, const VARIANT& PrToFileName)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x6ec, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName);
}
