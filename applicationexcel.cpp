// Machine generated IDispatch wrapper class(es) created with ClassWizard

#include "stdafx.h"
#include "applicationexcel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



/////////////////////////////////////////////////////////////////////////////
// _Applicationexcel properties

/////////////////////////////////////////////////////////////////////////////
// _Applicationexcel operations

LPDISPATCH _Applicationexcel::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetCreator()
{
	long result;
	InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetActiveCell()
{
	LPDISPATCH result;
	InvokeHelper(0x131, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetActiveChart()
{
	LPDISPATCH result;
	InvokeHelper(0xb7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetActivePrinter()
{
	CString result;
	InvokeHelper(0x132, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetActivePrinter(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x132, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

LPDISPATCH _Applicationexcel::GetActiveSheet()
{
	LPDISPATCH result;
	InvokeHelper(0x133, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetActiveWindow()
{
	LPDISPATCH result;
	InvokeHelper(0x2f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetActiveWorkbook()
{
	LPDISPATCH result;
	InvokeHelper(0x134, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetAddIns()
{
	LPDISPATCH result;
	InvokeHelper(0x225, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetAssistant()
{
	LPDISPATCH result;
	InvokeHelper(0x59e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::Calculate()
{
	InvokeHelper(0x117, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH _Applicationexcel::GetCells()
{
	LPDISPATCH result;
	InvokeHelper(0xee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetCharts()
{
	LPDISPATCH result;
	InvokeHelper(0x79, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetColumns()
{
	LPDISPATCH result;
	InvokeHelper(0xf1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetCommandBars()
{
	LPDISPATCH result;
	InvokeHelper(0x59f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetDDEAppReturnCode()
{
	long result;
	InvokeHelper(0x14c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::DDEExecute(long Channel, LPCTSTR String)
{
	static BYTE parms[] =
		VTS_I4 VTS_BSTR;
	InvokeHelper(0x14d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Channel, String);
}

long _Applicationexcel::DDEInitiate(LPCTSTR App, LPCTSTR Topic)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x14e, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		App, Topic);
	return result;
}

void _Applicationexcel::DDEPoke(long Channel, const VARIANT& Item, const VARIANT& Data)
{
	static BYTE parms[] =
		VTS_I4 VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x14f, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Channel, &Item, &Data);
}

VARIANT _Applicationexcel::DDERequest(long Channel, LPCTSTR Item)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_I4 VTS_BSTR;
	InvokeHelper(0x150, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		Channel, Item);
	return result;
}

void _Applicationexcel::DDETerminate(long Channel)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x151, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Channel);
}

VARIANT _Applicationexcel::Evaluate(const VARIANT& Name)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x1, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&Name);
	return result;
}

VARIANT _Applicationexcel::_Evaluate(const VARIANT& Name)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xfffffffb, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&Name);
	return result;
}

VARIANT _Applicationexcel::ExecuteExcel4Macro(LPCTSTR String)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x15e, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		String);
	return result;
}

LPDISPATCH _Applicationexcel::Intersect(LPDISPATCH Arg1, LPDISPATCH Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
		const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
		const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_DISPATCH VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT 
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x2fe, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
	return result;
}

LPDISPATCH _Applicationexcel::GetNames()
{
	LPDISPATCH result;
	InvokeHelper(0x1ba, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetRange(const VARIANT& Cell1, const VARIANT& Cell2)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&Cell1, &Cell2);
	return result;
}

LPDISPATCH _Applicationexcel::GetRows()
{
	LPDISPATCH result;
	InvokeHelper(0x102, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

VARIANT _Applicationexcel::Run(const VARIANT& Macro, const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, 
		const VARIANT& Arg10, const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, 
		const VARIANT& Arg20, const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, 
		const VARIANT& Arg30)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT 
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x103, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&Macro, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
	return result;
}

VARIANT _Applicationexcel::_Run2(const VARIANT& Macro, const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, 
		const VARIANT& Arg10, const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, 
		const VARIANT& Arg20, const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, 
		const VARIANT& Arg30)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT 
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x326, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&Macro, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
	return result;
}

LPDISPATCH _Applicationexcel::GetSelection()
{
	LPDISPATCH result;
	InvokeHelper(0x93, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SendKeys(const VARIANT& Keys, const VARIANT& Wait)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x17f, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Keys, &Wait);
}

LPDISPATCH _Applicationexcel::GetSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x1e5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetThisWorkbook()
{
	LPDISPATCH result;
	InvokeHelper(0x30a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::Union(LPDISPATCH Arg1, LPDISPATCH Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, const VARIANT& Arg11, 
		const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, const VARIANT& Arg21, 
		const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_DISPATCH VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT 
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x30b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
	return result;
}

LPDISPATCH _Applicationexcel::GetWindows()
{
	LPDISPATCH result;
	InvokeHelper(0x1ae, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetWorkbooks()
{
	LPDISPATCH result;
	InvokeHelper(0x23c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetWorksheetFunction()
{
	LPDISPATCH result;
	InvokeHelper(0x5a0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetWorksheets()
{
	LPDISPATCH result;
	InvokeHelper(0x1ee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetExcel4IntlMacroSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x245, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetExcel4MacroSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x243, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::ActivateMicrosoftApp(long Index)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x447, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Index);
}

void _Applicationexcel::AddChartAutoFormat(const VARIANT& Chart, LPCTSTR Name, const VARIANT& Description)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_BSTR VTS_VARIANT;
	InvokeHelper(0xd8, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Chart, Name, &Description);
}

void _Applicationexcel::AddCustomList(const VARIANT& ListArray, const VARIANT& ByRow)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x30c, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &ListArray, &ByRow);
}

BOOL _Applicationexcel::GetAlertBeforeOverwriting()
{
	BOOL result;
	InvokeHelper(0x3a2, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetAlertBeforeOverwriting(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3a2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

CString _Applicationexcel::GetAltStartupPath()
{
	CString result;
	InvokeHelper(0x139, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetAltStartupPath(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x139, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

BOOL _Applicationexcel::GetAskToUpdateLinks()
{
	BOOL result;
	InvokeHelper(0x3e0, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetAskToUpdateLinks(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3e0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetEnableAnimations()
{
	BOOL result;
	InvokeHelper(0x49c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetEnableAnimations(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x49c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Applicationexcel::GetAutoCorrect()
{
	LPDISPATCH result;
	InvokeHelper(0x479, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetBuild()
{
	long result;
	InvokeHelper(0x13a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetCalculateBeforeSave()
{
	BOOL result;
	InvokeHelper(0x13b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCalculateBeforeSave(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x13b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetCalculation()
{
	long result;
	InvokeHelper(0x13c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCalculation(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x13c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

VARIANT _Applicationexcel::GetCaller(const VARIANT& Index)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x13d, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		&Index);
	return result;
}

BOOL _Applicationexcel::GetCanPlaySounds()
{
	BOOL result;
	InvokeHelper(0x13e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetCanRecordSounds()
{
	BOOL result;
	InvokeHelper(0x13f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetCaption()
{
	CString result;
	InvokeHelper(0x8b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCaption(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x8b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

BOOL _Applicationexcel::GetCellDragAndDrop()
{
	BOOL result;
	InvokeHelper(0x140, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCellDragAndDrop(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x140, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

double _Applicationexcel::CentimetersToPoints(double Centimeters)
{
	double result;
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x43e, DISPATCH_METHOD, VT_R8, (void*)&result, parms,
		Centimeters);
	return result;
}

BOOL _Applicationexcel::CheckSpelling(LPCTSTR Word, const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase)
{
	BOOL result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x1f9, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms,
		Word, &CustomDictionary, &IgnoreUppercase);
	return result;
}

VARIANT _Applicationexcel::GetClipboardFormats(const VARIANT& Index)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x141, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		&Index);
	return result;
}

BOOL _Applicationexcel::GetDisplayClipboardWindow()
{
	BOOL result;
	InvokeHelper(0x142, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayClipboardWindow(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x142, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetCommandUnderlines()
{
	long result;
	InvokeHelper(0x143, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCommandUnderlines(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x143, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Applicationexcel::GetConstrainNumeric()
{
	BOOL result;
	InvokeHelper(0x144, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetConstrainNumeric(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x144, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

VARIANT _Applicationexcel::ConvertFormula(const VARIANT& Formula, long FromReferenceStyle, const VARIANT& ToReferenceStyle, const VARIANT& ToAbsolute, const VARIANT& RelativeTo)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x145, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&Formula, FromReferenceStyle, &ToReferenceStyle, &ToAbsolute, &RelativeTo);
	return result;
}

BOOL _Applicationexcel::GetCopyObjectsWithCells()
{
	BOOL result;
	InvokeHelper(0x3df, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCopyObjectsWithCells(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3df, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetCursor()
{
	long result;
	InvokeHelper(0x489, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCursor(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x489, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long _Applicationexcel::GetCustomListCount()
{
	long result;
	InvokeHelper(0x313, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetCutCopyMode()
{
	long result;
	InvokeHelper(0x14a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCutCopyMode(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x14a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long _Applicationexcel::GetDataEntryMode()
{
	long result;
	InvokeHelper(0x14b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDataEntryMode(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x14b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

CString _Applicationexcel::Get_Default()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetDefaultFilePath()
{
	CString result;
	InvokeHelper(0x40e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDefaultFilePath(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x40e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

void _Applicationexcel::DeleteChartAutoFormat(LPCTSTR Name)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0xd9, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Name);
}

void _Applicationexcel::DeleteCustomList(long ListNum)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x30f, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 ListNum);
}

LPDISPATCH _Applicationexcel::GetDialogs()
{
	LPDISPATCH result;
	InvokeHelper(0x2f9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetDisplayAlerts()
{
	BOOL result;
	InvokeHelper(0x157, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayAlerts(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x157, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetDisplayFormulaBar()
{
	BOOL result;
	InvokeHelper(0x158, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayFormulaBar(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x158, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetDisplayFullScreen()
{
	BOOL result;
	InvokeHelper(0x425, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayFullScreen(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x425, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetDisplayNoteIndicator()
{
	BOOL result;
	InvokeHelper(0x159, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayNoteIndicator(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x159, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetDisplayCommentIndicator()
{
	long result;
	InvokeHelper(0x4ac, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayCommentIndicator(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x4ac, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Applicationexcel::GetDisplayExcel4Menus()
{
	BOOL result;
	InvokeHelper(0x39f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayExcel4Menus(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x39f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetDisplayRecentFiles()
{
	BOOL result;
	InvokeHelper(0x39e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayRecentFiles(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x39e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetDisplayScrollBars()
{
	BOOL result;
	InvokeHelper(0x15a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayScrollBars(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x15a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetDisplayStatusBar()
{
	BOOL result;
	InvokeHelper(0x15b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayStatusBar(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x15b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Applicationexcel::DoubleClick()
{
	InvokeHelper(0x15d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

BOOL _Applicationexcel::GetEditDirectlyInCell()
{
	BOOL result;
	InvokeHelper(0x3a1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetEditDirectlyInCell(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3a1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetEnableAutoComplete()
{
	BOOL result;
	InvokeHelper(0x49b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetEnableAutoComplete(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x49b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetEnableCancelKey()
{
	long result;
	InvokeHelper(0x448, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetEnableCancelKey(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x448, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Applicationexcel::GetEnableSound()
{
	BOOL result;
	InvokeHelper(0x4ad, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetEnableSound(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4ad, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

VARIANT _Applicationexcel::GetFileConverters(const VARIANT& Index1, const VARIANT& Index2)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x3a3, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		&Index1, &Index2);
	return result;
}

LPDISPATCH _Applicationexcel::GetFileSearch()
{
	LPDISPATCH result;
	InvokeHelper(0x4b0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetFileFind()
{
	LPDISPATCH result;
	InvokeHelper(0x4b1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetFixedDecimal()
{
	BOOL result;
	InvokeHelper(0x15f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetFixedDecimal(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x15f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetFixedDecimalPlaces()
{
	long result;
	InvokeHelper(0x160, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetFixedDecimalPlaces(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x160, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

VARIANT _Applicationexcel::GetCustomListContents(long ListNum)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x312, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		ListNum);
	return result;
}

long _Applicationexcel::GetCustomListNum(const VARIANT& ListArray)
{
	long result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x311, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		&ListArray);
	return result;
}

VARIANT _Applicationexcel::GetOpenFilename(const VARIANT& FileFilter, const VARIANT& FilterIndex, const VARIANT& Title, const VARIANT& ButtonText, const VARIANT& MultiSelect)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x433, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&FileFilter, &FilterIndex, &Title, &ButtonText, &MultiSelect);
	return result;
}

VARIANT _Applicationexcel::GetSaveAsFilename(const VARIANT& InitialFilename, const VARIANT& FileFilter, const VARIANT& FilterIndex, const VARIANT& Title, const VARIANT& ButtonText)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x434, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		&InitialFilename, &FileFilter, &FilterIndex, &Title, &ButtonText);
	return result;
}

void _Applicationexcel::Goto(const VARIANT& Reference, const VARIANT& Scroll)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x1db, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Reference, &Scroll);
}

double _Applicationexcel::GetHeight()
{
	double result;
	InvokeHelper(0x7b, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetHeight(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x7b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

void _Applicationexcel::Help(const VARIANT& HelpFile, const VARIANT& HelpContextID)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x162, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &HelpFile, &HelpContextID);
}

BOOL _Applicationexcel::GetIgnoreRemoteRequests()
{
	BOOL result;
	InvokeHelper(0x164, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetIgnoreRemoteRequests(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x164, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

double _Applicationexcel::InchesToPoints(double Inches)
{
	double result;
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x43f, DISPATCH_METHOD, VT_R8, (void*)&result, parms,
		Inches);
	return result;
}

VARIANT _Applicationexcel::InputBox(LPCTSTR Prompt, const VARIANT& Title, const VARIANT& Default, const VARIANT& Left, const VARIANT& Top, const VARIANT& HelpFile, const VARIANT& HelpContextID, const VARIANT& Type)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x165, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
		Prompt, &Title, &Default, &Left, &Top, &HelpFile, &HelpContextID, &Type);
	return result;
}

BOOL _Applicationexcel::GetInteractive()
{
	BOOL result;
	InvokeHelper(0x169, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetInteractive(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x169, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

VARIANT _Applicationexcel::GetInternational(const VARIANT& Index)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x16a, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		&Index);
	return result;
}

BOOL _Applicationexcel::GetIteration()
{
	BOOL result;
	InvokeHelper(0x16b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetIteration(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x16b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

double _Applicationexcel::GetLeft()
{
	double result;
	InvokeHelper(0x7f, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetLeft(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x7f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

CString _Applicationexcel::GetLibraryPath()
{
	CString result;
	InvokeHelper(0x16e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::MacroOptions(const VARIANT& Macro, const VARIANT& Description, const VARIANT& HasMenu, const VARIANT& MenuText, const VARIANT& HasShortcutKey, const VARIANT& ShortcutKey, const VARIANT& Category, const VARIANT& StatusBar, 
		const VARIANT& HelpContextID, const VARIANT& HelpFile)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x46f, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Macro, &Description, &HasMenu, &MenuText, &HasShortcutKey, &ShortcutKey, &Category, &StatusBar, &HelpContextID, &HelpFile);
}

void _Applicationexcel::MailLogoff()
{
	InvokeHelper(0x3b1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Applicationexcel::MailLogon(const VARIANT& Name, const VARIANT& Password, const VARIANT& DownloadNewMail)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x3af, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Name, &Password, &DownloadNewMail);
}

VARIANT _Applicationexcel::GetMailSession()
{
	VARIANT result;
	InvokeHelper(0x3ae, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetMailSystem()
{
	long result;
	InvokeHelper(0x3cb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetMathCoprocessorAvailable()
{
	BOOL result;
	InvokeHelper(0x16f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

double _Applicationexcel::GetMaxChange()
{
	double result;
	InvokeHelper(0x170, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetMaxChange(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x170, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

long _Applicationexcel::GetMaxIterations()
{
	long result;
	InvokeHelper(0x171, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetMaxIterations(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x171, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Applicationexcel::GetMouseAvailable()
{
	BOOL result;
	InvokeHelper(0x175, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetMoveAfterReturn()
{
	BOOL result;
	InvokeHelper(0x176, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetMoveAfterReturn(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x176, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetMoveAfterReturnDirection()
{
	long result;
	InvokeHelper(0x478, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetMoveAfterReturnDirection(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x478, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH _Applicationexcel::GetRecentFiles()
{
	LPDISPATCH result;
	InvokeHelper(0x4b2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetName()
{
	CString result;
	InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::NextLetter()
{
	LPDISPATCH result;
	InvokeHelper(0x3cc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetNetworkTemplatesPath()
{
	CString result;
	InvokeHelper(0x184, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetODBCErrors()
{
	LPDISPATCH result;
	InvokeHelper(0x4b3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetODBCTimeout()
{
	long result;
	InvokeHelper(0x4b4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetODBCTimeout(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x4b4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

void _Applicationexcel::OnKey(LPCTSTR Key, const VARIANT& Procedure)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT;
	InvokeHelper(0x272, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Key, &Procedure);
}

void _Applicationexcel::OnRepeat(LPCTSTR Text, LPCTSTR Procedure)
{
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x301, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Text, Procedure);
}

void _Applicationexcel::OnTime(const VARIANT& EarliestTime, LPCTSTR Procedure, const VARIANT& LatestTime, const VARIANT& Schedule)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_BSTR VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x270, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &EarliestTime, Procedure, &LatestTime, &Schedule);
}

void _Applicationexcel::OnUndo(LPCTSTR Text, LPCTSTR Procedure)
{
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x302, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Text, Procedure);
}

CString _Applicationexcel::GetOnWindow()
{
	CString result;
	InvokeHelper(0x26f, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetOnWindow(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x26f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString _Applicationexcel::GetOperatingSystem()
{
	CString result;
	InvokeHelper(0x177, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetOrganizationName()
{
	CString result;
	InvokeHelper(0x178, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetPath()
{
	CString result;
	InvokeHelper(0x123, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetPathSeparator()
{
	CString result;
	InvokeHelper(0x179, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

VARIANT _Applicationexcel::GetPreviousSelections(const VARIANT& Index)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x17a, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		&Index);
	return result;
}

BOOL _Applicationexcel::GetPivotTableSelection()
{
	BOOL result;
	InvokeHelper(0x4b5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetPivotTableSelection(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4b5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetPromptForSummaryInfo()
{
	BOOL result;
	InvokeHelper(0x426, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetPromptForSummaryInfo(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x426, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Applicationexcel::Quit()
{
	InvokeHelper(0x12e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void _Applicationexcel::RecordMacro(const VARIANT& BasicCode, const VARIANT& XlmCode)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x305, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &BasicCode, &XlmCode);
}

BOOL _Applicationexcel::GetRecordRelative()
{
	BOOL result;
	InvokeHelper(0x17b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetReferenceStyle()
{
	long result;
	InvokeHelper(0x17c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetReferenceStyle(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x17c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

VARIANT _Applicationexcel::GetRegisteredFunctions(const VARIANT& Index1, const VARIANT& Index2)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0x307, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		&Index1, &Index2);
	return result;
}

BOOL _Applicationexcel::RegisterXLL(LPCTSTR Filename)
{
	BOOL result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x1e, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms,
		Filename);
	return result;
}

void _Applicationexcel::Repeat()
{
	InvokeHelper(0x12d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

BOOL _Applicationexcel::GetRollZoom()
{
	BOOL result;
	InvokeHelper(0x4b6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetRollZoom(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4b6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Applicationexcel::SaveWorkspace(const VARIANT& Filename)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0xd4, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Filename);
}

BOOL _Applicationexcel::GetScreenUpdating()
{
	BOOL result;
	InvokeHelper(0x17e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetScreenUpdating(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x17e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Applicationexcel::SetDefaultChart(const VARIANT& FormatName, const VARIANT& Gallery)
{
	static BYTE parms[] =
		VTS_VARIANT VTS_VARIANT;
	InvokeHelper(0xdb, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &FormatName, &Gallery);
}

long _Applicationexcel::GetSheetsInNewWorkbook()
{
	long result;
	InvokeHelper(0x3e1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetSheetsInNewWorkbook(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x3e1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Applicationexcel::GetShowChartTipNames()
{
	BOOL result;
	InvokeHelper(0x4b7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetShowChartTipNames(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4b7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetShowChartTipValues()
{
	BOOL result;
	InvokeHelper(0x4b8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetShowChartTipValues(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4b8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

CString _Applicationexcel::GetStandardFont()
{
	CString result;
	InvokeHelper(0x39c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetStandardFont(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x39c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

double _Applicationexcel::GetStandardFontSize()
{
	double result;
	InvokeHelper(0x39d, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetStandardFontSize(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x39d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

CString _Applicationexcel::GetStartupPath()
{
	CString result;
	InvokeHelper(0x181, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

VARIANT _Applicationexcel::GetStatusBar()
{
	VARIANT result;
	InvokeHelper(0x182, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetStatusBar(const VARIANT& newValue)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x182, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 &newValue);
}

CString _Applicationexcel::GetTemplatesPath()
{
	CString result;
	InvokeHelper(0x17d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetShowToolTips()
{
	BOOL result;
	InvokeHelper(0x183, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetShowToolTips(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x183, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

double _Applicationexcel::GetTop()
{
	double result;
	InvokeHelper(0x7e, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetTop(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x7e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

long _Applicationexcel::GetDefaultSaveFormat()
{
	long result;
	InvokeHelper(0x4b9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDefaultSaveFormat(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x4b9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

CString _Applicationexcel::GetTransitionMenuKey()
{
	CString result;
	InvokeHelper(0x136, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetTransitionMenuKey(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x136, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

long _Applicationexcel::GetTransitionMenuKeyAction()
{
	long result;
	InvokeHelper(0x137, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetTransitionMenuKeyAction(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x137, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Applicationexcel::GetTransitionNavigKeys()
{
	BOOL result;
	InvokeHelper(0x138, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetTransitionNavigKeys(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x138, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Applicationexcel::Undo()
{
	InvokeHelper(0x12f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

double _Applicationexcel::GetUsableHeight()
{
	double result;
	InvokeHelper(0x185, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

double _Applicationexcel::GetUsableWidth()
{
	double result;
	InvokeHelper(0x186, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetUserControl()
{
	BOOL result;
	InvokeHelper(0x4ba, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetUserControl(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4ba, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

CString _Applicationexcel::GetUserName_()
{
	CString result;
	InvokeHelper(0x187, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetUserName(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x187, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString _Applicationexcel::GetValue()
{
	CString result;
	InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetVbe()
{
	LPDISPATCH result;
	InvokeHelper(0x4bb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetVersion()
{
	CString result;
	InvokeHelper(0x188, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetVisible()
{
	BOOL result;
	InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetVisible(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Applicationexcel::Volatile(const VARIANT& Volatile)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x314, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &Volatile);
}

double _Applicationexcel::GetWidth()
{
	double result;
	InvokeHelper(0x7a, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetWidth(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x7a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

BOOL _Applicationexcel::GetWindowsForPens()
{
	BOOL result;
	InvokeHelper(0x18b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetWindowState()
{
	long result;
	InvokeHelper(0x18c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetWindowState(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x18c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long _Applicationexcel::GetDefaultSheetDirection()
{
	long result;
	InvokeHelper(0xe5, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDefaultSheetDirection(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xe5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long _Applicationexcel::GetCursorMovement()
{
	long result;
	InvokeHelper(0xe8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCursorMovement(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xe8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Applicationexcel::GetControlCharacters()
{
	BOOL result;
	InvokeHelper(0xe9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetControlCharacters(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0xe9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetEnableEvents()
{
	BOOL result;
	InvokeHelper(0x4bc, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetEnableEvents(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4bc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::Wait(const VARIANT& Time)
{
	BOOL result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x6ea, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms,
		&Time);
	return result;
}

BOOL _Applicationexcel::GetExtendList()
{
	BOOL result;
	InvokeHelper(0x701, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetExtendList(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x701, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Applicationexcel::GetOLEDBErrors()
{
	LPDISPATCH result;
	InvokeHelper(0x702, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetPhonetic(const VARIANT& Text)
{
	CString result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x703, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms,
		&Text);
	return result;
}

LPDISPATCH _Applicationexcel::GetCOMAddIns()
{
	LPDISPATCH result;
	InvokeHelper(0x704, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetDefaultWebOptions()
{
	LPDISPATCH result;
	InvokeHelper(0x705, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetProductCode()
{
	CString result;
	InvokeHelper(0x706, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString _Applicationexcel::GetUserLibraryPath()
{
	CString result;
	InvokeHelper(0x707, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetAutoPercentEntry()
{
	BOOL result;
	InvokeHelper(0x708, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetAutoPercentEntry(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x708, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Applicationexcel::GetLanguageSettings()
{
	LPDISPATCH result;
	InvokeHelper(0x709, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetAnswerWizard()
{
	LPDISPATCH result;
	InvokeHelper(0x70c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::CalculateFull()
{
	InvokeHelper(0x70d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

BOOL _Applicationexcel::FindFile()
{
	BOOL result;
	InvokeHelper(0x6eb, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetCalculationVersion()
{
	long result;
	InvokeHelper(0x70e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetShowWindowsInTaskbar()
{
	BOOL result;
	InvokeHelper(0x70f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetShowWindowsInTaskbar(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x70f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetFeatureInstall()
{
	long result;
	InvokeHelper(0x710, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetFeatureInstall(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x710, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL _Applicationexcel::GetReady()
{
	BOOL result;
	InvokeHelper(0x78c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetFindFormat()
{
	LPDISPATCH result;
	InvokeHelper(0x78e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetRefFindFormat(LPDISPATCH newValue)
{
	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper(0x78e, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms,
		 newValue);
}

LPDISPATCH _Applicationexcel::GetReplaceFormat()
{
	LPDISPATCH result;
	InvokeHelper(0x78f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetRefReplaceFormat(LPDISPATCH newValue)
{
	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper(0x78f, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms,
		 newValue);
}

LPDISPATCH _Applicationexcel::GetUsedObjects()
{
	LPDISPATCH result;
	InvokeHelper(0x790, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetCalculationState()
{
	long result;
	InvokeHelper(0x791, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetCalculationInterruptKey()
{
	long result;
	InvokeHelper(0x792, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetCalculationInterruptKey(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x792, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH _Applicationexcel::GetWatches()
{
	LPDISPATCH result;
	InvokeHelper(0x793, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetDisplayFunctionToolTips()
{
	BOOL result;
	InvokeHelper(0x794, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayFunctionToolTips(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x794, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long _Applicationexcel::GetAutomationSecurity()
{
	long result;
	InvokeHelper(0x795, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetAutomationSecurity(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x795, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH _Applicationexcel::GetFileDialog(long fileDialogType)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x796, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		fileDialogType);
	return result;
}

void _Applicationexcel::CalculateFullRebuild()
{
	InvokeHelper(0x799, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

BOOL _Applicationexcel::GetDisplayPasteOptions()
{
	BOOL result;
	InvokeHelper(0x79a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayPasteOptions(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x79a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetDisplayInsertOptions()
{
	BOOL result;
	InvokeHelper(0x79b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayInsertOptions(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x79b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetGenerateGetPivotData()
{
	BOOL result;
	InvokeHelper(0x79c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetGenerateGetPivotData(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x79c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Applicationexcel::GetAutoRecover()
{
	LPDISPATCH result;
	InvokeHelper(0x79d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetHwnd()
{
	long result;
	InvokeHelper(0x79e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long _Applicationexcel::GetHinstance()
{
	long result;
	InvokeHelper(0x79f, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::CheckAbort(const VARIANT& KeepAbort)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x7a0, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &KeepAbort);
}

LPDISPATCH _Applicationexcel::GetErrorCheckingOptions()
{
	LPDISPATCH result;
	InvokeHelper(0x7a2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetAutoFormatAsYouTypeReplaceHyperlinks()
{
	BOOL result;
	InvokeHelper(0x7a3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetAutoFormatAsYouTypeReplaceHyperlinks(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x7a3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Applicationexcel::GetSmartTagRecognizers()
{
	LPDISPATCH result;
	InvokeHelper(0x7a4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetNewWorkbook()
{
	LPDISPATCH result;
	InvokeHelper(0x61d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetSpellingOptions()
{
	LPDISPATCH result;
	InvokeHelper(0x7a5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetSpeech()
{
	LPDISPATCH result;
	InvokeHelper(0x7a6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetMapPaperSize()
{
	BOOL result;
	InvokeHelper(0x7a7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetMapPaperSize(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x7a7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL _Applicationexcel::GetShowStartupDialog()
{
	BOOL result;
	InvokeHelper(0x7a8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetShowStartupDialog(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x7a8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

CString _Applicationexcel::GetDecimalSeparator()
{
	CString result;
	InvokeHelper(0x711, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDecimalSeparator(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x711, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString _Applicationexcel::GetThousandsSeparator()
{
	CString result;
	InvokeHelper(0x712, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetThousandsSeparator(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x712, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

BOOL _Applicationexcel::GetUseSystemSeparators()
{
	BOOL result;
	InvokeHelper(0x7a9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetUseSystemSeparators(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x7a9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH _Applicationexcel::GetThisCell()
{
	LPDISPATCH result;
	InvokeHelper(0x7aa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH _Applicationexcel::GetRtd()
{
	LPDISPATCH result;
	InvokeHelper(0x7ab, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL _Applicationexcel::GetDisplayDocumentActionTaskPane()
{
	BOOL result;
	InvokeHelper(0x8cb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void _Applicationexcel::SetDisplayDocumentActionTaskPane(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x8cb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

void _Applicationexcel::DisplayXMLSourcePane(const VARIANT& XmlMap)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x8cc, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &XmlMap);
}

BOOL _Applicationexcel::GetArbitraryXMLSupportAvailable()
{
	BOOL result;
	InvokeHelper(0x8ce, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}
