// duanxin.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "duanxin.h"
#include "duanxinDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDuanxinApp

BEGIN_MESSAGE_MAP(CDuanxinApp, CWinApp)
	//{{AFX_MSG_MAP(CDuanxinApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDuanxinApp construction

CDuanxinApp::CDuanxinApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CDuanxinApp object

CDuanxinApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CDuanxinApp initialization

BOOL CDuanxinApp::InitInstance()
{
	AfxEnableControlContainer();
	//一下为添加皮肤的部分
	skinppLoadSkin(_T("Vista.ssk"));//Vista.ssk为项目下的皮肤文件
	
	
	//皮肤部分结束
	// Standard initialization
	// If you are not using these features and wish to reduce the size
	//  of your final executable, you should remove from the following
	//  the specific initialization routines you do not need.

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	CDuanxinDlg dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with OK
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
