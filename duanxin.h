// duanxin.h : main header file for the DUANXIN application
//

#if !defined(AFX_DUANXIN_H__CB4F3049_0B84_4637_B348_49C1006AFA39__INCLUDED_)
#define AFX_DUANXIN_H__CB4F3049_0B84_4637_B348_49C1006AFA39__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols
#define SKINSPACE _T("/SPATH:") ////  注意：这个必须添加在#include的下面！！！
/////////////////////////////////////////////////////////////////////////////
// CDuanxinApp:
// See duanxin.cpp for the implementation of this class
//

class CDuanxinApp : public CWinApp
{
public:
	CDuanxinApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDuanxinApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CDuanxinApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DUANXIN_H__CB4F3049_0B84_4637_B348_49C1006AFA39__INCLUDED_)
