// duanxinDlg.h : header file
//
//{{AFX_INCLUDES()
#include "_smsgate.h"
//}}AFX_INCLUDES
#include "SkinPPWTL.h"
#if !defined(AFX_DUANXINDLG_H__50BCD19A_22CF_4EE2_B43A_72290C3A73FB__INCLUDED_)
#define AFX_DUANXINDLG_H__50BCD19A_22CF_4EE2_B43A_72290C3A73FB__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CDuanxinDlg dialog

class CDuanxinDlg : public CDialog
{
// Construction
public:
	void ShowBMP();
	void saveMessage();
	void splitMessage();
	CDuanxinDlg(CWnd* pParent = NULL);	// standard constructor
	HICON m_hIconRed;    //串口打开时的红灯图标句柄
	HICON m_hIconOff;    //串口关闭时的指示图标句柄

// Dialog Data
	//{{AFX_DATA(CDuanxinDlg)
	enum { IDD = IDD_DUANXIN_DIALOG };
	CListBox	m_listbox_bmp;
	CStatic	m_com_openoff;
	CComboBox	m_comport;
	CListBox	m_clistbox;
	CTreeCtrl	m_ctrltree;
	C_Smsgate	m_smsgate_1;
	CString	m_cscs_bsic;
	CString	m_cscs_cid;
	CString	m_cscs_lac;
	CString	m_cscs_pd;
	CString	m_cscs_jd;
	CString	m_cscs_wd;
	CString	m_cscs_fxj;
	CString	m_cscs_qj;
	CString	m_cscs_hgj;
	CString	m_cscs_dmhb;
	CString	m_cscs_txhb;
	CString	m_cscs_txgg;
	CString	m_cscs_qbdp;
	CString	m_cscs_hbdp;
	CString	m_cscs_dpbz;
	CString	m_csjsy;
	CString	m_jzmc;
	CString	m_csrq;
	CString	m_txmc;
	CString	m_sqcs_bsic;
	CString	m_sqcs_cid;
	CString	m_sqcs_lac;
	CString	m_sqcs_pd;
	CString	m_sqcs_jd;
	CString	m_sqcs_wd;
	CString	m_sqcs_fxj;
	CString	m_sqcs_qj;
	CString	m_sqcs_hgj;
	CString	m_sqcs_dmhb;
	CString	m_sqcs_txhb;
	CString	m_sqcs_txgg;
	CString	m_sqcs_qbdp;
	CString	m_sqcs_dpbz;
	CString	m_bzxx;
	CString	m_sqcs_hbdp;
	CString	m_save_path;
	CString	m_local_telnumber;
	CString	m_sqcs_xhqd;
	CString	m_sqcs_ratio;
	CString	m_cscs_xhqd;
	CString	m_cscs_ratio;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDuanxinDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	virtual void CalcWindowRect(LPRECT lpClientRect, UINT nAdjustType = adjustBorder);
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CDuanxinDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnNewitem();
	afx_msg void OnOnRecvMsgSmsgate1();
	afx_msg void OnDestroy();
	afx_msg void OnSelchangeList1();
	afx_msg void Onmodifyparam();
	afx_msg void OnSavePath();
	afx_msg void OnAnalysisiReport();
	afx_msg void OnBrowse();
	afx_msg void OnConnectComport();
	afx_msg void OnDisconnectComport();
	afx_msg void OnSelchangeList2();
	afx_msg void OnTestReport();
	afx_msg void OnImportData();
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DUANXINDLG_H__50BCD19A_22CF_4EE2_B43A_72290C3A73FB__INCLUDED_)
