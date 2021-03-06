#if !defined(AFX__SMSGATE1_H__79D39E00_50BD_48CF_9793_51B9E368767E__INCLUDED_)
#define AFX__SMSGATE1_H__79D39E00_50BD_48CF_9793_51B9E368767E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Machine generated IDispatch wrapper class(es) created by Microsoft Visual C++

// NOTE: Do not modify the contents of this file.  If this class is regenerated by
//  Microsoft Visual C++, your modifications will be overwritten.

/////////////////////////////////////////////////////////////////////////////
// C_Smsgate wrapper class

class C_Smsgate : public CWnd
{
protected:
	DECLARE_DYNCREATE(C_Smsgate)
public:
	CLSID const& GetClsid()
	{
		static CLSID const clsid
			= { 0x45e0b554, 0x4b72, 0x439f, { 0xa5, 0x52, 0xbc, 0xc0, 0x4d, 0x98, 0x1b, 0x7c } };
		return clsid;
	}
	virtual BOOL Create(LPCTSTR lpszClassName,
		LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect,
		CWnd* pParentWnd, UINT nID,
		CCreateContext* pContext = NULL)
	{ return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID); }

    BOOL Create(LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect, CWnd* pParentWnd, UINT nID,
		CFile* pPersist = NULL, BOOL bStorage = FALSE,
		BSTR bstrLicKey = NULL)
	{ return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID,
		pPersist, bStorage, bstrLicKey); }

// Attributes
public:

// Operations
public:
	short GetCommPort();
	void SetCommPort(short nNewValue);
	VARIANT NewMsg();
	VARIANT Connect(short* WaitTime);
	CString Sendsms(BSTR* Msg, BSTR* Mobile, short* s_Report, BOOL* pV);
	CString SendAsc2(VARIANT* AscMsg, VARIANT* tomobile, VARIANT* smsreport);
	VARIANT AnswerCall();
	BOOL HangUpCall();
	VARIANT CallPhone(BSTR* PhoneNo, short* WaitTime);
	VARIANT GetPhoneNo();
	VARIANT ReadMsg(short* whyre);
	VARIANT DelSms(short* del_which);
	VARIANT RevAuto();
	VARIANT RevAutoClose();
	VARIANT ClosePort();
	VARIANT GetErrmsg();
	void SetErrmsg(const VARIANT& newValue);
	VARIANT M_model();
	VARIANT M_ltd();
	VARIANT M_ver();
	VARIANT M_ServiceNo();
	VARIANT M_imei();
	VARIANT InputAT(BSTR* strAT);
	VARIANT ReadNB(VARIANT* pb_why);
	long GetRevInterval();
	void SetRevInterval(long nNewValue);
	CString GetSmsService();
	void SetSmsService(LPCTSTR lpszNewValue);
	BOOL GetReadAndDel();
	void SetReadAndDel(BOOL bNewValue);
	CString GetSn();
	void SetSn(LPCTSTR lpszNewValue);
	BOOL Link();
	BOOL GetIsbusy();
	void SetIsbusy(BOOL bNewValue);
	VARIANT RdCenNo();
	CString GetSettings();
	void SetSettings(LPCTSTR lpszNewValue);
	CString GetCopyRight();
	void SetCopyRight(LPCTSTR lpszNewValue);
	VARIANT get_report_all(VARIANT* start_is, VARIANT* strCodeis);
	VARIANT NewReport();
	VARIANT Get_ID1();
	VARIANT Get_ID2();
	VARIANT Get_ID3();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX__SMSGATE1_H__79D39E00_50BD_48CF_9793_51B9E368767E__INCLUDED_)
