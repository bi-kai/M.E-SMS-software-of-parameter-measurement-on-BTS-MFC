// duanxinDlg.cpp : implementation file
//

#include "stdafx.h"
#include "duanxin.h"
#include "duanxinDlg.h"
#include "_smsgate.h"
#include <atlbase.h>
#include <afxcoll.h>
#include   <comdef.h>
#include "msword.h"
#include "excel.h"
#include "applicationexcel.h"
#include "rangeexcel.h"
#include "shapesexcel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

HTREEITEM m_hInsertItem;
short waittime;

CString s_msg;
CString s_mobile;
short s_report;
int s_pv;
BSTR bstr_s_msg;
BSTR bstr_s_mobile;
VARIANT s_wait_time;
CString s_newmessage;
bool flg_receive_msg;

CString message_data;//ÿ����Ϣ�ĺ�������

#define MAX_SECTION 260//Section��󳤶�
#define MAX_KEY 260//KeyValues��󳤶�
#define MAX_ALLSECTIONS 65535//����Section����󳤶�
#define MAX_ALLKEYS 65535//����KeyValue����󳤶�

BOOL modify_flag;

BITMAPINFOHEADER *m_bmpInfoHeader;
unsigned char *m_pDib;
unsigned char *m_pDibBits;
DWORD dwDibSize;
DWORD nFileLen;
int lWidth;
int lHeight;
int lBitCount;
int NumColor;
CString strFileName;//ͼƬ�ļ�·��
bool browse_flag;//��ȷ��ͼƬ��־λ
bool savepath_flag;//����·����ť�Ƿ����õı�־λ

CString  m_strPath;//(����word��excel��·��),bmpͼƬ·��
CString str_current_Path;//��ǰ·��
CString currentpath_buf_para;//ini���ò����ļ�·��
CString currentpath_buf;//ini�����ļ�·��

bool com_set_flag;//com���Ƿ����Ӻõı�־λ��0��δ���ӣ�1��������
bool RevAuto_once_flag;//�Զ�����״̬��������һ�εı�־λ

CStringArray data_bmp;//ͼƬ�б���е�ѡ��������ɵ��ַ�������
CStringArray data_data;//�����б���е�ѡ��������ɵ��ַ�������
CStringArray data_txt;//�����txt�����ļ�������ɵ��ַ�������
bool finded_bmp_flag;//ƥ������ͼƬ�ı�־λ
/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDuanxinDlg dialog

CDuanxinDlg::CDuanxinDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDuanxinDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDuanxinDlg)
	m_cscs_bsic = _T("");
	m_cscs_cid = _T("");
	m_cscs_lac = _T("");
	m_cscs_pd = _T("");
	m_cscs_jd = _T("");
	m_cscs_wd = _T("");
	m_cscs_fxj = _T("");
	m_cscs_qj = _T("");
	m_cscs_hgj = _T("");
	m_cscs_dmhb = _T("");
	m_cscs_txhb = _T("");
	m_cscs_txgg = _T("");
	m_cscs_qbdp = _T("");
	m_cscs_hbdp = _T("");
	m_cscs_dpbz = _T("");
	m_csjsy = _T("");
	m_jzmc = _T("");
	m_csrq = _T("");
	m_txmc = _T("");
	m_sqcs_bsic = _T("");
	m_sqcs_cid = _T("");
	m_sqcs_lac = _T("");
	m_sqcs_pd = _T("");
	m_sqcs_jd = _T("");
	m_sqcs_wd = _T("");
	m_sqcs_fxj = _T("");
	m_sqcs_qj = _T("");
	m_sqcs_hgj = _T("");
	m_sqcs_dmhb = _T("");
	m_sqcs_txhb = _T("");
	m_sqcs_txgg = _T("");
	m_sqcs_qbdp = _T("");
	m_sqcs_dpbz = _T("");
	m_bzxx = _T("");
	m_sqcs_hbdp = _T("");
	m_save_path = _T("");
	m_local_telnumber = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CDuanxinDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDuanxinDlg)
	DDX_Control(pDX, IDC_LIST2, m_listbox_bmp);
	DDX_Control(pDX, IDC_COM_OPENOFF, m_com_openoff);
	DDX_Control(pDX, IDC_COMBO_COMSELECT, m_comport);
	DDX_Control(pDX, IDC_LIST1, m_clistbox);
	DDX_Control(pDX, IDC_SMSGATE1, m_smsgate_1);
	DDX_Text(pDX, IDC_EDIT12, m_cscs_bsic);
	DDX_Text(pDX, IDC_EDIT11, m_cscs_cid);
	DDX_Text(pDX, IDC_EDIT10, m_cscs_lac);
	DDX_Text(pDX, IDC_EDIT35, m_cscs_pd);
	DDX_Text(pDX, IDC_EDIT15, m_cscs_jd);
	DDX_Text(pDX, IDC_EDIT16, m_cscs_wd);
	DDX_Text(pDX, IDC_EDIT6, m_cscs_fxj);
	DDX_Text(pDX, IDC_EDIT5, m_cscs_qj);
	DDX_Text(pDX, IDC_EDIT21, m_cscs_hgj);
	DDX_Text(pDX, IDC_EDIT23, m_cscs_dmhb);
	DDX_Text(pDX, IDC_EDIT25, m_cscs_txhb);
	DDX_Text(pDX, IDC_EDIT24, m_cscs_txgg);
	DDX_Text(pDX, IDC_EDIT22, m_cscs_qbdp);
	DDX_Text(pDX, IDC_EDIT33, m_cscs_hbdp);
	DDX_Text(pDX, IDC_EDIT32, m_cscs_dpbz);
	DDX_Text(pDX, IDC_EDIT1, m_csjsy);
	DDX_Text(pDX, IDC_EDIT2, m_jzmc);
	DDX_Text(pDX, IDC_EDIT3, m_csrq);
	DDX_Text(pDX, IDC_EDIT4, m_txmc);
	DDX_Text(pDX, IDC_EDIT9, m_sqcs_bsic);
	DDX_Text(pDX, IDC_EDIT14, m_sqcs_cid);
	DDX_Text(pDX, IDC_EDIT7, m_sqcs_lac);
	DDX_Text(pDX, IDC_EDIT34, m_sqcs_pd);
	DDX_Text(pDX, IDC_EDIT13, m_sqcs_jd);
	DDX_Text(pDX, IDC_EDIT8, m_sqcs_wd);
	DDX_Text(pDX, IDC_EDIT17, m_sqcs_fxj);
	DDX_Text(pDX, IDC_EDIT18, m_sqcs_qj);
	DDX_Text(pDX, IDC_EDIT20, m_sqcs_hgj);
	DDX_Text(pDX, IDC_EDIT26, m_sqcs_dmhb);
	DDX_Text(pDX, IDC_EDIT28, m_sqcs_txhb);
	DDX_Text(pDX, IDC_EDIT27, m_sqcs_txgg);
	DDX_Text(pDX, IDC_EDIT29, m_sqcs_qbdp);
	DDX_Text(pDX, IDC_EDIT31, m_sqcs_dpbz);
	DDX_Text(pDX, IDC_EDIT19, m_bzxx);
	DDX_Text(pDX, IDC_EDIT30, m_sqcs_hbdp);
	DDX_Text(pDX, IDC_EDIT36, m_save_path);
	DDX_Text(pDX, IDC_EDIT37, m_local_telnumber);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CDuanxinDlg, CDialog)
	//{{AFX_MSG_MAP(CDuanxinDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_NEWITEM, OnNewitem)
	ON_WM_DESTROY()
	ON_LBN_SELCHANGE(IDC_LIST1, OnSelchangeList1)
	ON_BN_CLICKED(IDC_BUTTON5, Onmodifyparam)
	ON_BN_CLICKED(IDC_BUTTON4, OnSavePath)
	ON_BN_CLICKED(IDC_BUTTON3, OnAnalysisiReport)
	ON_BN_CLICKED(IDC_BUTTON1, OnBrowse)
	ON_BN_CLICKED(IDC_BUTTON8, OnConnectComport)
	ON_BN_CLICKED(IDC_BUTTON9, OnDisconnectComport)
	ON_LBN_SELCHANGE(IDC_LIST2, OnSelchangeList2)
	ON_BN_CLICKED(IDC_BUTTON2, OnTestReport)
	ON_BN_CLICKED(IDC_BUTTON6, OnImportData)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDuanxinDlg message handlers

BOOL CDuanxinDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	/*************���浱ǰ·��**********************************/
	char current_path[MAX_PATH];
	GetCurrentDirectory(MAX_PATH,current_path);
	str_current_Path=current_path;
//	CString current_puth_exe=str_current_Path+"";
	
	/**************����è*******************************************/
//	InitDriverTree();
// 	m_smsgate_1.SetCommPort(3);
// 	m_smsgate_1.SetSmsService("+8613800290500");
// 	m_smsgate_1.SetSettings("9600,n,8,1");
// 	m_smsgate_1.RevAuto();
// 	m_smsgate_1.SetReadAndDel(TRUE);
// 	waittime=10;
// 	 s_wait_time=m_smsgate_1.Connect(&waittime);

//	m_smsgate_1.SetCommPort(1);
	m_smsgate_1.SetSmsService("+8613800290500");
	m_smsgate_1.SetSettings("9600,n,8,1");
//	m_smsgate_1.RevAuto();
	m_smsgate_1.SetReadAndDel(TRUE);




	 flg_receive_msg=0;//������̣���־λ
		//AfxMessageBox("right!");
	 message_data="";
	 com_set_flag=0;//comδ����
	 m_hIconRed  = AfxGetApp()->LoadIcon(IDI_ICON3);
	 m_hIconOff	= AfxGetApp()->LoadIcon(IDI_ICON2);
	 GetDlgItem(IDC_BUTTON9)->EnableWindow(FALSE);//û���Ӷ˿�ʱ��������Ͽ�
	 RevAuto_once_flag=0;
	 /************�ؼ���ʼ��***************/
	 CString execute_exe;
	 currentpath_buf_para=str_current_Path+"\\Config_para.ini";
	 currentpath_buf=str_current_Path+"\\Config.ini";
	 ::GetPrivateProfileString("config_param","BSIC","unknown",m_sqcs_bsic.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
	 m_sqcs_bsic.ReleaseBuffer();

		 ::GetPrivateProfileString("config_param","CID","unknown",m_sqcs_cid.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_cid.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","LAC","unknown",m_sqcs_lac.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_lac.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","pinduan","unknown",m_sqcs_pd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_pd.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","jingdu","unknown",m_sqcs_jd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_jd.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","weidu","unknown",m_sqcs_wd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_wd.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","fangxiangjiao","unknown",m_sqcs_fxj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_fxj.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","qingjiao","unknown",m_sqcs_qj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_qj.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","henggunjiao","unknown",m_sqcs_hgj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_hgj.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","dimianhaiba","unknown",m_sqcs_dmhb.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_dmhb.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","tianxianhaiba","unknown",m_sqcs_txhb.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_txhb.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","tianxianguagao","unknown",m_sqcs_txgg.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_txgg.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","qianbandianping","unknown",m_sqcs_qbdp.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_qbdp.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","houbandianping","unknown",m_sqcs_hbdp.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_hbdp.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","dianpingbizhi","unknown",m_sqcs_dpbz.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_sqcs_dpbz.ReleaseBuffer();
		 ::GetPrivateProfileString("config_param","local_telnumber","unknown",m_local_telnumber.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
		 m_local_telnumber.ReleaseBuffer();
//		 ::GetPrivateProfileString("config_param","execute","unknown",execute_exe.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf_para);
//  	 execute_exe.ReleaseBuffer();
		 UpdateData(FALSE);
		 /***********��ȡע���������Ƿ��Ѿ���װ************/
		 HKEY hKEY; //�����йص� hKEY, �ڲ�ѯ����ʱҪ�رա�
		 LPCTSTR data_Set="Software\\Microsoft\\Windows\\CurrentVersion\\";
		 CString str_owner;
		 //����·�� data_Set ��ص� hKEY����һ������Ϊ�������ƣ��ڶ���������
		 //��ʾҪ���ʵļ���λ�ã���������������Ϊ0��KEY_READ��ʾ�Բ�ѯ�ķ�ʽ��
		 //����ע���hKEY�򱣴�˺������򿪵ļ��ľ����
		 long ret0=(::RegOpenKeyEx(HKEY_LOCAL_MACHINE,data_Set, 0, KEY_READ, &hKEY));
		 if(ret0!=ERROR_SUCCESS) //����޷���hKEY������ֹ�����ִ��
		 {
			 MessageBox("����: �޷����йص�hKEY!");
			 return 1;
		 }
		 //��ѯ�йص����� (�û����� owner_Get)��
		 DWORD type_1=REG_SZ ; 
		 DWORD cbData_1=80;  
		 //hKEYΪ�ղ�RegOpenKeyEx()�������򿪵ļ��ľ����"RegisteredOwner"��
		 //��ʾҪ�� ѯ�ļ�ֵ����type_1��ʾ��ѯ���ݵ����ͣ�owner_Get��������
		 //��ѯ�����ݣ�cbData_1��ʾԤ���õ����ݳ��ȡ�
		 long ret1=::RegQueryValueEx(hKEY, "softwaredone", NULL,&type_1,(LPBYTE)(LPCSTR)str_owner, &cbData_1);
		 if(ret1!=ERROR_SUCCESS)
		 {
			 MessageBox("����: �޷���ѯ�й�ע�����Ϣ!");
			 return 1;
		 }
		 // �������ǰҪ�ر��Ѿ��򿪵� hKEY��
		::RegCloseKey(hKEY); 
		 /************ע����ȡ���***************************/
		 if (str_owner!="111")
		 {
			 if ((ShellExecute(NULL,"open","HL340.EXE",NULL,NULL,SW_SHOWNORMAL)>(HANDLE)32)&&(ShellExecute(NULL,"open","zckj.bat",NULL,NULL,SW_SHOWNORMAL)>(HANDLE)32))
			 {
				 str_owner="111";
			 } 
			 else
			 {
				 str_owner="0";
			 }
//			 AfxMessageBox(str_owner,MB_OK,0);
		 } 
//		 ::WritePrivateProfileString("config_param","execute",execute_exe,currentpath_buf_para);
		 /*************д��ע�����������****************/
		 //�����йص� hKEY, �ڳ�������Ҫ�رա�
//		 HKEY hKEY;  
//		 LPCTSTR data_Set="Software\\Microsoft\\Windows\\CurrentVersion\\";
		 //����·�� data_Set ��ص�hKEY��KEY_WRITE��ʾ��д�ķ�ʽ�򿪡�
		 long ret2=(::RegOpenKeyEx(HKEY_LOCAL_MACHINE,data_Set, 0, KEY_WRITE, &hKEY));
		 if(ret2!=ERROR_SUCCESS)
		 {
			 MessageBox("����: �޷����йص�hKEY!");
			 return 1;
		 }
		 //�޸��й�����(�û����� owner_Set)��Ҫ�Ƚ�CString��ת��ΪLPBYTE��
//		 DWORD type_1=REG_SZ;
//		 DWORD cbData_1=str_owner.GetLength()+1;  
		 //��RegQureyValueEx()���ƣ�hKEY��ʾ�Ѵ򿪵ļ��ľ����"RegisteredOwner"
		 //��ʾҪ���ʵļ�ֵ����owner_Set��ʾ�µļ�ֵ��type_1��cbData_1��ʾ��ֵ��
		 //���������ͺ����ݳ���
		 long ret3=::RegSetValueEx(hKEY, "softwaredone", NULL,type_1,(LPBYTE)(LPCSTR)str_owner, cbData_1);
		 if(ret3!=ERROR_SUCCESS)
		 {
			 MessageBox("����: �޷��޸��й�ע�����Ϣ!");
			 return 1;
		 }
	::RegCloseKey(hKEY);
		 /*************ע���д�����******************/
		 /*************������ini�����ļ�*********************/
		 CString test;
		 char dd[125];
		 DWORD pp;
		 
		 
		 for (int ju=0;ju<125;ju++)
		 {
			 dd[ju]='\0';
		 }
		 
		 pp=::GetPrivateProfileSectionNames(dd,125,".\\Config.ini");
		 int ii=125,idol=0;
		 int pos=0;
		 CStringArray strArr;
		 int ps=0,pos1=0;
		 char buf[125]={0};
		 CString buf2;
		 while (1)//�õ�������ǰ�¼�����Ϣ��������
		 {
			 pos1=0;
			 if (dd[pos]=='\0')
			 {
				 if (dd[pos+1]=='\0')
				 {
				 break;
			 }
		 }
		 
		 while (dd[pos]!='\0')
		 {
			 buf[pos1]=dd[pos];
			 pos++;pos1++;
		 }
		 buf[pos1]='\0';
		 buf2.Format("%s",buf);
		 strArr.Add(buf2);
	//	 AfxMessageBox(strArr.GetAt(ps),MB_OK,0);
		 m_clistbox.InsertString(0,strArr.GetAt(ps));//��ȡ�����ļ��еļ�¼��������ǰ����Ϣ
		 data_data.Add(strArr.GetAt(ps));
		 ps++;pos++;
	}//end of while(1)

// 		  	for (int p=0;p<data_data.GetSize();p++)
// 		  	{
// 		  		AfxMessageBox(data_data.GetAt(p),MB_OK,0);
// 	 		}

	 CString str1_list;

		 m_clistbox.GetText(0,str1_list);
		 //AfxMessageBox(str1,MB_OK,0);
		 ::GetPrivateProfileString(str1_list,"BSIC","unknown",m_cscs_bsic.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_bsic.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"CID","unknown",m_cscs_cid.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_cid.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"LAC","unknown",m_cscs_lac.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_lac.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"pinduan","unknown",m_cscs_pd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_pd.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"jingdu","unknown",m_cscs_jd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_jd.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"weidu","unknown",m_cscs_wd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_wd.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"fangxiangjiao","unknown",m_cscs_fxj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_fxj.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"qingjiao","unknown",m_cscs_qj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_qj.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"henggunjiao","unknown",m_cscs_hgj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_hgj.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"dimianhaiba","unknown",m_cscs_dmhb.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_dmhb.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"tianxianhaiba","unknown",m_cscs_txhb.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_txhb.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"tianxianguagao","unknown",m_cscs_txgg.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_txgg.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"qianbandianping","unknown",m_cscs_qbdp.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_qbdp.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"houbandianping","unknown",m_cscs_hbdp.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_hbdp.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"dianpingbizhi","unknown",m_cscs_dpbz.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_dpbz.ReleaseBuffer();
		 ::GetPrivateProfileString(str1_list,"date","unknown",m_csrq.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		 m_cscs_dpbz.ReleaseBuffer();
		 UpdateData(FALSE);
	 /*******������ʼ������*****************/
	 m_clistbox.SetCurSel(0);
	modify_flag=0;
	browse_flag=0;
	savepath_flag=0;//����·����ť��־λ
	finded_bmp_flag=0;

	AfxOleInit(); 
	AfxEnableControlContainer();


	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CDuanxinDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CDuanxinDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CDuanxinDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CDuanxinDlg::OnNewitem() 
{
	// TODO: Add your control notification handler code here
	s_msg="gaga,hi!";
	s_mobile="+8615094036912";
	s_report=1;
	s_pv=0;
	bstr_s_msg = s_msg.AllocSysString();//CString ����� AllocSysString ������ CString ת���� BSTR
	bstr_s_mobile = s_mobile.AllocSysString();
	m_smsgate_1.Sendsms(&bstr_s_msg,&bstr_s_mobile,&s_report,&s_pv);
}

//DEL void CDuanxinDlg::OnBeginlabeleditMytree(NMHDR* pNMHDR, LRESULT* pResult) 
//DEL {
//DEL 	TV_DISPINFO* pTVDispInfo = (TV_DISPINFO*)pNMHDR;
//DEL 	// TODO: Add your control notification handler code here
//DEL 	m_treectrl.GetEditControl()->LimitText(16);
//DEL 	*pResult = 0;
//DEL }

//DEL void CDuanxinDlg::OnEndlabeleditMytree(NMHDR* pNMHDR, LRESULT* pResult) 
//DEL {
//DEL 	TV_DISPINFO* pTVDispInfo = (TV_DISPINFO*)pNMHDR;
//DEL 	// TODO: Add your control notification handler code here
//DEL 	CString strName;
//DEL 	m_treectrl.GetEditControl()->GetWindowText(strName);
//DEL 	if (strName.IsEmpty())
//DEL 	{
//DEL 		AfxMessageBox(_T("���������Ʋ���Ϊ�գ�����������"));
//DEL 		CEdit* pEdit=m_treectrl.EditLabel(m_hInsertItem);
//DEL 		ASSERT(pEdit!=NULL);
//DEL 		return;
//DEL 	}
//DEL 
//DEL 	HTREEITEM hRoot=m_treectrl.GetRootItem();
//DEL 	HTREEITEM hFind=m_treectrl.FindWindow(hRoot,strName);
//DEL 	if (hFind==NULL)
//DEL 	{
//DEL 		char msg[64]={0};
//DEL 		sprintf(msg,"���������������%s,ȷ��ô��",strName);
//DEL 		if (MessageBox(msg,_T("��ʾ"),MB_OKCANCEL)==IDOK)
//DEL 		{
//DEL 			*pResult=TRUE;
//DEL 		} 
//DEL 		else
//DEL 		{
//DEL 			m_treectrl.DeleteItem(m_hInsertItem);
//DEL 		}
//DEL 	}
//DEL 	else
//DEL 	{
//DEL 		AfxMessageBox(_T("���������Ѿ����ڣ����������룡"));
//DEL 		CEdit* pEdit=m_treectrl.EditLabel(m_hInsertItem);
//DEL 		ASSERT(pEdit!=NULL);
//DEL 
//DEL 	*pResult = 0;
//DEL 	}
//DEL 
//DEL }

//DEL void CDuanxinDlg::InitDriverTree()
//DEL {
//DEL 	char *pDriver,buf[50]={0};
//DEL 	//�õ����������̺�
//DEL 	GetLogicalDriveStrings(sizeof(buf),buf);
//DEL 	//���������Ϊ��
//DEL 	for(pDriver=buf;*pDriver;pDriver+=strlen(pDriver)+1)
//DEL 	{
//DEL 		//Ҷ�ӽڵ�ṹ��
//DEL 		TVINSERTSTRUCT tvInsert;
//DEL 		tvInsert.hParent = NULL;
//DEL 		tvInsert.hInsertAfter = NULL;
//DEL 		tvInsert.item.mask = TVIF_TEXT;
//DEL 		tvInsert.item.pszText = pDriver; 
//DEL 		
//DEL //		HTREEITEM hDriver = m_treectrl.InsertItem(&tvInsert);
//DEL 		//���ýڵ�����Ϊ1����ʾ�ýڵ��Ѿ�չ���������ٴ�չ��ʱ�����ٽ��а󶨣�
//DEL 		m_treectrl.SetItemData(hDriver,1);
//DEL 		//�Դ�������Ϊ���������²����ļ����а�
//DEL 		InsertNode(pDriver,hDriver);
//DEL      }
//DEL }

//DEL void CDuanxinDlg::InsertNode(CString szPath, HTREEITEM hNode)
//DEL {
//DEL 	HANDLE hFile;
//DEL 	WIN32_FIND_DATA wData;
//DEL 	
//DEL 	szPath+="\\*";
//DEL 	hFile=FindFirstFile(szPath,&wData);
//DEL 	//����ʧ��
//DEL 	if(hFile==INVALID_HANDLE_VALUE)
//DEL 		return;
//DEL 	do
//DEL 	{
//DEL 		//����2�������ļ���"."��".."
//DEL 		if(wData.cFileName[0]=='.')
//DEL 			continue;
//DEL 		//������ҵ����ļ��Ǹ��ļ���
//DEL 		if(wData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
//DEL 		{
//DEL 			HTREEITEM hTemp=m_treectrl.InsertItem(wData.cFileName,0,0,hNode,TVI_SORT);
//DEL 			//���һ����ʱ�ڵ�����ʾ+��
//DEL 			m_treectrl.InsertItem(NULL,0,0,hTemp,TVI_SORT); 
//DEL 		}
//DEL 		else
//DEL 			m_treectrl.InsertItem(wData.cFileName,0,0,hNode,TVI_SORT);
//DEL 		
//DEL      }while(FindNextFile(hFile,&wData));
//DEL 
//DEL }



//DEL void CDuanxinDlg::OnItemexpandingMytree(NMHDR* pNMHDR, LRESULT* pResult) 
//DEL {
//DEL 	NM_TREEVIEW* pNMTreeView = (NM_TREEVIEW*)pNMHDR;
//DEL 	// TODO: Add your control notification handler code here
//DEL 	//�ж���չ�����Ǻ�£
//DEL 	if(TVE_EXPAND==pNMTreeView->action)
//DEL 	{
//DEL 		HTREEITEM  hNode=pNMTreeView->itemNew.hItem; 
//DEL 		//�жϽڵ������Ƿ�Ϊ0����û��չ����������а�
//DEL 		if(!m_treectrl.GetItemData(hNode))
//DEL 		{
//DEL 			m_treectrl.DeleteItem(m_treectrl.GetChildItem(hNode));
//DEL 			m_treectrl.SetItemData(hNode,1);
//DEL 			InsertNode(GetFullPath(hNode),hNode);
//DEL 		}
//DEL      }
//DEL 
//DEL 	*pResult = 0;
//DEL }

//DEL CString CDuanxinDlg::GetFullPath(HTREEITEM hNode)
//DEL {
//DEL 	CString szRet=m_treectrl.GetItemText(hNode);
//DEL 	while(hNode=m_treectrl.GetParentItem(hNode))
//DEL 		szRet=m_treectrl.GetItemText(hNode)+"\\"+szRet;
//DEL      return szRet;
//DEL }

BEGIN_EVENTSINK_MAP(CDuanxinDlg, CDialog)
    //{{AFX_EVENTSINK_MAP(CDuanxinDlg)
	ON_EVENT(CDuanxinDlg, IDC_SMSGATE1, 2 /* OnRecvMsg */, OnOnRecvMsgSmsgate1, VTS_NONE)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void CDuanxinDlg::OnOnRecvMsgSmsgate1() 
{
	// TODO: Add your control notification handler code here
	if(com_set_flag==1)//�˿����Ӻ�ſ�ʼ������Ϣ
	{
	if (flg_receive_msg==0)//�¼�����
	{
		flg_receive_msg=1;
	VARIANT s_msg;
	CStringArray s_sa,s_message;//�洢ÿ����Ϣ��3���ֶ�,�洢ÿ����Ϣ
	int i=0,pos,message_counter=0;
 	CString a,b,buf;
	
 	a.Format("%c",'\002');
	b.Format("%c",'\001');
	
		
		s_msg=m_smsgate_1.NewMsg();

	if(s_msg.vt==VT_BSTR)
	{		
		s_newmessage=s_msg.bstrVal;//��ǰ��Ϣ��ȫ������
		while (1)//�õ�������ǰ�¼�����Ϣ��������
		{
			pos=s_newmessage.Find(b);
			if (pos>0)//������Ϣ
			{
				s_message.Add(s_newmessage.Left(pos));
				s_newmessage=s_newmessage.Mid(pos+1);
				message_counter++;
			}
			else//����һ����Ϣ�����߶�����Ϣ�����һ����Ϣ
			{
				s_message.Add(s_newmessage);
				message_counter++;
				break;
			}
			
		}//end of while(1)

		i=0;
		int p;
		for(int j=0;j<message_counter;j++)//��ÿ����Ϣ���δ���
		{
			p=0;
			buf=s_message.GetAt(j);
			//AfxMessageBox(buf,MB_OKCANCEL,0);
		while (1)//��ȡ���ĵ�����Ϣ����ȡ3������:���룬���ݣ�ʱ��
		{
			
			pos=buf.Find(a);
			if (pos>=0)//������Ϣ�ĵ�һ���ڶ��ֶ�
			{	p++;
				s_sa.Add(buf.Left(pos));
				buf=buf.Mid(pos+1);
				if(p%2==0)//3���еĵڶ��Σ���Ϣ�����ֶ�
				{
					//m_clistbox.InsertString(0,s_sa.GetAt(i));
					message_data=s_sa.GetAt(i);
					splitMessage();
				}
				i++;
			}
			else//������Ϣ�ĵ����ֶΣ�ʱ���ֶΣ�
			{
				p++;
				s_sa.Add(buf);
				i++;
				m_csrq=s_sa.GetAt(i);
				UpdateData(FALSE);
				break;
			}

		}//end of while(1)

		}//end of for	
		
	}
	else
	{
		AfxMessageBox("NOT VT_BSTR!",MB_OK,0);
	}
	flg_receive_msg=0;//�ͷŽ��̻���
	}
	}
}

//DEL void CDuanxinDlg::OnClose() 
//DEL {
//DEL 	// TODO: Add your message handler code here and/or call default
//DEL 	
//DEL 	CDialog::OnClose();
//DEL }

void CDuanxinDlg::OnDestroy() 
{
	m_smsgate_1.ClosePort();
	SysFreeString(bstr_s_mobile);
	SysFreeString(bstr_s_msg);
	CDialog::OnDestroy();
	AfxPostQuitMessage(0);
	// TODO: Add your message handler code here
	m_smsgate_1.ClosePort();
}

//DEL void CDuanxinDlg::OnDblclkList1() 
//DEL {
//DEL 	// TODO: Add your control notification handler code here
//DEL 	AfxMessageBox("ListBox click!",MB_OK,0);
//DEL }

void CDuanxinDlg::splitMessage()//��ÿ����Ϣ�ĺ������ݵĴ���
{
	CString divide_flag;
	int pos1=0,message_counter1=0,p1=0;//��Ϣ�жεĸ���
	CStringArray s_message1;//ÿ����Ϣ�������ԼӺ����֣��洢���ֳ�����ÿ�ε���Ϣ
	divide_flag.Format("%c",'\053');
//	m_clistbox.InsertString(0,message_data);
	int rightmessage_flag=0;
	bool CID_exist_flag=0;//BSIC���Ƿ�������б��еı�־λ
	while(1)
	{
	pos1=message_data.Find(divide_flag);
	if (pos1>=0)//ǰ14���ֶ�//if ((pos1>0)||(p1<14))//ǰ14���ֶ�
	{
//		if (message_data.Left(pos1).IsEmpty())
//		{
//			message_data=message_data.Mid(pos1+1);
//			continue;
//		}
		
		
		s_message1.Add(message_data.Left(pos1));
		message_data=message_data.Mid(pos1+1);
		//AfxMessageBox(s_message1.GetAt(message_counter1));
		if (p1==0)
		{
			
			m_cscs_cid=s_message1.GetAt(p1);
/****************����µ������Ϣ��CID���Ƿ�������б���***********************************/
			for (int ii=0;ii<data_data.GetSize();ii++)
			{
				if (m_cscs_cid==data_data.GetAt(ii))
				{
					CID_exist_flag=1;
					break;
				}
				else
					CID_exist_flag=0;
				
			}
/******************���CID�����********************************/	
			if (CID_exist_flag==0)
				m_clistbox.InsertString(0,s_message1.GetAt(p1));
			rightmessage_flag+=1;//�������������
		}
		else if (p1==1)
		{

			m_cscs_bsic=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==2)
		{
			m_cscs_lac=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		} 
		else if(p1==3)
		{
			m_cscs_pd=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==4)
		{
			m_cscs_jd=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==5)
		{
			m_cscs_wd=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==6)
		{
			m_cscs_fxj=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==7)
		{
			m_cscs_qj=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==8)
		{
			m_cscs_hgj=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==9)
		{
			m_cscs_dmhb=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==10)
		{
			m_cscs_txhb=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==11)
		{
			m_cscs_txgg=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==12)
		{
			m_cscs_qbdp=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}
		else if (p1==13)
		{
			m_cscs_hbdp=s_message1.GetAt(p1);
			rightmessage_flag+=1;//�������������
		}

		message_counter1++;
		p1++;
	}
	else if(rightmessage_flag==14)//���һ���ֶ�
	{
		s_message1.Add(message_data);
		m_cscs_dpbz=s_message1.GetAt(p1);
//		UpdateData(FALSE);//�˹��鿴����Ҫ�Զ�ˢ����ֵ
		message_counter1++;
		p1++;
		saveMessage();//���汾���յ�����Ϣ
		break;
	}
	else
	{
		AfxMessageBox("Received a garbage message!",MB_OK,0);
		break;
	}
	}//end of while(1)
}

void CDuanxinDlg::saveMessage()
{
	::WritePrivateProfileString(m_cscs_cid,"BSIC",m_cscs_bsic,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"CID",m_cscs_cid,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"LAC",m_cscs_lac,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"pinduan",m_cscs_pd,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"jingdu",m_cscs_jd,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"weidu",m_cscs_wd,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"fangxiangjiao",m_cscs_fxj,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"qingjiao",m_cscs_qj,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"henggunjiao",m_cscs_hgj,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"dimianhaiba",m_cscs_dmhb,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"tianxianhaiba",m_cscs_txhb,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"tianxianguagao",m_cscs_txgg,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"qianbandianping",m_cscs_qbdp,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"houbandianping",m_cscs_hbdp,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"dianpingbizhi",m_cscs_dpbz,currentpath_buf);
	::WritePrivateProfileString(m_cscs_cid,"date",m_csrq,currentpath_buf);

}

//DEL void CDuanxinDlg::OnSetfocusList1() 
//DEL {
//DEL 	// TODO: Add your control notification handler code here
//DEL 
//DEL }

//DEL void CDuanxinDlg::OnSelcancelList1() 
//DEL {
//DEL 	// TODO: Add your control notification handler code here
//DEL 
//DEL }

void CDuanxinDlg::OnSelchangeList1() 
{
	// TODO: Add your control notification handler code here
	CString str1,str2;
	CFile file;
	bool finded_bmp_flag=0;//�ҵ���Ӧ���ݵ�ͼƬ�ı�־λ
	int index=m_clistbox.GetCurSel();
	int j=0,i=0;
	if (index!=LB_ERR)
	{
		m_clistbox.GetText(index,str1);
		str2=str1;
		//AfxMessageBox(str1,MB_OK,0);
		::GetPrivateProfileString(str1,"BSIC","unknown",m_cscs_bsic.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_bsic.ReleaseBuffer();
		::GetPrivateProfileString(str1,"CID","unknown",m_cscs_cid.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_cid.ReleaseBuffer();
		::GetPrivateProfileString(str1,"LAC","unknown",m_cscs_lac.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_lac.ReleaseBuffer();
		::GetPrivateProfileString(str1,"pinduan","unknown",m_cscs_pd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_pd.ReleaseBuffer();
		::GetPrivateProfileString(str1,"jingdu","unknown",m_cscs_jd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_jd.ReleaseBuffer();
		::GetPrivateProfileString(str1,"weidu","unknown",m_cscs_wd.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_wd.ReleaseBuffer();
		::GetPrivateProfileString(str1,"fangxiangjiao","unknown",m_cscs_fxj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_fxj.ReleaseBuffer();
		::GetPrivateProfileString(str1,"qingjiao","unknown",m_cscs_qj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_qj.ReleaseBuffer();
		::GetPrivateProfileString(str1,"henggunjiao","unknown",m_cscs_hgj.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_hgj.ReleaseBuffer();
		::GetPrivateProfileString(str1,"dimianhaiba","unknown",m_cscs_dmhb.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_dmhb.ReleaseBuffer();
		::GetPrivateProfileString(str1,"tianxianhaiba","unknown",m_cscs_txhb.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_txhb.ReleaseBuffer();
		::GetPrivateProfileString(str1,"tianxianguagao","unknown",m_cscs_txgg.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_txgg.ReleaseBuffer();
		::GetPrivateProfileString(str1,"qianbandianping","unknown",m_cscs_qbdp.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_qbdp.ReleaseBuffer();
		::GetPrivateProfileString(str1,"houbandianping","unknown",m_cscs_hbdp.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_hbdp.ReleaseBuffer();
		::GetPrivateProfileString(str1,"dianpingbizhi","unknown",m_cscs_dpbz.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_dpbz.ReleaseBuffer();
		::GetPrivateProfileString(str1,"date","unknown",m_csrq.GetBuffer(MAX_PATH),MAX_PATH,currentpath_buf);
		m_cscs_dpbz.ReleaseBuffer();
		UpdateData(FALSE);

	}
/**********************������ͼƬ�Ĺ���********************************/
	str1=str2;
	
	for (i=0;i<data_bmp.GetSize();i++)
	{
		if (str1.GetLength()==(data_bmp.GetAt(i).GetLength()-4))
		{
			
			int pp=0;
			while (pp<str1.GetLength())
			{
				if ((data_bmp.GetAt(i).GetAt(pp)<str1.GetAt(pp))||(data_bmp.GetAt(i).GetAt(pp)>str1.GetAt(pp)))
				{
					break;
				}
				finded_bmp_flag=0;
				if (pp==str1.GetLength()-1)
				{
					m_listbox_bmp.SetCurSel(i);
					finded_bmp_flag=1;
					j=i;
//					AfxMessageBox(data_bmp.GetAt(i),MB_OK,0);
				}
				pp++;
			}
		}
		
	}//����ܹ��ҵ��Ļ�������ʱi��Ϊ��Ӧ��ͼƬ���б��е����к�
	
	if (finded_bmp_flag==1)
	{
		finded_bmp_flag=1;//�ҵ���ͼƬ
		CString str_tmp;
			m_listbox_bmp.GetText(j,str_tmp);		
			strFileName = m_strPath+str_tmp;
			
			
			if(!file.Open(strFileName, CFile::modeRead))
			{
				browse_flag=0;
				return;
			}
			//	AfxMessageBox(strFileName,MB_OK,0);
			browse_flag=1;//��ȷ��ͼƬ
			BITMAPFILEHEADER bmfHeader;
			nFileLen = file.GetLength();
			dwDibSize = nFileLen - sizeof(BITMAPFILEHEADER);
			if (m_pDib != NULL)
			{
				delete[] m_pDib;
				m_pDib = NULL;
			}
			m_pDib = new unsigned char[dwDibSize];
			if (file.Read((LPSTR)&bmfHeader, sizeof(bmfHeader)) != sizeof(bmfHeader))
				return;
			if (bmfHeader.bfType != ((WORD)('M'<<8) | 'B'))
				return ;
			if (file.Read(m_pDib, dwDibSize) != dwDibSize)
				return ;
			m_bmpInfoHeader = (BITMAPINFOHEADER*)m_pDib;
			lHeight = m_bmpInfoHeader->biHeight; //ͼ��
			lWidth = m_bmpInfoHeader->biWidth; //ͼ���
			lBitCount = m_bmpInfoHeader->biBitCount; //ͼ��λ��
			switch(lBitCount)
			{
			case 1:
				NumColor = 2;
				break;
			case 4:
				NumColor = 16;
				break;
			case 8:
				NumColor = 256;
				break;
			case 24:
				NumColor = 0;
				break;
			default:
				return;
			}
			m_pDibBits = m_pDib + sizeof(BITMAPINFOHEADER) + NumColor * sizeof (RGBQUAD);
			ShowBMP();		
	}
	else
	{		
		finded_bmp_flag=0;//û���ҵ�ͼƬ
		strFileName = str_current_Path+"\\bai.bmp";
		if(!file.Open(strFileName, CFile::modeRead))
		{
			browse_flag=0;
			return;
		}
		//	AfxMessageBox(strFileName,MB_OK,0);
		browse_flag=1;//��ȷ��ͼƬ
		BITMAPFILEHEADER bmfHeader;
		nFileLen = file.GetLength();
		dwDibSize = nFileLen - sizeof(BITMAPFILEHEADER);
		if (m_pDib != NULL)
		{
			delete[] m_pDib;
			m_pDib = NULL;
		}
		m_pDib = new unsigned char[dwDibSize];
		if (file.Read((LPSTR)&bmfHeader, sizeof(bmfHeader)) != sizeof(bmfHeader))
			return;
		if (bmfHeader.bfType != ((WORD)('M'<<8) | 'B'))
			return ;
		if (file.Read(m_pDib, dwDibSize) != dwDibSize)
			return ;
		m_bmpInfoHeader = (BITMAPINFOHEADER*)m_pDib;
		lHeight = m_bmpInfoHeader->biHeight; //ͼ��
		lWidth = m_bmpInfoHeader->biWidth; //ͼ���
		lBitCount = m_bmpInfoHeader->biBitCount; //ͼ��λ��
		switch(lBitCount)
		{
		case 1:
			NumColor = 2;
			break;
		case 4:
			NumColor = 16;
			break;
		case 8:
			NumColor = 256;
			break;
		case 24:
			NumColor = 0;
			break;
		default:
			return;
		}
		m_pDibBits = m_pDib + sizeof(BITMAPINFOHEADER) + NumColor * sizeof (RGBQUAD);
			ShowBMP();	
	}
}

void CDuanxinDlg::Onmodifyparam() 
{
	// TODO: Add your control notification handler code here
if (modify_flag==0)
	{
		modify_flag=1;
		GetDlgItem(IDC_BUTTON5)->SetWindowText("�������");
		GetDlgItem(IDC_EDIT9)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT14)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT7)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT34)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT13)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT8)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT17)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT18)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT20)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT26)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT28)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT27)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT29)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT30)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT31)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT37)->EnableWindow(TRUE);
	} 
	else
	{
		modify_flag=0;
		GetDlgItem(IDC_BUTTON5)->SetWindowText("�޸Ĳ���");
		GetDlgItem(IDC_EDIT9)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT14)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT7)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT34)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT13)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT8)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT17)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT18)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT20)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT26)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT28)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT27)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT29)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT30)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT31)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT37)->EnableWindow(FALSE);
		UpdateData(TRUE);

		::WritePrivateProfileString("config_param","BSIC",m_sqcs_bsic,currentpath_buf_para);
		::WritePrivateProfileString("config_param","CID",m_sqcs_cid,currentpath_buf_para);
		::WritePrivateProfileString("config_param","LAC",m_sqcs_lac,currentpath_buf_para);
		::WritePrivateProfileString("config_param","pinduan",m_sqcs_pd,currentpath_buf_para);
		::WritePrivateProfileString("config_param","jingdu",m_sqcs_jd,currentpath_buf_para);
		::WritePrivateProfileString("config_param","weidu",m_sqcs_wd,currentpath_buf_para);
		::WritePrivateProfileString("config_param","fangxiangjiao",m_sqcs_fxj,currentpath_buf_para);
		::WritePrivateProfileString("config_param","qingjiao",m_sqcs_qj,currentpath_buf_para);
		::WritePrivateProfileString("config_param","henggunjiao",m_sqcs_hgj,currentpath_buf_para);
		::WritePrivateProfileString("config_param","dimianhaiba",m_sqcs_dmhb,currentpath_buf_para);
		::WritePrivateProfileString("config_param","tianxianhaiba",m_sqcs_txhb,currentpath_buf_para);
		::WritePrivateProfileString("config_param","tianxianguagao",m_sqcs_txgg,currentpath_buf_para);
		::WritePrivateProfileString("config_param","qianbandianping",m_sqcs_qbdp,currentpath_buf_para);
		::WritePrivateProfileString("config_param","houbandianping",m_sqcs_hbdp,currentpath_buf_para);
		::WritePrivateProfileString("config_param","dianpingbizhi",m_sqcs_dpbz,currentpath_buf_para);
		::WritePrivateProfileString("config_param","local_telnumber",m_local_telnumber,currentpath_buf_para);
	}
}

void CDuanxinDlg::OnSavePath() 
{
	// TODO: Add your control notification handler code here

	// 		CFileDialog file_open_dlg(TRUE,NULL,NULL,OFN_HIDEREADONLY,"BMP Files(*.bmp)|*.bmp|All Files(*.*)|*.*||");
	// 		file_open_dlg.m_ofn.lpstrTitle = "Open Image File";
	// 		if(file_open_dlg.DoModal() != IDOK)
	// 		return;
// 		CString file_name = file_open_dlg.GetPathName();
	((CListBox*)GetDlgItem(IDC_LIST2))->ResetContent();//���ʱ���ԭ��CListBox�е�����
	data_bmp.RemoveAll();
	CString str;
	BROWSEINFO bi;
	char name[MAX_PATH];
	ZeroMemory(&bi,sizeof(BROWSEINFO));
	bi.hwndOwner = GetSafeHwnd();
	bi.pszDisplayName = name;
	bi.lpszTitle = "��ѡ��ͼƬĿ¼";
	bi.ulFlags = BIF_RETURNFSANCESTORS;
	LPITEMIDLIST idl = SHBrowseForFolder(&bi);
	if(idl == NULL)
		return;
	SHGetPathFromIDList(idl, str.GetBuffer(MAX_PATH));//����Ŀ��־���б�ת��Ϊ�ĵ�ϵͳ·��
	str.ReleaseBuffer();
	m_strPath = str;//Ϊ�Ի�������һ�༭���Ӧ��CString�ͱ��������沢��ʾѡ�е�·����
	if(str.GetAt(str.GetLength()-1)!='\\')
	m_strPath+="\\";
	m_save_path=m_strPath;
	UpdateData(FALSE);

//	if (savepath_flag==0)//����·��
//	{
// 		savepath_flag=1;
// 		GetDlgItem(IDC_BUTTON2)->EnableWindow(TRUE);
// 		GetDlgItem(IDC_BUTTON3)->EnableWindow(TRUE);
//	} 
// 	else//��������·��
// 	{
// 		GetDlgItem(IDC_BUTTON4)->SetWindowText(_T("����·��"));
// 	}
/********************���ͼƬ�����б�************************/

	HANDLE hFind_txt;
	WIN32_FIND_DATA FindFileData;//Ѱ���ļ���־
	CString m_strFolder=m_strPath+"*.bmp";//�����ҵ�Ŀ¼
	CString str_Folder=m_strFolder;//������չ��Ϊtxt���ļ�
	hFind_txt = FindFirstFile(str_Folder,&FindFileData);
//	CString strFileName;//�洢�ļ���
	if(hFind_txt != INVALID_HANDLE_VALUE)
	{
		//�鵽�ĵ�һ���ļ�
		strFileName=FindFileData.cFileName;
		m_listbox_bmp.AddString(strFileName);
		data_bmp.Add(strFileName);
//		cout<<strFileName<<endl;
		
	}
	while(FindNextFile(hFind_txt,&FindFileData))
	{
		strFileName=FindFileData.cFileName;
		m_listbox_bmp.AddString(strFileName);
		data_bmp.Add(strFileName);
//		cout<<strFileName<<endl;
	}
	FindClose(hFind_txt);
	m_listbox_bmp.SetCurSel(0);
// 	for (int p=0;p<data_bmp.GetSize();p++)
// 	{
// 		AfxMessageBox(data_bmp.GetAt(p),MB_OK,0);
// 	}
}



void CDuanxinDlg::OnAnalysisiReport() 
{
	// TODO: Add your control notification handler code here
	/********************�γ�����***********************/
	UpdateData(TRUE);
	CStringArray str_array;
	str_array.Add("������Ա:");str_array.Add(m_csjsy);str_array.Add("��������:");str_array.Add(m_csrq);
	str_array.Add("��Ŀ����");str_array.Add("�������");str_array.Add("���Բ���");
	str_array.Add("��վ����:");str_array.Add(m_jzmc);
	str_array.Add("��������:");str_array.Add(m_txmc);
	str_array.Add("BSIC");str_array.Add(m_sqcs_bsic);str_array.Add(m_cscs_bsic);
	str_array.Add("CID");str_array.Add(m_sqcs_cid);str_array.Add(m_cscs_cid);
	str_array.Add("LAC");str_array.Add(m_sqcs_lac);str_array.Add(m_cscs_lac);
	str_array.Add("Ƶ��");str_array.Add(m_sqcs_pd);str_array.Add(m_cscs_pd);
	str_array.Add("����");str_array.Add(m_sqcs_jd);str_array.Add(m_cscs_jd);
	str_array.Add("γ��");str_array.Add(m_sqcs_wd);str_array.Add(m_cscs_wd);
	str_array.Add("�����");str_array.Add(m_sqcs_fxj);str_array.Add(m_cscs_fxj);
	str_array.Add("���");str_array.Add(m_sqcs_qj);str_array.Add(m_cscs_qj);
	str_array.Add("�����");str_array.Add(m_sqcs_hgj);str_array.Add(m_cscs_hgj);
	str_array.Add("���溣��");str_array.Add(m_sqcs_dmhb);str_array.Add(m_cscs_dmhb);
	str_array.Add("���ߺ���");str_array.Add(m_sqcs_txhb);str_array.Add(m_cscs_txhb);
	str_array.Add("���߹Ҹ�");str_array.Add(m_sqcs_txgg);str_array.Add(m_cscs_txgg);
	str_array.Add("ǰ���ƽ");str_array.Add(m_sqcs_qbdp);str_array.Add(m_cscs_qbdp);
	str_array.Add("����ƽ");str_array.Add(m_sqcs_hbdp);str_array.Add(m_cscs_hbdp);
	str_array.Add("��ƽ��ֵ");str_array.Add(m_sqcs_dpbz);str_array.Add(m_cscs_dpbz);

	_Application app;
	COleVariant vTrue((short)TRUE),	vFalse((short)FALSE);
	COleVariant   VOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
	app.CreateDispatch(_T("Word.Application"));
	app.SetVisible(FALSE);
	//Create New Doc
	Documents docs=app.GetDocuments();
	CComVariant tpl(_T("")),Visble,DocType(0),NewTemplate(false);
	docs.Add(&tpl,&NewTemplate,&DocType,&Visble);
	//Add Content:Text
	Selection sel=app.GetSelection();
	sel.TypeText(_T("\t\t\t\t\t\t\t������̬���Է�������\r\n"));
	// COleDateTime dt=COleDateTime::GetCurrentTime();
	// CString strDT=dt.Format("%Y-%m-%d");
	// CString str("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
	// str+=strDT;
	// str+="\r\n";
	// sel.TypeText(str);
	//Add Table
	_Document saveDoc=app.GetActiveDocument();
	Tables tables=saveDoc.GetTables();
	CComVariant defaultBehavior(1),AutoFitBehavior(1);
	tables.Add(sel.GetRange(),21,4,&defaultBehavior,&AutoFitBehavior);
	Table table=tables.Item(1);
	/*************���********************************/
	Cell c1=table.Cell(2,3);/***�ڶ���***/
	Cell c2=table.Cell(2,4);
	c1.Merge(c2);
	
	for(int i=3;i<5;i++)
	{
		c1=table.Cell(i,2);/***��3,4��***/
		c2=table.Cell(i,3);
		Cell c3=table.Cell(i,4);
		c1.Merge(c2);
		c1.Merge(c3);
		c3.ReleaseDispatch();
	}
	
	for (i=5;i<20;i++)
	{
		c1=table.Cell(i,3);/***��5~19��***/
		c2=table.Cell(i,4);
		c1.Merge(c2);
	}
	
	for (i=20;i<=21;i++)
	{
		c1=table.Cell(i,2);/***��20,21��***/
		c2=table.Cell(i,3);
		Cell c3=table.Cell(i,4);
		c1.Merge(c2);
		c1.Merge(c3);
		Cell c4=table.Cell(i,1);
		c4.Merge(c1);
		c3.ReleaseDispatch();
		c4.ReleaseDispatch();
	}
	
	c1.ReleaseDispatch();
	c2.ReleaseDispatch();
	
	/************����*********************************/
	for (i=0;i<str_array.GetSize();i++)
	{
		sel.TypeText(_T(str_array.GetAt(i)));//��1��
		sel.MoveRight(COleVariant((short)1),COleVariant(short(1)),COleVariant(short(0)));
	}
//	sel.TypeText(_T("��ע��Ϣ:"));//��ע��Ϣ
	CString bzxx_tmp="��ע��Ϣ:"+m_bzxx;
	sel.TypeText(_T(bzxx_tmp));//��ע��Ϣ
	sel.MoveRight(COleVariant((short)1),COleVariant(short(1)),COleVariant(short(0)));	
	/********************����ͼƬ**************************/
	sel.TypeText(_T("��վͼƬ"));
	if((browse_flag==1)&&(finded_bmp_flag==1))
	{
		InlineShapes inlineshapes = sel.GetInlineShapes();
		CString picture1=strFileName;
		inlineshapes.AddPicture((LPCTSTR)picture1,COleVariant((short)FALSE),COleVariant((short)TRUE),&_variant_t(sel.GetRange()));
		inlineshapes.ReleaseDispatch();
	}
	
	CString  final_save_path_doc=m_save_path;
	final_save_path_doc+=m_cscs_cid;
	final_save_path_doc+=".doc";

	saveDoc=app.GetActiveDocument();
//	saveDoc.SaveAs(COleVariant(final_save_path_doc),COleVariant((short)0),vFalse, COleVariant(""), vTrue, COleVariant(""),vFalse, vFalse, vFalse, vFalse, vFalse,VOptional, VOptional, VOptional, VOptional, VOptional);

	app.SetVisible(TRUE);
	table.ReleaseDispatch();
	tables.ReleaseDispatch();
	sel.ReleaseDispatch();
	docs.ReleaseDispatch();
	saveDoc.ReleaseDispatch();
	app.SetVisible(TRUE);
	app.ReleaseDispatch();
/**********************excel���ֿ�ʼ*****************************************************************************/
	_Applicationexcel app_excel;
	Workbooks books;
	_Workbook book;
	Worksheets sheets;
	_Worksheet sheet;
	Rangeexcel range;
	Rangeexcel cols;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CStringArray excel_strarray;
	excel_strarray.Add("A2");excel_strarray.Add("B2");excel_strarray.Add("C2");excel_strarray.Add("D2");
	excel_strarray.Add("A3");excel_strarray.Add("B3");excel_strarray.Add("C3");
	excel_strarray.Add("A4");excel_strarray.Add("B4");
	excel_strarray.Add("A5");excel_strarray.Add("B5");
	excel_strarray.Add("A6");excel_strarray.Add("B6");excel_strarray.Add("C6");
	excel_strarray.Add("A7");excel_strarray.Add("B7");excel_strarray.Add("C7");
	excel_strarray.Add("A8");excel_strarray.Add("B8");excel_strarray.Add("C8");
	excel_strarray.Add("A9");excel_strarray.Add("B9");excel_strarray.Add("C9");
	excel_strarray.Add("A10");excel_strarray.Add("B10");excel_strarray.Add("C10");
	excel_strarray.Add("A11");excel_strarray.Add("B11");excel_strarray.Add("C11");
	excel_strarray.Add("A12");excel_strarray.Add("B12");excel_strarray.Add("C12");
	excel_strarray.Add("A13");excel_strarray.Add("B13");excel_strarray.Add("C13");
	excel_strarray.Add("A14");excel_strarray.Add("B14");excel_strarray.Add("C14");
	excel_strarray.Add("A15");excel_strarray.Add("B15");excel_strarray.Add("C15");
	excel_strarray.Add("A16");excel_strarray.Add("B16");excel_strarray.Add("C16");
	excel_strarray.Add("A17");excel_strarray.Add("B17");excel_strarray.Add("C17");
	excel_strarray.Add("A18");excel_strarray.Add("B18");excel_strarray.Add("C18");
	excel_strarray.Add("A19");excel_strarray.Add("B19");excel_strarray.Add("C19");
	excel_strarray.Add("A20");excel_strarray.Add("B20");excel_strarray.Add("C20");

	if( !app_excel.CreateDispatch("Excel.Application")){
		this->MessageBox("�޷�����ExcelӦ�ã�");
		return;
	}
	
	books=app_excel.GetWorkbooks();//��ȡ����������
	book=books.Add(covOptional); //���һ��������
	sheets=book.GetSheets();//��ȡ��������
	sheet=sheets.GetItem(COleVariant((short)1));//��ȡ��һ��������
/****************************ѭ����������*********************************************/
	for (i=0;i<str_array.GetSize();i++)
	{
		range=sheet.GetRange(COleVariant(excel_strarray.GetAt(i)),COleVariant(excel_strarray.GetAt(i)));//��2��
		range.SetValue2(COleVariant(str_array.GetAt(i)));
		cols=range.GetEntireColumn();//���ÿ��Ϊ�Զ���Ӧ
		cols.AutoFit();
	}

	range=sheet.GetRange(COleVariant("B1"),COleVariant("B1"));//ѡ��������A1:A1��Ԫ������
	range.SetValue2(COleVariant("������̬���Է�������"));//���ñ���

	CString bzxx_excel_tmp="��ע��Ϣ:"+m_bzxx;
	range=sheet.GetRange(COleVariant("A21"),COleVariant("A21"));//ѡ��������A1:A1��Ԫ������
	range.SetValue2(COleVariant(bzxx_excel_tmp));//��ע��Ϣ
	/********************����ͼƬ**************************/
	if((browse_flag==1)&&(finded_bmp_flag==1))
	{
	Shapesexcel shapesexcel=sheet.GetShapes();//��Sheet�����ϻ��һ��Shapes 
	CString picture2=strFileName;
	range=sheet.GetRange(COleVariant(_T("A22")),COleVariant(_T("H28")));    // ���Range������������ͼƬ
	shapesexcel.AddPicture(_T(picture2),false,true,0, 300,400, 300);
	}
/****************************ѭ���������ݣ�������***************************************/	
	app_excel.SetVisible(TRUE);//��ʾExcel��񣬲�����״̬Ϊ�û��ɿ���
	app_excel.SetUserControl(TRUE);
	
	//ͨ��Workbook�����SaveAs��������ʵ�ֱ���
	CString  final_save_path_excel=m_save_path;
	final_save_path_excel+=m_cscs_cid;
	final_save_path_excel+=".xls";
//	book.SaveAs(COleVariant(final_save_path_excel),covOptional,covOptional,covOptional,covOptional,covOptional,(long)0,covOptional,covOptional,covOptional,covOptional,covOptional);
	
app_excel.ReleaseDispatch();
books.ReleaseDispatch();
book.ReleaseDispatch();
sheets.ReleaseDispatch();
sheet.ReleaseDispatch();
range.ReleaseDispatch();
cols.ReleaseDispatch();
/***********************excel���ֽ���***************************************************************************/
}

void CDuanxinDlg::OnBrowse() 
{	
	CFileDialog dlg(TRUE, "*.BMP", NULL, NULL,"λͼ�ļ�(*.BMP)|*.bmp;*.BMP|");	
	CFile file;
	if (dlg.DoModal() == IDOK)
	{
		strFileName = dlg.GetPathName();
		if(!file.Open(strFileName, CFile::modeRead))
		{
			browse_flag=0;
			return;
		}
		//	AfxMessageBox(strFileName,MB_OK,0);
		browse_flag=1;//��ȷ��ͼƬ
		BITMAPFILEHEADER bmfHeader;
		nFileLen = file.GetLength();
		dwDibSize = nFileLen - sizeof(BITMAPFILEHEADER);
		if (m_pDib != NULL)
		{
			delete[] m_pDib;
			m_pDib = NULL;
		}
		m_pDib = new unsigned char[dwDibSize];
		if (file.Read((LPSTR)&bmfHeader, sizeof(bmfHeader)) != sizeof(bmfHeader))
			return;
		if (bmfHeader.bfType != ((WORD)('M'<<8) | 'B'))
			return ;
		if (file.Read(m_pDib, dwDibSize) != dwDibSize)
			return ;
		m_bmpInfoHeader = (BITMAPINFOHEADER*)m_pDib;
		lHeight = m_bmpInfoHeader->biHeight; //ͼ��
		lWidth = m_bmpInfoHeader->biWidth; //ͼ���
		lBitCount = m_bmpInfoHeader->biBitCount; //ͼ��λ��
		switch(lBitCount)
		{
		case 1:
			NumColor = 2;
			break;
		case 4:
			NumColor = 16;
			break;
		case 8:
			NumColor = 256;
			break;
		case 24:
			NumColor = 0;
			break;
		default:
			return;
		}
		m_pDibBits = m_pDib + sizeof(BITMAPINFOHEADER) + NumColor * sizeof (RGBQUAD);
		
	}
	else
	{
		browse_flag=0;
	}
	
ShowBMP();	
}

void CDuanxinDlg::ShowBMP()
{
	CDC *pDC;
	CRect rect;
	CWnd *pWnd = GetDlgItem(IDC_STATIC_SHOW);
	pWnd->GetClientRect(&rect);
	pDC = pWnd->GetDC();
	SetStretchBltMode(pDC->m_hDC,HALFTONE);//��ֹ����Ӧ����ͼ����ʾʧ��
	StretchDIBits(pDC->m_hDC,rect.left, rect.top, rect.right, rect.bottom, 0, 0,lWidth, lHeight, m_pDibBits, (BITMAPINFO*)m_bmpInfoHeader, BI_RGB, SRCCOPY);	
}

void CDuanxinDlg::OnConnectComport() 
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_BUTTON8)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON9)->EnableWindow(TRUE);


 	m_smsgate_1.SetCommPort(m_comport.GetCurSel()+1);
// 	m_smsgate_1.SetSmsService("+8613800290500");
// 	m_smsgate_1.SetSettings("9600,n,8,1");
	if (RevAuto_once_flag==0)
	{
		RevAuto_once_flag=1;
		m_smsgate_1.RevAuto();
	}
 	
// 	m_smsgate_1.SetReadAndDel(TRUE);
	waittime=10;
	s_wait_time=m_smsgate_1.Connect(&waittime);


	if(s_wait_time.vt==VT_BSTR)
	{
		CString com_connect_feedback=s_wait_time.bstrVal;
		if(com_connect_feedback=="y")//�Ѿ���ȷ����
		{
			m_com_openoff.SetIcon(m_hIconRed);
			com_set_flag=1;//com������
		}		
	}
	
	
	
}

void CDuanxinDlg::OnDisconnectComport() 
{
	// TODO: Add your control notification handler code here
	com_set_flag=0;
	m_smsgate_1.ClosePort();
	m_com_openoff.SetIcon(m_hIconOff);
	GetDlgItem(IDC_BUTTON9)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON8)->EnableWindow(TRUE);
}

// void CDuanxinDlg::OnSelchangeList2() 
// {
// 	// TODO: Add your control notification handler code here
// 	CString str1;
// 	int index=m_listbox_bmp.GetCurSel();
// 	if (index!=LB_ERR)
// 	{
// 		m_listbox_bmp.GetText(index,str1);
// 
// 
// 
// 	strFileName = m_strPath+str1;
// 	AfxMessageBox(strFileName,MB_OK,0);
// 	if(!file.Open(strFileName, CFile::modeRead))
// 	{
// 		browse_flag=0;
// 		return;
// 	}
// 	//	AfxMessageBox(strFileName,MB_OK,0);
// 	browse_flag=1;//��ȷ��ͼƬ
// 	BITMAPFILEHEADER bmfHeader;
// 	nFileLen = file.GetLength();
// 	dwDibSize = nFileLen - sizeof(BITMAPFILEHEADER);
// 	if (m_pDib != NULL)
// 	{
// 		delete[] m_pDib;
// 		m_pDib = NULL;
// 	}
// 	m_pDib = new unsigned char[dwDibSize];
// 	if (file.Read((LPSTR)&bmfHeader, sizeof(bmfHeader)) != sizeof(bmfHeader))
// 		return;
// 	if (bmfHeader.bfType != ((WORD)('M'<<8) | 'B'))
// 		return ;
// 	if (file.Read(m_pDib, dwDibSize) != dwDibSize)
// 		return ;
// 	m_bmpInfoHeader = (BITMAPINFOHEADER*)m_pDib;
// 	lHeight = m_bmpInfoHeader->biHeight; //ͼ��
// 	lWidth = m_bmpInfoHeader->biWidth; //ͼ���
// 	lBitCount = m_bmpInfoHeader->biBitCount; //ͼ��λ��
// 	switch(lBitCount)
// 	{
// 	case 1:
// 		NumColor = 2;
// 		break;
// 	case 4:
// 		NumColor = 16;
// 		break;
// 	case 8:
// 		NumColor = 256;
// 		break;
// 	case 24:
// 		NumColor = 0;
// 		break;
// 	default:
// 		return;
// 	}
// 		m_pDibBits = m_pDib + sizeof(BITMAPINFOHEADER) + NumColor * sizeof (RGBQUAD);
// 	}
// }

void CDuanxinDlg::CalcWindowRect(LPRECT lpClientRect, UINT nAdjustType) 
{
	// TODO: Add your specialized code here and/or call the base class
	
	CDialog::CalcWindowRect(lpClientRect, nAdjustType);
}

void CDuanxinDlg::OnSelchangeList2() 
{
	// TODO: Add your control notification handler code here
	CString str1;
	CFile file;

	int index=m_listbox_bmp.GetCurSel();
	if (index!=LB_ERR)
	{
		m_listbox_bmp.GetText(index,str1);		
		strFileName = m_strPath+str1;


		if(!file.Open(strFileName, CFile::modeRead))
		{
			browse_flag=0;
			return;
		}
		//	AfxMessageBox(strFileName,MB_OK,0);
		browse_flag=1;//��ȷ��ͼƬ
		BITMAPFILEHEADER bmfHeader;
		nFileLen = file.GetLength();
		dwDibSize = nFileLen - sizeof(BITMAPFILEHEADER);
		if (m_pDib != NULL)
		{
			delete[] m_pDib;
			m_pDib = NULL;
		}
		m_pDib = new unsigned char[dwDibSize];
		if (file.Read((LPSTR)&bmfHeader, sizeof(bmfHeader)) != sizeof(bmfHeader))
			return;
		if (bmfHeader.bfType != ((WORD)('M'<<8) | 'B'))
			return ;
		if (file.Read(m_pDib, dwDibSize) != dwDibSize)
			return ;
		m_bmpInfoHeader = (BITMAPINFOHEADER*)m_pDib;
		lHeight = m_bmpInfoHeader->biHeight; //ͼ��
		lWidth = m_bmpInfoHeader->biWidth; //ͼ���
		lBitCount = m_bmpInfoHeader->biBitCount; //ͼ��λ��
		switch(lBitCount)
		{
		case 1:
			NumColor = 2;
			break;
		case 4:
			NumColor = 16;
			break;
		case 8:
			NumColor = 256;
			break;
		case 24:
			NumColor = 0;
			break;
		default:
			return;
		}
		m_pDibBits = m_pDib + sizeof(BITMAPINFOHEADER) + NumColor * sizeof (RGBQUAD);
		ShowBMP();	
	}	
}

void CDuanxinDlg::OnTestReport() 
{
	// TODO: Add your control notification handler code here
		UpdateData(TRUE);
	CStringArray str_array;
	str_array.Add("������Ա:");str_array.Add(m_csjsy);str_array.Add("��������:");str_array.Add(m_csrq);
	str_array.Add("��Ŀ����");str_array.Add("���Բ���");
	str_array.Add("��վ����:");str_array.Add(m_jzmc);
	str_array.Add("��������:");str_array.Add(m_txmc);
	str_array.Add("BSIC");str_array.Add(m_cscs_bsic);
	str_array.Add("CID");str_array.Add(m_cscs_cid);
	str_array.Add("LAC");str_array.Add(m_cscs_lac);
	str_array.Add("Ƶ��");str_array.Add(m_cscs_pd);
	str_array.Add("����");str_array.Add(m_cscs_jd);
	str_array.Add("γ��");str_array.Add(m_cscs_wd);
	str_array.Add("�����");str_array.Add(m_cscs_fxj);
	str_array.Add("���");str_array.Add(m_cscs_qj);
	str_array.Add("�����");str_array.Add(m_cscs_hgj);
	str_array.Add("���溣��");str_array.Add(m_cscs_dmhb);
	str_array.Add("���ߺ���");str_array.Add(m_cscs_txhb);
	str_array.Add("���߹Ҹ�");str_array.Add(m_cscs_txgg);
	str_array.Add("ǰ���ƽ");str_array.Add(m_cscs_qbdp);
	str_array.Add("����ƽ");str_array.Add(m_cscs_hbdp);
	str_array.Add("��ƽ��ֵ");str_array.Add(m_cscs_dpbz);

	_Application app;
	COleVariant vTrue((short)TRUE),	vFalse((short)FALSE);
	COleVariant   VOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
	app.CreateDispatch(_T("Word.Application"));
	app.SetVisible(FALSE);
	//Create New Doc
	Documents docs=app.GetDocuments();
	CComVariant tpl(_T("")),Visble,DocType(0),NewTemplate(false);
	docs.Add(&tpl,&NewTemplate,&DocType,&Visble);
	//Add Content:Text
	Selection sel=app.GetSelection();
	sel.TypeText(_T("\t\t\t\t\t\t\t������̬���Է�������\r\n"));
	// COleDateTime dt=COleDateTime::GetCurrentTime();
	// CString strDT=dt.Format("%Y-%m-%d");
	// CString str("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
	// str+=strDT;
	// str+="\r\n";
	// sel.TypeText(str);
	//Add Table
	_Document saveDoc=app.GetActiveDocument();
	Tables tables=saveDoc.GetTables();
	CComVariant defaultBehavior(1),AutoFitBehavior(1);
	tables.Add(sel.GetRange(),21,4,&defaultBehavior,&AutoFitBehavior);
	Table table=tables.Item(1);
	/*************���********************************/
	Cell c1=table.Cell(2,3);/***�ڶ���***/
	Cell c2=table.Cell(2,4);
	
	for(int i=2;i<20;i++)
	{
		c1=table.Cell(i,2);/***��3,4��***/
		c2=table.Cell(i,3);
		Cell c3=table.Cell(i,4);
		c1.Merge(c2);
		c1.Merge(c3);
		c3.ReleaseDispatch();
	}
	
	for (i=20;i<=21;i++)
	{
		c1=table.Cell(i,2);/***��20,21��***/
		c2=table.Cell(i,3);
		Cell c3=table.Cell(i,4);
		c1.Merge(c2);
		c1.Merge(c3);
		Cell c4=table.Cell(i,1);
		c4.Merge(c1);
		c3.ReleaseDispatch();
		c4.ReleaseDispatch();
	}
	
	c1.ReleaseDispatch();
	c2.ReleaseDispatch();
	
	/************����*********************************/
	for (i=0;i<str_array.GetSize();i++)
	{
		sel.TypeText(_T(str_array.GetAt(i)));//��1��
		sel.MoveRight(COleVariant((short)1),COleVariant(short(1)),COleVariant(short(0)));
	}
//	sel.TypeText(_T("��ע��Ϣ:"));//��ע��Ϣ
	CString bzxx_tmp="��ע��Ϣ:"+m_bzxx;
	sel.TypeText(_T(bzxx_tmp));//��ע��Ϣ
	sel.MoveRight(COleVariant((short)1),COleVariant(short(1)),COleVariant(short(0)));	
	/********************����ͼƬ**************************/
	sel.TypeText(_T("��վͼƬ"));
	if((browse_flag==1)&&(finded_bmp_flag==1))
	{
		InlineShapes inlineshapes = sel.GetInlineShapes();
		CString picture1=strFileName;
		inlineshapes.AddPicture((LPCTSTR)picture1,COleVariant((short)FALSE),COleVariant((short)TRUE),&_variant_t(sel.GetRange()));
		inlineshapes.ReleaseDispatch();
	}
	
	CString  final_save_path_doc=m_save_path;
	final_save_path_doc+=m_cscs_cid;
	final_save_path_doc+=".doc";

	saveDoc=app.GetActiveDocument();
//	saveDoc.SaveAs(COleVariant(final_save_path_doc),COleVariant((short)0),vFalse, COleVariant(""), vTrue, COleVariant(""),vFalse, vFalse, vFalse, vFalse, vFalse,VOptional, VOptional, VOptional, VOptional, VOptional);

	app.SetVisible(TRUE);
	table.ReleaseDispatch();
	tables.ReleaseDispatch();
	sel.ReleaseDispatch();
	docs.ReleaseDispatch();
	saveDoc.ReleaseDispatch();
	app.SetVisible(TRUE);
	app.ReleaseDispatch();
/**********************excel���ֿ�ʼ*****************************************************************************/
	_Applicationexcel app_excel;
	Workbooks books;
	_Workbook book;
	Worksheets sheets;
	_Worksheet sheet;
	Rangeexcel range;
	Rangeexcel cols;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CStringArray excel_strarray;
	excel_strarray.Add("A2");excel_strarray.Add("B2");excel_strarray.Add("C2");excel_strarray.Add("D2");
	excel_strarray.Add("A3");excel_strarray.Add("B3");
	excel_strarray.Add("A4");excel_strarray.Add("B4");
	excel_strarray.Add("A5");excel_strarray.Add("B5");
	excel_strarray.Add("A6");excel_strarray.Add("B6");
	excel_strarray.Add("A7");excel_strarray.Add("B7");
	excel_strarray.Add("A8");excel_strarray.Add("B8");
	excel_strarray.Add("A9");excel_strarray.Add("B9");
	excel_strarray.Add("A10");excel_strarray.Add("B10");
	excel_strarray.Add("A11");excel_strarray.Add("B11");
	excel_strarray.Add("A12");excel_strarray.Add("B12");
	excel_strarray.Add("A13");excel_strarray.Add("B13");
	excel_strarray.Add("A14");excel_strarray.Add("B14");
	excel_strarray.Add("A15");excel_strarray.Add("B15");
	excel_strarray.Add("A16");excel_strarray.Add("B16");
	excel_strarray.Add("A17");excel_strarray.Add("B17");
	excel_strarray.Add("A18");excel_strarray.Add("B18");
	excel_strarray.Add("A19");excel_strarray.Add("B19");
	excel_strarray.Add("A20");excel_strarray.Add("B20");

	if( !app_excel.CreateDispatch("Excel.Application")){
		this->MessageBox("�޷�����ExcelӦ�ã�");
		return;
	}
	
	books=app_excel.GetWorkbooks();//��ȡ����������
	book=books.Add(covOptional); //���һ��������
	sheets=book.GetSheets();//��ȡ��������
	sheet=sheets.GetItem(COleVariant((short)1));//��ȡ��һ��������
/****************************ѭ����������*********************************************/
	for (i=0;i<str_array.GetSize();i++)
	{
		range=sheet.GetRange(COleVariant(excel_strarray.GetAt(i)),COleVariant(excel_strarray.GetAt(i)));//��2��
		range.SetValue2(COleVariant(str_array.GetAt(i)));
		cols=range.GetEntireColumn();//���ÿ��Ϊ�Զ���Ӧ
		cols.AutoFit();
	}

	range=sheet.GetRange(COleVariant("B1"),COleVariant("B1"));//ѡ��������A1:A1��Ԫ������
	range.SetValue2(COleVariant("������̬���Է�������"));//���ñ���

	CString bzxx_excel_tmp="��ע��Ϣ:"+m_bzxx;
	range=sheet.GetRange(COleVariant("A21"),COleVariant("A21"));//ѡ��������A1:A1��Ԫ������
	range.SetValue2(COleVariant(bzxx_excel_tmp));//��ע��Ϣ
	/********************����ͼƬ**************************/
	if((browse_flag==1)&&(finded_bmp_flag==1))
	{
	Shapesexcel shapesexcel=sheet.GetShapes();//��Sheet�����ϻ��һ��Shapes 
	CString picture2=strFileName;
	range=sheet.GetRange(COleVariant(_T("A22")),COleVariant(_T("H28")));    // ���Range������������ͼƬ
	shapesexcel.AddPicture(_T(picture2),false,true,0, 300,400, 300);
	}
/****************************ѭ���������ݣ�������***************************************/	
	app_excel.SetVisible(TRUE);//��ʾExcel��񣬲�����״̬Ϊ�û��ɿ���
	app_excel.SetUserControl(TRUE);
	
	//ͨ��Workbook�����SaveAs��������ʵ�ֱ���
	CString  final_save_path_excel=m_save_path;
	final_save_path_excel+=m_cscs_cid;
	final_save_path_excel+=".xls";
//	book.SaveAs(COleVariant(final_save_path_excel),covOptional,covOptional,covOptional,covOptional,covOptional,(long)0,covOptional,covOptional,covOptional,covOptional,covOptional);
	
app_excel.ReleaseDispatch();
books.ReleaseDispatch();
book.ReleaseDispatch();
sheets.ReleaseDispatch();
sheet.ReleaseDispatch();
range.ReleaseDispatch();
cols.ReleaseDispatch();
/***********************excel���ֽ���***************************************************************************/
}

void CDuanxinDlg::OnImportData() 
{
	// TODO: Add your control notification handler code here
	CString str;
	BROWSEINFO bi;
	char name[MAX_PATH];
	ZeroMemory(&bi,sizeof(BROWSEINFO));
	bi.hwndOwner = GetSafeHwnd();
	bi.pszDisplayName = name;
	bi.lpszTitle = "��ѡ������Ŀ¼";
	bi.ulFlags = BIF_RETURNFSANCESTORS;
	LPITEMIDLIST idl = SHBrowseForFolder(&bi);
	if(idl == NULL)
		return;
	SHGetPathFromIDList(idl, str.GetBuffer(MAX_PATH));//����Ŀ��־���б�ת��Ϊ�ĵ�ϵͳ·��
	str.ReleaseBuffer();
	m_strPath = str;//Ϊ�Ի�������һ�༭���Ӧ��CString�ͱ��������沢��ʾѡ�е�·����
	if(str.GetAt(str.GetLength()-1)!='\\')
		m_strPath+="\\";
	m_save_path=m_strPath;
	UpdateData(FALSE);
	/********************���ͼƬ�����б�************************/
	CStdioFile stdio_file;
	
	HANDLE hFind_txt;
	WIN32_FIND_DATA FindFileData;//Ѱ���ļ���־
	CString m_strFolder=m_strPath+"*.txt";//�����ҵ�Ŀ¼
	CString str_Folder=m_strFolder;//������չ��Ϊtxt���ļ�
	hFind_txt = FindFirstFile(str_Folder,&FindFileData);
	//	CString strFileName;//�洢�ļ���
	if(hFind_txt != INVALID_HANDLE_VALUE)
	{
		//�鵽�ĵ�һ���ļ�
		strFileName=FindFileData.cFileName;
		data_txt.Add(strFileName);
/*******************�������������ݲ���*******************************/
		CString str_tmp_1=m_strPath+strFileName;
		if(!stdio_file.Open(str_tmp_1,CFile::modeRead))
		{
			AfxMessageBox("�򿪴���׼���˳���",MB_OK,0);
			return;
		}
		stdio_file.ReadString(str_tmp_1);
//		AfxMessageBox(str_tmp_1);
		message_data=str_tmp_1;
		splitMessage();
		//�ر��ļ�		
		stdio_file.Close();
	}
	while(FindNextFile(hFind_txt,&FindFileData))
	{
		strFileName=FindFileData.cFileName;
		data_txt.Add(strFileName);
/*******************�������������ݲ���*******************************/
		CString str_tmp_2=m_strPath+strFileName;
		if(!stdio_file.Open(str_tmp_2,CFile::modeRead))
		{
			AfxMessageBox("�򿪴���׼���˳���",MB_OK,0);
			return;
		}
		stdio_file.ReadString(str_tmp_2);
//		AfxMessageBox(str_tmp_2);
		message_data=str_tmp_2;
		splitMessage();
		//�ر��ļ�		
		stdio_file.Close();
	}
	FindClose(hFind_txt);
// 	 	for (int p=0;p<data_txt.GetSize();p++)
// 	 	{
// 	 		AfxMessageBox(data_txt.GetAt(p),MB_OK,0);
// 	 	}
}
