// TAXDlg.cpp : implementation file
//

#include "stdafx.h"
#include "TAX.h"
#include "TAXDlg.h"
#include "comutil.h"
#include "comdef.h"
#include "para.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

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
// CTAXDlg dialog

CTAXDlg::CTAXDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTAXDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTAXDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CTAXDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTAXDlg)
	DDX_Control(pDX, IDC_EDIT_PATH, m_edit_path);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CTAXDlg, CDialog)
	//{{AFX_MSG_MAP(CTAXDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_SELECT, OnButtonSelect)
	ON_BN_CLICKED(IDC_BUTTON_CAL, OnButtonCal)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTAXDlg message handlers

BOOL CTAXDlg::OnInitDialog()
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

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CTAXDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CTAXDlg::OnPaint() 
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
HCURSOR CTAXDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CTAXDlg::OnButtonSelect() 
{
	// TODO: Add your control notification handler code here
	CFileDialog hFileDlg(true,NULL ,  NULL,   OFN_FILEMUSTEXIST | OFN_READONLY | OFN_PATHMUSTEXIST,  TEXT("EXCEL�ļ� (*.xlsx)|*.xlsx|�����ļ�(*.*)|*.*|"),  NULL);
	if(hFileDlg.DoModal() == IDOK)
	{
		m_str_path = hFileDlg.GetPathName();
		UpdateData(FALSE);
	}
	m_edit_path.SetWindowText(m_str_path);
}

void CTAXDlg::OnButtonCal() 
{
	// TODO: Add your control notification handler code here
	if (m_str_path.IsEmpty())
	{
		AfxMessageBox(_T("����ѡ�����ģ��"), MB_OKCANCEL | MB_ICONQUESTION);
		return ;
	}
	//��ȡexcel ����

	Para m_paraPrice;
	Para m_paraAmount;
	m_paraPrice.init();
	m_paraAmount.init();

	if (!m_oExcelApp.CreateDispatch( _T( "Excel.Application" ), NULL ) )
	{
	   ::MessageBox( NULL, _T( "����Excel����ʧ�ܣ�" ), _T( "������ʾ��" ), MB_OK | MB_ICONERROR);
	   exit(1);
	}
 
	//����Ϊ��ʾ
	m_oExcelApp.SetVisible(FALSE);
	m_oWorkBooks.AttachDispatch( m_oExcelApp.GetWorkbooks(), TRUE ); //û��������䣬������ļ�����ʧ�ܡ�
 
	LPDISPATCH lpDisp = NULL;
	COleVariant covTrue((short)TRUE);
	COleVariant covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR); 
	Range  oCurCell;
 
	// ���ļ�
	lpDisp = m_oWorkBooks.Open( m_str_path,_variant_t(vtMissing),_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing),
		_variant_t(vtMissing) );
	// ��û��WorkBook( ������ )
	m_oWorkBook.AttachDispatch( lpDisp, TRUE );
	// ��û��WorkSheet( ������ )
	m_oWorkSheet.AttachDispatch( m_oWorkBook.GetActiveSheet(), TRUE );
	// ���ʹ�õ�����Range( ���� )
	m_oCurrRange.AttachDispatch( m_oWorkSheet.GetUsedRange(), TRUE );
 
	// ���ʹ�õ�����
	long lgUsedRowNum = 0;
	m_oCurrRange.AttachDispatch( m_oCurrRange.GetRows(), TRUE );
	lgUsedRowNum = m_oCurrRange.GetCount();
	// ���ʹ�õ�����
	long lgUsedColumnNum = 0;
	m_oCurrRange.AttachDispatch( m_oCurrRange.GetColumns(), TRUE );
	lgUsedColumnNum = m_oCurrRange.GetCount();
	// ��ȡSheet������
	CString strSheetName = m_oWorkSheet.GetName();
 
	//�õ�ȫ��Cells����ʱ,CurrRange��cells�ļ���
	m_oCurrRange.AttachDispatch( m_oWorkSheet.GetCells(), TRUE );

	//��ȡpara
	CString m_str;
	for (int i=1;i<=3;i++)	//row ,ֻ��ȡǰ3��
	{
		for (int j=2;j<=lgUsedColumnNum;j++)
		{
			oCurCell.AttachDispatch( m_oCurrRange.GetItem( COleVariant( (long)(i)), COleVariant( (long)j) ).pdispVal, TRUE );
			VARIANT varItemName = oCurCell.GetText();
			m_str=varItemName.bstrVal;
			if (m_str.IsEmpty())
			{
				m_paraPrice.m_nLenth=j-2;
				break;
			}
			m_paraPrice.m_nArray[i-1][j-2]=atof(m_str);
		}
	}
	
	//��ȡ�ɹ������������
	int m_nAmount;
	int m_nTotalAmount;
	double m_fFee;

	m_oCurrRange.SetItem( _variant_t( (long)(lgUsedRowNum+1) ), _variant_t( (long)(1) ),COleVariant("����") );
	
	for (int j=2;j<=lgUsedColumnNum;j++)
	{
			
		oCurCell.AttachDispatch( m_oCurrRange.GetItem( COleVariant( (long)(4)), COleVariant( (long)j) ).pdispVal, TRUE );
		VARIANT varItemName = oCurCell.GetText();
		m_str=varItemName.bstrVal;
		if (m_str.IsEmpty())
		{
			break;
		}
		m_nAmount=atoi(m_str);
		
		oCurCell.AttachDispatch( m_oCurrRange.GetItem( COleVariant( (long)(5)), COleVariant( (long)j) ).pdispVal, TRUE );
		varItemName = oCurCell.GetText();
		m_str=varItemName.bstrVal;
		if (m_str.IsEmpty())
		{
			break;
		}

		m_nTotalAmount=atoi(m_str);
		
		//�״μ������
		if (j==2)
		{
			m_fFee=(float)(m_nAmount*m_paraPrice.m_nArray[1][0]);
			//m_str.Format(_T("%f"),m_fFee);
			//m_oCurrRange.SetItem( _variant_t( (long)(lgUsedRowNum+1) ), _variant_t( (long)(1) ),COleVariant(m_str) );
		}
		else
		{
			//�ۼƲɹ��������ż۸�
			m_paraPrice.GetBestPrice(m_nTotalAmount);	//���βɹ����ż۸�
			//�����ֵ
			Para m_paraFee;		// �۸�  ��  ����
			m_paraFee.init();
			// �۸�ֵ
			if (m_paraPrice.m_nBPpos==0)
			{
				m_paraFee.m_nArray[0][0]=m_paraPrice.m_nArray[1][0];
				//���ż۸�Ϊһ���ģ������������
				m_fFee=(float)(m_nAmount*m_paraPrice.m_nArray[1][0]);
			}
			else
			{
				for (int j=0,i=m_paraPrice.m_nBPpos;i>=0;i--)
				{
					m_paraFee.m_nArray[0][j++]=m_paraPrice.m_nArray[1][i];
				}
				
				// ����۸���
				m_paraFee.m_nArray[1][0]=m_nTotalAmount-m_paraPrice.m_nArray[0][m_paraPrice.m_nBPpos];
				m_paraFee.m_nLenth=1;
				// ʵ�ʲɹ����ּ�
				int m_ntimes;
				int l=1;
				m_nAmount-=(int)m_paraFee.m_nArray[1][0];
				for (m_ntimes=m_paraPrice.m_nBPpos-1;m_ntimes>=0;m_ntimes--)
				{
					if (m_nAmount>m_paraPrice.m_nArray[2][m_ntimes])	// ʵ�ʲɹ������ڵ�ǰ�ּ�
					{
						m_nAmount-=(int)m_paraPrice.m_nArray[2][m_ntimes];
						m_paraFee.m_nArray[1][l]=m_paraPrice.m_nArray[2][m_ntimes];
						m_paraFee.m_nLenth++;
					}
					else
					{
						m_paraFee.m_nArray[1][l]=m_nAmount;
						m_paraFee.m_nLenth++;
						break;
					}
					l++;
				}
				// ������ü���
				m_fFee=0;
				for (int k=0;k<m_paraFee.m_nLenth;k++)
				{
					m_paraFee.m_nArray[2][k]=m_paraFee.m_nArray[0][k]*m_paraFee.m_nArray[1][k];
					m_fFee+=m_paraFee.m_nArray[2][k];
				}
			}
		}
		//excel  д��
		m_str.Format(_T("%f"),m_fFee);
		m_oCurrRange.SetItem( _variant_t( (long)(lgUsedRowNum+1) ), _variant_t( (long)(j) ),COleVariant(m_str) );
	}
	

	m_oWorkBook.SaveAs( COleVariant( m_str_path ),
						_variant_t(vtMissing),
						_variant_t(vtMissing),
						_variant_t(vtMissing),
						_variant_t(vtMissing),
						_variant_t(vtMissing),
						0,
						_variant_t(vtMissing),
						_variant_t(vtMissing),
						_variant_t(vtMissing),
						_variant_t(vtMissing),
						_variant_t(vtMissing) );


	AfxMessageBox(_T("������ɣ�������д��EXCEL��"), MB_OKCANCEL | MB_ICONQUESTION);
	// �����б�ؼ�����
//	m_pExcelOperDlg->initListCtrlColumn( lgUsedColumnNum );
//	m_pExcelOperDlg->updateListCtrlData( arrayStr, lgUsedRowNum );
 
//EXCEL_OUT:
	// �ر�
	m_oWorkBook.Close( covOptional, COleVariant( m_str_path ), covOptional );
	m_oWorkBooks.Close();
	// �ͷ�
	m_oCurrRange.ReleaseDispatch();
	m_oWorkSheet.ReleaseDispatch();
	m_oWorkSheets.ReleaseDispatch();
	m_oWorkBook.ReleaseDispatch();
	m_oWorkBooks.ReleaseDispatch();
	m_oExcelApp.ReleaseDispatch();
	m_oExcelApp.Quit();  // ����������Ƴ�Excel��������������е�EXCEL���̻��Զ�����



}
