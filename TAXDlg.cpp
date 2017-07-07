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
	CFileDialog hFileDlg(true,NULL ,  NULL,   OFN_FILEMUSTEXIST | OFN_READONLY | OFN_PATHMUSTEXIST,  TEXT("EXCEL文件 (*.xlsx)|*.xlsx|所有文件(*.*)|*.*|"),  NULL);
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
		AfxMessageBox(_T("请先选择计算模板"), MB_OKCANCEL | MB_ICONQUESTION);
		return ;
	}
	//读取excel 内容

	Para m_paraPrice;
	Para m_paraAmount;
	m_paraPrice.init();
	m_paraAmount.init();

	if (!m_oExcelApp.CreateDispatch( _T( "Excel.Application" ), NULL ) )
	{
	   ::MessageBox( NULL, _T( "创建Excel服务失败！" ), _T( "错误提示！" ), MB_OK | MB_ICONERROR);
	   exit(1);
	}
 
	//设置为显示
	m_oExcelApp.SetVisible(FALSE);
	m_oWorkBooks.AttachDispatch( m_oExcelApp.GetWorkbooks(), TRUE ); //没有这条语句，下面打开文件返回失败。
 
	LPDISPATCH lpDisp = NULL;
	COleVariant covTrue((short)TRUE);
	COleVariant covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR); 
	Range  oCurCell;
 
	// 打开文件
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
	// 获得活动的WorkBook( 工作簿 )
	m_oWorkBook.AttachDispatch( lpDisp, TRUE );
	// 获得活动的WorkSheet( 工作表 )
	m_oWorkSheet.AttachDispatch( m_oWorkBook.GetActiveSheet(), TRUE );
	// 获得使用的区域Range( 区域 )
	m_oCurrRange.AttachDispatch( m_oWorkSheet.GetUsedRange(), TRUE );
 
	// 获得使用的行数
	long lgUsedRowNum = 0;
	m_oCurrRange.AttachDispatch( m_oCurrRange.GetRows(), TRUE );
	lgUsedRowNum = m_oCurrRange.GetCount();
	// 获得使用的列数
	long lgUsedColumnNum = 0;
	m_oCurrRange.AttachDispatch( m_oCurrRange.GetColumns(), TRUE );
	lgUsedColumnNum = m_oCurrRange.GetCount();
	// 读取Sheet的名称
	CString strSheetName = m_oWorkSheet.GetName();
 
	//得到全部Cells，此时,CurrRange是cells的集合
	m_oCurrRange.AttachDispatch( m_oWorkSheet.GetCells(), TRUE );

	//读取para
	CString m_str;
	for (int i=1;i<=3;i++)	//row ,只读取前3行
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
	
	//读取采购量并计算费用
	int m_nAmount;
	int m_nTotalAmount;
	double m_fFee;

	m_oCurrRange.SetItem( _variant_t( (long)(lgUsedRowNum+1) ), _variant_t( (long)(1) ),COleVariant("费用") );
	
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
		
		//首次计算费用
		if (j==2)
		{
			m_fFee=(float)(m_nAmount*m_paraPrice.m_nArray[1][0]);
			//m_str.Format(_T("%f"),m_fFee);
			//m_oCurrRange.SetItem( _variant_t( (long)(lgUsedRowNum+1) ), _variant_t( (long)(1) ),COleVariant(m_str) );
		}
		else
		{
			//累计采购量找最优价格
			m_paraPrice.GetBestPrice(m_nTotalAmount);	//本次采购最优价格
			//计算表赋值
			Para m_paraFee;		// 价格  量  费用
			m_paraFee.init();
			// 价格赋值
			if (m_paraPrice.m_nBPpos==0)
			{
				m_paraFee.m_nArray[0][0]=m_paraPrice.m_nArray[1][0];
				//最优价格为一个的，后续无需计算
				m_fFee=(float)(m_nAmount*m_paraPrice.m_nArray[1][0]);
			}
			else
			{
				for (int j=0,i=m_paraPrice.m_nBPpos;i>=0;i--)
				{
					m_paraFee.m_nArray[0][j++]=m_paraPrice.m_nArray[1][i];
				}
				
				// 计算价格量
				m_paraFee.m_nArray[1][0]=m_nTotalAmount-m_paraPrice.m_nArray[0][m_paraPrice.m_nBPpos];
				m_paraFee.m_nLenth=1;
				// 实际采购量分级
				int m_ntimes;
				int l=1;
				m_nAmount-=(int)m_paraFee.m_nArray[1][0];
				for (m_ntimes=m_paraPrice.m_nBPpos-1;m_ntimes>=0;m_ntimes--)
				{
					if (m_nAmount>m_paraPrice.m_nArray[2][m_ntimes])	// 实际采购量大于当前分级
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
				// 单项费用计算
				m_fFee=0;
				for (int k=0;k<m_paraFee.m_nLenth;k++)
				{
					m_paraFee.m_nArray[2][k]=m_paraFee.m_nArray[0][k]*m_paraFee.m_nArray[1][k];
					m_fFee+=m_paraFee.m_nArray[2][k];
				}
			}
		}
		//excel  写入
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


	AfxMessageBox(_T("计算完成！数据已写入EXCEL表"), MB_OKCANCEL | MB_ICONQUESTION);
	// 更新列表控件数据
//	m_pExcelOperDlg->initListCtrlColumn( lgUsedColumnNum );
//	m_pExcelOperDlg->updateListCtrlData( arrayStr, lgUsedRowNum );
 
//EXCEL_OUT:
	// 关闭
	m_oWorkBook.Close( covOptional, COleVariant( m_str_path ), covOptional );
	m_oWorkBooks.Close();
	// 释放
	m_oCurrRange.ReleaseDispatch();
	m_oWorkSheet.ReleaseDispatch();
	m_oWorkSheets.ReleaseDispatch();
	m_oWorkBook.ReleaseDispatch();
	m_oWorkBooks.ReleaseDispatch();
	m_oExcelApp.ReleaseDispatch();
	m_oExcelApp.Quit();  // 这条语句是推出Excel程序，任务管理器中的EXCEL进程会自动结束



}
