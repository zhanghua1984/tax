// TAXDlg.h : header file
//

#if !defined(AFX_TAXDLG_H__9675FB14_F613_49D2_8006_658E1EB035B7__INCLUDED_)
#define AFX_TAXDLG_H__9675FB14_F613_49D2_8006_658E1EB035B7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "excel.h"
/////////////////////////////////////////////////////////////////////////////
// CTAXDlg dialog

class CTAXDlg : public CDialog
{
// Construction
public:
	CTAXDlg(CWnd* pParent = NULL);	// standard constructor
	
	CString m_str_path;
	// Excel应用对象
	_Application  m_oExcelApp;   // Excel程序
	_Worksheet  m_oWorkSheet;   // 工作表
	_Workbook  m_oWorkBook;   // 工作簿
	Workbooks  m_oWorkBooks;  // 工作簿集合
	Worksheets m_oWorkSheets;  // 工作表集合
	Range m_oCurrRange;   // 使用区域

// Dialog Data
	//{{AFX_DATA(CTAXDlg)
	enum { IDD = IDD_TAX_DIALOG };
	CEdit	m_edit_path;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTAXDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CTAXDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnButtonSelect();
	afx_msg void OnButtonCal();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TAXDLG_H__9675FB14_F613_49D2_8006_658E1EB035B7__INCLUDED_)
