// TAX.h : main header file for the TAX application
//

#if !defined(AFX_TAX_H__612B1EC9_6937_40F1_A76B_8A2624B55647__INCLUDED_)
#define AFX_TAX_H__612B1EC9_6937_40F1_A76B_8A2624B55647__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CTAXApp:
// See TAX.cpp for the implementation of this class
//

class CTAXApp : public CWinApp
{
public:
	CTAXApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTAXApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CTAXApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TAX_H__612B1EC9_6937_40F1_A76B_8A2624B55647__INCLUDED_)
