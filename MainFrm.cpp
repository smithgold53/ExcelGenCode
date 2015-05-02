// MainFrm.cpp : implementation of the CMainFrame class
//

#include "stdafx.h"
#include "MainFrm.h"
#include "GuiUtils.h"
#include "resource.h"
#include ".\mainfrm.h"


// CMainFrame

IMPLEMENT_DYNCREATE(CMainFrame, CGuiMainFrame)

BEGIN_MESSAGE_MAP(CMainFrame, CGuiMainFrame)
	ON_WM_CREATE()
END_MESSAGE_MAP()




// CMainFrame construction/destruction

CMainFrame::CMainFrame()
{
	//SetLocalLang(1);
	SetEnableLogin(false);
	m_szModuleID = _T("EGC");
	m_szSize = CSize(400, 150);
	m_szVersion = _T("1.0");
}

CMainFrame::~CMainFrame()
{
}


int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CGuiMainFrame::OnCreate(lpCreateStruct) == -1)
		return -1;
	ModifyStyle(WS_MAXIMIZEBOX|WS_SIZEBOX|WS_MINIMIZEBOX, 0);
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	CenterWindow();
	//m_wndTV.Create(this);
	m_wndGen.Create(this);

	return 0;
}
#include "Excel.h"
void CMainFrame::OnInitializes()
{

}
