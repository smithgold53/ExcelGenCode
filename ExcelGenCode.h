// ExcelGenCode.h : main header file for the ExcelGenCode application
//
#pragma once

#ifndef __AFXWIN_H__
#endif

#include "resource.h"       // main symbols


// CExcelGenCodeApp:
// See ExcelGenCode.cpp for the implementation of this class
//

class CExcelGenCodeApp : public CWinApp
{
public:
	CExcelGenCodeApp();


// Overrides
public:
	virtual BOOL InitInstance();

// Implementation
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

extern CExcelGenCodeApp theApp;