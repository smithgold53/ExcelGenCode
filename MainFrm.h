// MainFrm.h : interface of the CMainFrame class
//

#ifndef MAINFRAME_H
#define MAINFRAME_H
#include "GuiMainFrame.h"
#include "GenCode.h"


class CMainFrame : public CGuiMainFrame
{
	
protected: // create from serialization only
	DECLARE_DYNCREATE(CMainFrame)

// Attributes
public:
	//CTestView m_wndTV;
	//Cabc1 m_wndAA;
	CGenCode m_wndGen;
	CMainFrame();
	virtual ~CMainFrame();
// Operations
protected: 
	void OnInitializes();

// Implementation

// Generated message map functions
protected:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	DECLARE_MESSAGE_MAP()

};
#endif

