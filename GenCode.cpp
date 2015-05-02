#include "GenCode.h"
#include "MainFrm.h"
#include "Excel.h"


static void _OnBrowseSelectFnc(CWnd *pWnd){
	CGenCode *pVw = (CGenCode *)pWnd;
	pVw->OnBrowseSelect();
	
} 

/*static void _OnURLChangeFnc(CWnd *pWnd){
	((CGenCode *)pWnd)->OnURLChange();
} */
/*static void _OnURLSetfocusFnc(CWnd *pWnd){
	((CGenCode *)pWnd)->OnURLSetfocus();} */ 
/*static void _OnURLKillfocusFnc(CWnd *pWnd){
	((CGenCode *)pWnd)->OnURLKillfocus();
} */
static int _OnURLCheckValueFnc(CWnd *pWnd){
	return ((CGenCode *)pWnd)->OnURLCheckValue();
} 
static void _OnGenCodeSelectFnc(CWnd *pWnd){
	CGenCode *pVw = (CGenCode *)pWnd;
	pVw->OnGenCodeSelect();
} 
static int _OnAddGenCodeFnc(CWnd *pWnd){
	 return ((CGenCode*)pWnd)->OnAddGenCode();
} 
static int _OnEditGenCodeFnc(CWnd *pWnd){
	 return ((CGenCode*)pWnd)->OnEditGenCode();
} 
static int _OnDeleteGenCodeFnc(CWnd *pWnd){
	 return ((CGenCode*)pWnd)->OnDeleteGenCode();
} 
static int _OnSaveGenCodeFnc(CWnd *pWnd){
	 return ((CGenCode*)pWnd)->OnSaveGenCode();
} 
static int _OnCancelGenCodeFnc(CWnd *pWnd){
	 return ((CGenCode*)pWnd)->OnCancelGenCode();
} 

CGenCode::CGenCode(){

	m_nDlgWidth = 405;
	m_nDlgHeight = 100;
	SetDefaultValues();
}
CGenCode::~CGenCode(){
}

void CGenCode::OnCreateComponents(){
	m_wndExcelGenCode.Create(this, _T("Excel Gen Code"), 5, 5, 395, 60);
	m_wndBrowse.Create(this, _T("Browse"), 305, 30, 390, 55);
	m_wndURL.Create(this,10, 30, 300, 55); 
	m_wndGenCode.Create(this, _T("Gen Code"), 305, 65, 390, 90);
	//m_wndNoteNote.Create(this, _T("Note Note"), 5, 65, 300, 90);

}
void CGenCode::OnInitializeComponents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	m_wndURL.SetLimitText(35);
	m_wndURL.SetCheckValue(true);


}
void CGenCode::OnSetWindowEvents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	m_wndBrowse.SetEvent(WE_CLICK, _OnBrowseSelectFnc);
	//m_wndURL.SetEvent(WE_CHANGE, _OnURLChangeFnc);
	//m_wndURL.SetEvent(WE_SETFOCUS, _OnURLSetfocusFnc);
	//m_wndURL.SetEvent(WE_KILLFOCUS, _OnURLKillfocusFnc);
	m_wndURL.SetEvent(WE_CHECKVALUE, _OnURLCheckValueFnc);
	m_wndGenCode.SetEvent(WE_CLICK, _OnGenCodeSelectFnc);
	AddEvent(1, _T("Add	Ctrl+A"), _OnAddGenCodeFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditGenCodeFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteGenCodeFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSaveGenCodeFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelGenCodeFnc, 0, 'T', VK_CONTROL);
	SetMode(VM_NONE);
	MessageBox(_T("L\x1B0u ý: \x110\x1ED5i Uni\x63o\x64\x65 \x63\x1EE7\x61 \x45x\x63\x65l s\x61ng \x43String tr\x1B0\x1EDB\x63 khi g\x65n :-\x62\x64"));
}
void CGenCode::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndURL.GetDlgCtrlID(), m_szURL);

}
void CGenCode::GetDataToScreen(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CGenCode::GetScreenToData(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();

}
void CGenCode::SetDefaultValues(){

	m_szURL.Empty();

}
int CGenCode::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CMainFrame *pMF = (CMainFrame *) AfxGetMainWnd();
 		CString szSQL;
 		CRecord rs(&pMF->m_db);
  		switch(nMode){
 		case VM_ADD: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 3, 4, -1);
 			SetDefaultValues();
 			break;
 		case VM_EDIT: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 3, 4, -1);
 			break;
 		case VM_VIEW: 
 			EnableControls(FALSE);
 			EnableButtons(FALSE, 3, 4, -1);
 			break;
 		case VM_NONE: 
 			EnableControls(FALSE);
 			EnableButtons(TRUE, 0, 1, -1);
 			SetDefaultValues();
 			break;
 		};
 		UpdateData(FALSE);
 		return nOldMode;
}
void CGenCode::OnBrowseSelect(){
	CMainFrame *pMF = (CMainFrame *) AfxGetMainWnd();
	this->UpdateData();
	CFile f;
	CString strFilter =(_T("Excel Files (*.xls)|*.xls"));
	CFileDialog FileDlg(TRUE, _T(".xls"), NULL, 0, strFilter);
	if( FileDlg.DoModal() == IDOK )
	{
		if( f.Open(FileDlg.GetFileName(), CFile::modeRead) == FALSE )
			return;
		CArchive ar(&f, CArchive::load);
		CString filePath;
		filePath = FileDlg.GetPathName();
		//fileName = FileDlg.GetFileName();
		m_szURL.Format(_T("%s"), filePath);
		ar.Close();
	}
	else
		return;
	f.Close();
	this->UpdateData(FALSE);
	
} 
/*void CGenCode::OnURLChange(){
	
} */
/*void CGenCode::OnURLSetfocus(){
	
} */
/*void CGenCode::OnURLKillfocus(){
	
} */
int CGenCode::OnURLCheckValue(){
	return 0;
} 

struct mergecell{
	int first_row;
	int last_row;
	int first_col;
	int last_col;
};

void CGenCode::OnGenCodeSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	CExcel xls;
	xls.Load(m_szURL);
	CString tmpStr;
	tmpStr.Empty();
	xls.SetWorksheet(0);

	//_msg(_T("%d: %d"), xls.GetTotalColumns(), xls.GetTotalRows());
	for (int i =0; i < xls.GetTotalColumns(); i++){
		int w = xls.GetColWidth(i);
		tmpStr.AppendFormat(_T("\r\txls.SetColumnWidth(%d, %d);"), i, w);
	}
	unsigned short fCol=0, fRow=0, lCol=0, lRow=0;
	CArray<mergecell, mergecell> mergecells;

	for (int i =0; i < xls.GetTotalColumns(); i++){
		for (int j =0; j < xls.GetTotalRows(); j++){
			CString szText = xls.ReadString(j, i); 
			xls.GetMerge(j, i, &fRow, &lRow, &fCol, &lCol);
			
			bool found = false;
			for (int n =0; n < mergecells.GetCount(); n++){
				mergecell mc = mergecells[n];
				if(mc.first_row == fRow && mc.first_col == fCol && mc.last_row == lRow && mc.last_col == lCol)
				{
					 
					found = true;
					break;
				}
			}
			if(!found){
				mergecell mc;
				mc.first_row = fRow;
				mc.first_col = fCol;
				mc.last_row = lRow;
				mc.last_col = lCol;
				mergecells.Add(mc);
			}
			if(!szText.IsEmpty())
			{
				/*CellFormat *cf = xls.GetCellStyle(j, i);
				int border, cellstyle, fontsize;
				bool bold, italic;
				border = cf->GetBorder();
				cellstyle = cf->GetCellStyle();
				fontsize = cf->GetFontSize();
				bold = cf->IsBold();

				italic = cf->IsItalic();*/
				CellFormat *cf = xls.GetCellStyle(j, i);
				int fontsize = cf->GetFontSize();
				tmpStr.AppendFormat(_T("\r\txls.SetCellText(%d, %d, _T(\"%s\"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, %d);"), i, j, szText, fontsize);
			}		
		} //for j
	} //for i

	for (int i =0; i < mergecells.GetCount(); i++){
		mergecell mc = mergecells[i];
		//tmpStr.AppendFormat(_T("\r\txls.MergeCell(%d, %d, %d, %d);"), mc.first_col, mc.first_row, mc.last_col, mc.last_row);
		tmpStr.AppendFormat(_T("\r\txls.SetMerge(%d, %d, %d, %d);"),mc.first_row, mc.last_row, mc.first_col, mc.last_col);
	}
	//_msg(_T("%s"), tmpStr);

	//Save to file text
	this->UpdateData();
	CFile f;
	CString strFilter = (_T("Text Files (*.txt)|*.txt"));
	CFileDialog FileD(FALSE, _T(".txt"), NULL, 0, strFilter);
	if( FileD.DoModal() == IDOK )
	{
	f.Open(FileD.GetFileName(), CFile::modeCreate | CFile::modeWrite);
		CArchive ar(&f, CArchive::store);
		ar << tmpStr;
		ar.Close();
	}
	else
		return;
	f.Close();
	
	
} 
int CGenCode::OnAddGenCode(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CMainFrame *pMF = (CMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CGenCode::OnEditGenCode(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CMainFrame *pMF = (CMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CGenCode::OnDeleteGenCode(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CMainFrame *pMF = (CMainFrame *)AfxGetMainWnd();
 	CString szSQL;
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO)
 		return -1;
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") );
 	int ret = pMF->ExecSQL(szSQL);
 	if(ret >= 0){
 		SetMode(VM_NONE);
 		OnCancelGenCode();
 	}
	return 0;
}
int CGenCode::OnSaveGenCode(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT)
 		return -1;
 	if(!IsValidateData())
 		return -1;
 	CMainFrame *pMF = (CMainFrame *)AfxGetMainWnd();
 	CString szSQL;
 	if(GetMode() == VM_ADD){
 		//szSQL = m_tblTbl.GetInsertSQL();
 	}
 	else if(GetMode() == VM_EDIT){
 		CString szWhere;
 		szWhere.Format(_T(" WHERE"));
 		//szSQL = m_tblTbl.GetUpdateSQL(_T("createdby,createddate"));
 		szSQL += szWhere;
 	}
 _fmsg(_T("%s"), szSQL);
 	int ret = pMF->ExecSQL(szSQL);
 	if(ret > 0)
 	{
 		//OnGenCodeListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CGenCode::OnCancelGenCode(){
 	if(GetMode() == VM_EDIT)
 	{
 		SetMode(VM_VIEW);
 	} 
 	else 
 	{
 		SetMode(VM_NONE);
 	} 
 	CMainFrame *pMF = (CMainFrame *)AfxGetMainWnd();
 	return 0;
} 
int CGenCode::OnGenCodeListLoadData(){
	return 0;
}

