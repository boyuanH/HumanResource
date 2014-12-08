// StaffDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "HumanResouce.h"
#include "StaffDlg.h"
#include "afxdialogex.h"
#include "ADOConnection.h"

// StaffDlg 对话框

IMPLEMENT_DYNAMIC(StaffDlg, CDialogEx)

StaffDlg::StaffDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(StaffDlg::IDD, pParent)
	, m_findBusinessEntityID("")
{

}

StaffDlg::~StaffDlg()
{
}

void StaffDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT1, m_findBusinessEntityID);
	DDX_Control(pDX, IDC_LIST1, m_List);
}


BEGIN_MESSAGE_MAP(StaffDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON4, &StaffDlg::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON1, &StaffDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// StaffDlg 消息处理程序


BOOL StaffDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	DWORD dwExStyle=LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES|LVS_EX_HEADERDRAGDROP|LVS_EX_ONECLICKACTIVATE;
	m_List.ModifyStyle(0,LVS_REPORT|LVS_SINGLESEL|LVS_SHOWSELALWAYS);
	m_List.SetExtendedStyle(dwExStyle);
	m_List.InsertColumn(0,_T("BusinessEntityID"),LVCFMT_CENTER,50,0);
	m_List.InsertColumn(1,_T("NationalIDNumber"),LVCFMT_CENTER,100,0);
	m_List.InsertColumn(2,_T("JobTitle"),	LVCFMT_CENTER,200,0);	
	m_List.InsertColumn(3,_T("MaritalStatus"),LVCFMT_CENTER,100,0);
	m_List.InsertColumn(4,_T("Gender"),	LVCFMT_CENTER,60,0);
	m_List.InsertColumn(5,_T("HireDate"),	LVCFMT_CENTER,100,0);
	m_List.InsertColumn(6,_T("SalariedFlag"),LVCFMT_CENTER,100,0);

	CADOConnection adoConn;
	_bstr_t vSQL="select [BusinessEntityID],[NationalIDNumber],[JobTitle],[MaritalStatus],[Gender],[HireDate],[SalariedFlag] from [AdventureWorks2008].[HumanResources].[Employee]";
	_RecordsetPtr m_pRecordset;
	adoConn.OnInitAdo();
	m_pRecordset=adoConn.GetRecordset(vSQL);
	if(m_pRecordset->adoEOF)
		return FALSE;
	m_List.DeleteAllItems();
	m_pRecordset->MoveFirst();
	CString businessEntityID,nationalIDNumber,jobTitle,maritalStatus,gender,hireDate,salariedFlag;
	int nPos = 0;

	while (!m_pRecordset->adoEOF)
	{
		businessEntityID = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("BusinessEntityID");
		nationalIDNumber = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("NationalIDNumber");
		jobTitle = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("JobTitle");
		maritalStatus = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("MaritalStatus");
		gender = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("Gender");
		hireDate = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("HireDate");
		salariedFlag = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("SalariedFlag");

		m_List.InsertItem(nPos,businessEntityID);
		m_List.SetItemText(nPos,1,nationalIDNumber);
		m_List.SetItemText(nPos,2,jobTitle);
		m_List.SetItemText(nPos,3,maritalStatus);
		m_List.SetItemText(nPos,4,gender);
		m_List.SetItemText(nPos,5,hireDate);
		m_List.SetItemText(nPos,6,salariedFlag);

		m_pRecordset->MoveNext();
		nPos++;
	}
	adoConn.ExitConnect();
	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}

void StaffDlg::OnBnClickedButton4()
{
	//查找
	UpdateData(true);
	if (m_findBusinessEntityID == "")
	{
		AfxMessageBox(_T("请输入要查询的BusinessEntityID"));
		return;
	}	

	CADOConnection adoConn;
	_bstr_t vSQL="select [BusinessEntityID],[NationalIDNumber],[JobTitle],[MaritalStatus],[Gender],[HireDate],[SalariedFlag] from [AdventureWorks2008].[HumanResources].[Employee] where BusinessEntityID = "+(_bstr_t)m_findBusinessEntityID;
	_RecordsetPtr m_pRecordset;
	adoConn.OnInitAdo();
	m_pRecordset=adoConn.GetRecordset(vSQL);
	if(m_pRecordset->adoEOF)
	{
		AfxMessageBox(_T("查无此人!"));
		return ;
	}
	m_List.DeleteAllItems();
	CString businessEntityID,nationalIDNumber,jobTitle,maritalStatus,gender,hireDate,salariedFlag;

	businessEntityID = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("BusinessEntityID");
	nationalIDNumber = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("NationalIDNumber");
	jobTitle = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("JobTitle");
	maritalStatus = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("MaritalStatus");
	gender = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("Gender");
	hireDate = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("HireDate");
	salariedFlag = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("SalariedFlag");
	
	m_List.InsertItem(0,businessEntityID);
	m_List.SetItemText(0,1,nationalIDNumber);
	m_List.SetItemText(0,2,jobTitle);
	m_List.SetItemText(0,3,maritalStatus);
	m_List.SetItemText(0,4,gender);
	m_List.SetItemText(0,5,hireDate);
	m_List.SetItemText(0,6,salariedFlag);

	adoConn.ExitConnect();
}

void StaffDlg::OnBnClickedButton1()
{
	// 删除
	CString str; 
	int i = -1;
	BOOL isSelected = FALSE;
	for(i=0; i<m_List.GetItemCount(); i++) 
	{ 
		if( m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED ) 
		{ 
			str.Format(_T(" 选中了第%d行"), i+1); 
		//	AfxMessageBox(str); 
			isSelected = TRUE;
			break;
		} 
	} 
	if (!isSelected)
	{
		AfxMessageBox(_T("没选择"));
		return;
	}

	TCHAR delID[1024]; 
	LVITEM lvi; 
	lvi.iItem = i; //行
	lvi.iSubItem = 0; //列
	lvi.mask = LVIF_TEXT; 
	lvi.pszText = delID; 
	lvi.cchTextMax = 1024; 
	m_List.GetItem(&lvi); 
	
	//数据库删除,从Item获得的ID
	CADOConnection m_ADOConn;
	m_ADOConn.OnInitAdo(); 
	_RecordsetPtr m_pRecordset;
	_bstr_t empSQL="select [BusinessEntityID] from [AdventureWorks2008].[HumanResources].[Employee] where [BusinessEntityID] = "+(_bstr_t)delID;
	m_pRecordset=m_ADOConn.GetRecordset(empSQL);
	if(m_pRecordset->adoEOF)
	{
		AfxMessageBox(_T("删除错误，查无此人!"));
		return ;
	}
	m_ADOConn.ExitConnect();
//	m_ADOConn.OnInitAdo(); 
	_bstr_t delSQL1="delete from [AdventureWorks2008].[HumanResources].[Employee] where [BusinessEntityID] = "+(_bstr_t)delID;
	_bstr_t delSQL2="delete from [AdventureWorks2008].[HumanResources].[EmployeeDepartmentHistory] where [BusinessEntityID] = "+(_bstr_t)delID;
	_bstr_t delSQL3="delete from [AdventureWorks2008].[HumanResources].[JobCandidate] where [BusinessEntityID] = "+(_bstr_t)delID;
	_bstr_t delSQL4="delete from [AdventureWorks2008].[HumanResources].[EmployeePayHistory] where [BusinessEntityID] = "+(_bstr_t)delID;

	m_ADOConn.ExecuteSQL(delSQL1);
	m_ADOConn.ExecuteSQL(delSQL2);
	m_ADOConn.ExecuteSQL(delSQL3);
	m_ADOConn.ExecuteSQL(delSQL4);

	m_ADOConn.ExitConnect();
	//重新显示全部的list
	_bstr_t vSQL="select [BusinessEntityID],[NationalIDNumber],[JobTitle],[MaritalStatus],[Gender],[HireDate],[SalariedFlag] from [AdventureWorks2008].[HumanResources].[Employee] where BusinessEntityID = "+(_bstr_t)m_findBusinessEntityID;
	m_pRecordset=m_ADOConn.GetRecordset(vSQL);
	m_List.DeleteAllItems();
	if(m_pRecordset->adoEOF)
		return ;	
	m_pRecordset->MoveFirst();
	CString businessEntityID,nationalIDNumber,jobTitle,maritalStatus,gender,hireDate,salariedFlag;
	int nPos = 0;

	while (!m_pRecordset->adoEOF)
	{
		businessEntityID = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("BusinessEntityID");
		nationalIDNumber = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("NationalIDNumber");
		jobTitle = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("JobTitle");
		maritalStatus = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("MaritalStatus");
		gender = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("Gender");
		hireDate = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("HireDate");
		salariedFlag = (LPCTSTR)(_bstr_t)m_pRecordset->GetCollect("SalariedFlag");

		m_List.InsertItem(nPos,businessEntityID);
		m_List.SetItemText(nPos,1,nationalIDNumber);
		m_List.SetItemText(nPos,2,jobTitle);
		m_List.SetItemText(nPos,3,maritalStatus);
		m_List.SetItemText(nPos,4,gender);
		m_List.SetItemText(nPos,5,hireDate);
		m_List.SetItemText(nPos,6,salariedFlag);

		m_pRecordset->MoveNext();
		nPos++;
	}
	m_ADOConn.ExitConnect();
	m_List.SetItemState(i, 0, LVIS_SELECTED|LVIS_FOCUSED); 
}
