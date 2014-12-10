// StaffDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "HumanResouce.h"
#include "StaffDlg.h"
#include "afxdialogex.h"
#include "ADOConnection.h"

// StaffDlg �Ի���


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
	ON_NOTIFY(NM_CLICK, IDC_LIST1, &StaffDlg::OnNMClickList1)
	ON_BN_CLICKED(IDC_BUTTON2, &StaffDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, &StaffDlg::OnBnClickedButton3)
END_MESSAGE_MAP()


// StaffDlg ��Ϣ�������


BOOL StaffDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
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
	listUpdate();
	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}

void StaffDlg::OnBnClickedButton4()
{
	//����
	UpdateData(true);
	if (m_findBusinessEntityID == "")
	{
		AfxMessageBox(_T("������Ҫ��ѯ��BusinessEntityID"));
		return;
	}	

	CADOConnection adoConn;
	_bstr_t vSQL="select [BusinessEntityID],[NationalIDNumber],[JobTitle],[MaritalStatus],[Gender],[HireDate],[SalariedFlag] from [AdventureWorks2008].[HumanResources].[Employee] where BusinessEntityID = "+(_bstr_t)m_findBusinessEntityID;
	_RecordsetPtr m_pRecordset;
	adoConn.OnInitAdo();
	m_pRecordset=adoConn.GetRecordset(vSQL);
	if(m_pRecordset->adoEOF)
	{
		AfxMessageBox(_T("���޴���!"));
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
	// ɾ��
	CString str; 
	int i = -1;
	BOOL isSelected = FALSE;
	for(i=0; i<m_List.GetItemCount(); i++) 
	{ 
		if( m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED ) 
		{ 
			str.Format(_T(" ѡ���˵�%d��"), i+1); 
		//	AfxMessageBox(str); 
			isSelected = TRUE;
			break;
		} 
	} 
	if (!isSelected)
	{
		AfxMessageBox(_T("ûѡ��"));
		return;
	}

	TCHAR delID[1024]; 
	LVITEM lvi; 
	lvi.iItem = i; //��
	lvi.iSubItem = 0; //��
	lvi.mask = LVIF_TEXT; 
	lvi.pszText = delID; 
	lvi.cchTextMax = 1024; 
	m_List.GetItem(&lvi); 
	
	//���ݿ�ɾ��,��Item��õ�ID
	CADOConnection m_ADOConn;
	m_ADOConn.OnInitAdo(); 
 	_RecordsetPtr m_pRecordset;
	
	CString delStr;
	delStr.Format(_T("%s"),delID);
	
	_bstr_t delSQL1="delete from [AdventureWorks2008].[HumanResources].[JobCandidate] where [BusinessEntityID] = "+(_bstr_t)delStr;
	_bstr_t delSQL2="delete from [AdventureWorks2008].[HumanResources].[EmployeePayHistory] where [BusinessEntityID] = "+(_bstr_t)delStr;
	_bstr_t delSQL3="delete from [AdventureWorks2008].[HumanResources].[EmployeeDepartmentHistory] where [BusinessEntityID] = "+(_bstr_t)delStr;
	_bstr_t delSQL4="delete from [AdventureWorks2008].[HumanResources].[Employee] where [BusinessEntityID] = "+(_bstr_t)delStr;
	m_ADOConn.ExecuteSQL(delSQL1);
	m_ADOConn.ExecuteSQL(delSQL2);
	m_ADOConn.ExecuteSQL(delSQL3);
	m_ADOConn.ExecuteSQL(delSQL4);

	//������ʾȫ����list
	listUpdate();
	m_List.SetItemState(i, 0, LVIS_SELECTED|LVIS_FOCUSED); 
}


void StaffDlg::OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	for (int j = 0;j < 7; j++)
	{
		GetDlgItem(1011+j)->SetWindowText(_T(""));
	}

	int i = -1;
	BOOL isSelected = FALSE;
	for(i=0; i<m_List.GetItemCount(); i++) 
	{ 
		if( m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED ) 
		{ 
			isSelected = TRUE;
			break;
		} 
	} 
	
	if (isSelected)
	{
		for (int j = 0;j<7;j++)
		{
			TCHAR sztChar[1024]; 
			LVITEM lvi; 
			lvi.iItem = i; //��
			lvi.iSubItem = j; //��
			lvi.mask = LVIF_TEXT; 
			lvi.pszText = sztChar; 
			lvi.cchTextMax = 1024; 
			m_List.GetItem(&lvi); 
			int temp = j+2;
			GetDlgItem(1011+j)->SetWindowText(sztChar);
		}
	}

	*pResult = 0;
}


void StaffDlg::OnBnClickedButton2() 
{
	//���
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString businessEntityID,nationalIDNumber,jobTitle,maritalStatus,gender,hireDate,salariedFlag;
	GetDlgItem(1011)->GetWindowText(businessEntityID);
	GetDlgItem(1012)->GetWindowText(nationalIDNumber);
	GetDlgItem(1013)->GetWindowText(jobTitle);
	GetDlgItem(1014)->GetWindowText(maritalStatus);
	GetDlgItem(1015)->GetWindowText(gender);
	GetDlgItem(1016)->GetWindowText(hireDate);
	GetDlgItem(1017)->GetWindowText(salariedFlag);
	if (((businessEntityID.IsEmpty())&&(nationalIDNumber.IsEmpty())&&(jobTitle.IsEmpty())&&(maritalStatus.IsEmpty())&&(gender.IsEmpty())&&(hireDate.IsEmpty())&&(salariedFlag.IsEmpty())))
	{
		MessageBox(_T("������������"));
		return;
	}

	if (ruleCheck(2) != 0)
	{
		return;
	}

	CADOConnection adoConn;
	if (salariedFlag == _T("-1"))
	{
		salariedFlag = _T("1");
	}
	_bstr_t vSQL="insert into [AdventureWorks2008].[HumanResources].[Employee]([BusinessEntityID],[NationalIDNumber],[LoginID],[JobTitle],[MaritalStatus],[Gender],[HireDate],[SalariedFlag]) values ('"+ (_bstr_t)businessEntityID + "','"+(_bstr_t)nationalIDNumber+ "','"+(_bstr_t)(jobTitle+nationalIDNumber)+ "','"+(_bstr_t)jobTitle+ "','"+(_bstr_t)maritalStatus+ "','"+(_bstr_t)gender+ "','"+(_bstr_t)hireDate+ "','"+(_bstr_t)salariedFlag+ "')";
	_RecordsetPtr m_pRecordset;
	adoConn.OnInitAdo();
	adoConn.GetRecordset(vSQL);
	adoConn.m_pConnection->Close();
	listUpdate();
}


void StaffDlg::OnBnClickedButton3()
{
	//�޸�
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	//���List�е�ID��Edit�е�ID�Ƿ�һ�£�ID�������޸�
	UpdateData(TRUE);
	int i = -1;
	BOOL isSelected = FALSE;
	for(i=0; i<m_List.GetItemCount(); i++) 
	{ 
		if( m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED ) 
		{ 
			isSelected = TRUE;
			break;
		} 
	} 

	if (!isSelected)
	{
		MessageBox(_T("ѡ����Ҫ�޸ĵ���"));
		return;
	}

	TCHAR sztChar[1024]; 
	LVITEM lvi; 
	lvi.iItem = i; //��
	lvi.iSubItem = 0; //��
	lvi.mask = LVIF_TEXT; 
	lvi.pszText = sztChar; 
	lvi.cchTextMax = 1024; 
	m_List.GetItem(&lvi); 
	CString idLStr;
	idLStr.Format(_T("%s"),sztChar);
	
	CString idStr;
	GetDlgItem(IDC_EDIT2)->GetWindowTextW(idStr);

	if (idLStr != idStr)
	{
		MessageBox(_T("BusinessEntityID���������"));
		return;
	}

	CString businessEntityID,nationalIDNumber,jobTitle,maritalStatus,gender,hireDate,salariedFlag;
	GetDlgItem(1011)->GetWindowText(businessEntityID);
	GetDlgItem(1012)->GetWindowText(nationalIDNumber);
	GetDlgItem(1013)->GetWindowText(jobTitle);
	GetDlgItem(1014)->GetWindowText(maritalStatus);
	GetDlgItem(1015)->GetWindowText(gender);
	GetDlgItem(1016)->GetWindowText(hireDate);
	GetDlgItem(1017)->GetWindowText(salariedFlag);

	if (((businessEntityID.IsEmpty())&&(nationalIDNumber.IsEmpty())&&(jobTitle.IsEmpty())&&(maritalStatus.IsEmpty())&&(gender.IsEmpty())&&(hireDate.IsEmpty())&&(salariedFlag.IsEmpty())))
	{
		MessageBox(_T("������������"));
		return;
	}

	if (ruleCheck(1) != 0)
	{
		return;
	}

	CADOConnection adoConn;
	if (salariedFlag == _T("-1"))
	{
		salariedFlag = _T("1");
	}
	_bstr_t vSQL= "UPDATE [AdventureWorks2008].[HumanResources].[Employee] SET [NationalIDNumber] = '"+(_bstr_t)nationalIDNumber +"',JobTitle = '"+(_bstr_t)jobTitle +"',MaritalStatus = '"+(_bstr_t)maritalStatus +"',Gender = '"+(_bstr_t)gender +"',HireDate = '"+(_bstr_t)hireDate +"',SalariedFlag = '"+(_bstr_t)salariedFlag +"' WHERE [BusinessEntityID] = '"+(_bstr_t)businessEntityID +"'";
	_RecordsetPtr m_pRecordset;
	adoConn.OnInitAdo();
	adoConn.GetRecordset(vSQL);
	listUpdate();
	adoConn.m_pConnection->Close();
}

int StaffDlg::ruleCheck(int mode)
{
	//mode = 1 �������Ϸ��ԣ�mode = 2 ���Ϸ�����Ψһ��
	//���ڼ��editBox�в����ĺϷ����Լ����ֲ�����Ψһ��
	UpdateData(TRUE);	
	CString businessEntityID,nationalIDNumber,jobTitle,maritalStatus,gender,hireDate,salariedFlag;
	GetDlgItem(1011)->GetWindowText(businessEntityID);
	GetDlgItem(1012)->GetWindowText(nationalIDNumber);
	GetDlgItem(1013)->GetWindowText(jobTitle);
	GetDlgItem(1014)->GetWindowText(maritalStatus);
	GetDlgItem(1015)->GetWindowText(gender);
	GetDlgItem(1016)->GetWindowText(hireDate);
	GetDlgItem(1017)->GetWindowText(salariedFlag);
	// ���������ݵĺϷ��� ���ϸ񷵻�ֵΪ 1
	if (!(businessEntityID.SpanIncluding(_T("0123456789")) == businessEntityID))
	{
		MessageBox(_T("businessEntityID ��ʽ����"));
		return 1;
	}
	if (!(nationalIDNumber.SpanIncluding(_T("0123456789")) == nationalIDNumber))
	{
		MessageBox(_T("nationalIDNumber ��ʽ����"));
		return 1;
	}
	if ( !((maritalStatus == _T("M") || maritalStatus == _T("S"))))
	{
		MessageBox(_T("maritalStatus ��������"));
		return 1;
	}

	if ( !((gender == _T("M") || gender == _T("F"))))
	{
		MessageBox(_T("gender ��������"));
		return 1;
	}

	if ( !((salariedFlag == _T("0") || salariedFlag == _T("-1"))))
	{
		MessageBox(_T("salariedFlag ��������"));
		return 1;
	}

	if (jobTitle.GetLength()>50 ||jobTitle.GetLength()<0)
	{
		MessageBox(_T("jobTitle ���ݴ���"));
		return 1;
	}

	//hireDate��ʽ��� ��׼��ʽXXXX-XX-XX������XXΪ���� 1999-01-30
	if (hireDate.GetLength() != 10)
	{
		MessageBox(_T("hireDate ����"));
		return 1;
	}

	CString temp = hireDate.Left(4);
	if ((_ttoi(temp)-2050 > 0) && (1990 -_ttoi(temp) > 0))
	{
		MessageBox(_T("hireDate ����"));
		return 1;
	}
	temp = hireDate.Mid(5,2);
	if ((_ttoi(temp)-12 > 0) && (0 -_ttoi(temp) > 0))
	{
		MessageBox(_T("hireDate ����"));
		return 1;
	}
	temp = hireDate.Right(2);
	if ((_ttoi(temp)-31 > 0) && (0 -_ttoi(temp) > 0))
	{
		MessageBox(_T("hireDate ����"));
		return 1;
	}
	//���ϼ��Ϸ���
	if (1 == mode)
	{
		return 0;
	}
	// ��鲿�ֲ�����Ψһ�� ���ϸ񷵻�ֵΪ 2 businessEntityID,nationalIDNumber
	
	CADOConnection adoConn;
	_bstr_t vSQL="select [BusinessEntityID] from [AdventureWorks2008].[HumanResources].[Employee] where BusinessEntityID = "+(_bstr_t)businessEntityID;
	_RecordsetPtr m_pRecordset;
	adoConn.OnInitAdo();
	m_pRecordset=adoConn.GetRecordset(vSQL);
	if(!(m_pRecordset->adoEOF))
	{
		MessageBox(_T("BusinessEntityID �ظ�"));
		return 2;
	}

	vSQL="select [NationalIDNumber] from [AdventureWorks2008].[HumanResources].[Employee] where NationalIDNumber = "+(_bstr_t)nationalIDNumber;
	
	m_pRecordset=adoConn.GetRecordset(vSQL);
	if(!(m_pRecordset->adoEOF))
	{
		MessageBox(_T("NationalIDNumber �ظ�"));
		return 2;
	}
	adoConn.ExitConnect();
	//�ϸ񷵻� 0 
	return 0;
}

  void StaffDlg::listUpdate()
 {
	 CADOConnection adoConn;
 	_RecordsetPtr m_pRecordset;
	CString businessEntityID,nationalIDNumber,jobTitle,maritalStatus,gender,hireDate,salariedFlag;
 	_bstr_t vSQL="select [BusinessEntityID],[NationalIDNumber],[JobTitle],[MaritalStatus],[Gender],[HireDate],[SalariedFlag] from [AdventureWorks2008].[HumanResources].[Employee]";
 	m_List.DeleteAllItems();
 	m_pRecordset = adoConn.GetRecordset(vSQL);
 	if(m_pRecordset->adoEOF)
 		return ;	
 	m_pRecordset->MoveFirst();
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
 }