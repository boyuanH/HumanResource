#include "stdafx.h"
#include "ADOConnection.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

CADOConnection::CADOConnection(void)
{
}


CADOConnection::~CADOConnection(void)
{
}

void CADOConnection::OnInitAdo()
{
	::CoInitialize(NULL);
	try
	{
		m_pConnection.CreateInstance("ADODB.Connection");
		_bstr_t strConnect;
		//  strConnect="Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AdventureWorks2008;Data Source=HUANGBOYUAN-PC";
		strConnect="File Name=ADO.udl";
		m_pConnection->Open(strConnect,"","",adModeUnknown);
	}
	catch(_com_error e)
	{
		AfxMessageBox(_T("打开数据库失败"));
		AfxMessageBox(e.ErrorMessage());
	}
}
_RecordsetPtr& CADOConnection::GetRecordset(_bstr_t bstrSQL)
{
	try
	{
		if(m_pConnection==NULL)
			OnInitAdo();
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);

	}
	catch(_com_error e)
	{
		AfxMessageBox(_T("打开记录集失败"));
		AfxMessageBox(e.ErrorMessage());
	}
	return m_pRecordset;
}
BOOL CADOConnection::ExecuteSQL(_bstr_t bstrSQL)
{
	try
	{
		if(m_pConnection==NULL)
			OnInitAdo();
		m_pConnection->Execute(bstrSQL,NULL,adCmdText);
		return true;
	}
	catch(_com_error e)
	{
		CString str;
		str.Format(_T("不能打开记录集!%s"),e.ErrorMessage());
		AfxMessageBox(str);
		return false;
	}
}
void CADOConnection::ExitConnect()
{
	if(m_pRecordset!=NULL)
		m_pRecordset->Close();
	m_pConnection->Close();
	::CoUninitialize();
}