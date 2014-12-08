
#if !defined(AFX_ADOCONN_H__INCLUDED_)
#define AFX_ADOCONN_H__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#pragma once
class CADOConnection
{
public:
	_RecordsetPtr m_pRecordset;//记录集指针
	_ConnectionPtr m_pConnection;//数据库连接指针

public:
	CADOConnection(void);
	virtual ~CADOConnection(void);
	void OnInitAdo();
	_RecordsetPtr& GetRecordset(_bstr_t bstrSQL);
	BOOL ExecuteSQL(_bstr_t bstrSQL);
	void ExitConnect();
};

#endif