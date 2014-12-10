#pragma once
#include "afxcmn.h"


// StaffDlg 对话框

class StaffDlg : public CDialogEx
{
	DECLARE_DYNAMIC(StaffDlg)

public:
	StaffDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~StaffDlg();

// 对话框数据
	enum { IDD = IDD_DIALOG_STAFF };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton4();
	CString m_findBusinessEntityID;
	virtual BOOL OnInitDialog();
	CListCtrl m_List;

	afx_msg void OnBnClickedButton1();
	afx_msg void OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
private:
	int ruleCheck(int mode);
	void listUpdate();
};
