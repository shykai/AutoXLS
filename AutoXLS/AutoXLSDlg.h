
// AutoXLSDlg.h : ͷ�ļ�
//

#pragma once
#include "afxwin.h"
#include "afxcmn.h"

#include "PublicDef.h"

#include <map>

// CAutoXLSDlg �Ի���
class CAutoXLSDlg : public CDialogEx
{
// ����
public:
	CAutoXLSDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_AUTOXLS_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	CString titleName;
	int titleCount;
	BOOL isPlus;
	CListCtrl titleList;
	int stuCount;
	CString plusTitle;
	CString totalTitle;
	BOOL isSumScore;

public:
	afx_msg void OnBnClickedAdd();
	afx_msg void OnBnClickedModify();
	afx_msg void OnBnClickedRemove();
	afx_msg void OnBnClickedSave();
	afx_msg void OnClickedIsplus();
	afx_msg void OnNMClickList(NMHDR *pNMHDR, LRESULT *pResult);

private:
	CString toString(int numVal);

	char * charToWchar(char *s) {

		int w_nlen = MultiByteToWideChar(CP_ACP, 0, s, -1, NULL, 0);

		char *ret;

		ret = (char*)malloc(sizeof(WCHAR)*w_nlen);

		memset(ret, 0, sizeof(ret));

		MultiByteToWideChar(CP_ACP, 0, s, -1, (WCHAR*)ret, w_nlen);

		return ret;

	}
public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnBnClickedRadioClassType();
	afx_msg void OnCbnSelchangeClass();
	afx_msg void OnBnClickedClassAdd();
	afx_msg void OnBnClickedClassDel();
	afx_msg void OnBnClickedClassReset();
	CComboBox classList;
	int m_classType;

private:
	struct strucNode
	{
		CString className;
		MatchClassType classType;
		MatchNodes nodeList;
	};
	typedef std::list<strucNode*> ClassMap;
	ClassMap classMap;

	strucNode* nowNode;

	void refreshNodeList();
};
