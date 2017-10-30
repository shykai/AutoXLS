
// AutoXLSDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "AutoXLS.h"
#include "AutoXLSDlg.h"
#include "afxdialogex.h"

#include "MakeXls.cpp"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAutoXLSDlg �Ի���



CAutoXLSDlg::CAutoXLSDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_AUTOXLS_DIALOG, pParent)
	, titleName(_T(""))
	, titleCount(1)
	, isPlus(FALSE)
	, stuCount(1)
	, plusTitle(_T("������"))
	, totalTitle(_T("�ܷ�"))
	, isSumScore(FALSE)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CAutoXLSDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_Title, titleName);
	DDX_Text(pDX, IDC_TitleCount, titleCount);
	DDX_Check(pDX, IDC_IsPlus, isPlus);
	DDX_Control(pDX, IDC_LIST, titleList);
	DDX_Text(pDX, IDC_StuCount, stuCount);
	DDX_Text(pDX, IDC_PlusTitle, plusTitle);
	DDX_Text(pDX, IDC_TotalTitle, totalTitle);
	DDX_Check(pDX, IDC_IsSum, isSumScore);
}

BEGIN_MESSAGE_MAP(CAutoXLSDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CAutoXLSDlg::OnBnClickedAdd)
	ON_BN_CLICKED(IDC_BUTTON2, &CAutoXLSDlg::OnBnClickedModify)
	ON_BN_CLICKED(IDC_BUTTON3, &CAutoXLSDlg::OnBnClickedRemove)
	ON_BN_CLICKED(IDC_Save, &CAutoXLSDlg::OnBnClickedSave)
	ON_NOTIFY(NM_CLICK, IDC_LIST, &CAutoXLSDlg::OnNMClickList)
	ON_BN_CLICKED(IDC_IsPlus, &CAutoXLSDlg::OnClickedIsplus)
END_MESSAGE_MAP()


// CAutoXLSDlg ��Ϣ�������

BOOL CAutoXLSDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	DWORD dwStyle = titleList.GetExtendedStyle();
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	titleList.SetExtendedStyle(dwStyle);

	titleList.InsertColumn(0, _T("���"), LVCFMT_LEFT, 40);
	titleList.InsertColumn(1, _T("��������"), LVCFMT_LEFT, 120);
	titleList.InsertColumn(2, _T("С������"), LVCFMT_LEFT, 60);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CAutoXLSDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CAutoXLSDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CAutoXLSDlg::OnBnClickedAdd()
{
	UpdateData(TRUE);

	int rowNo = titleList.GetItemCount();


	titleList.InsertItem(rowNo, toString(rowNo+1));//������
	
	titleList.SetItemText(rowNo, 1, titleName);//��������
	titleList.SetItemText(rowNo, 2, toString(titleCount));//��������

	titleName = "";
	titleCount = 1;

	GetDlgItem(IDC_Title)->SetFocus();

	UpdateData(FALSE);
}


void CAutoXLSDlg::OnBnClickedModify()
{
	UpdateData(TRUE);

	POSITION pos = titleList.GetFirstSelectedItemPosition();
	if (pos != NULL)
	{
		//�õ��кţ�ͨ��POSITIONת��
		int rowNo = (int)titleList.GetNextSelectedItem(pos);
		
		titleList.SetItemText(rowNo, 1, titleName);//��������
		titleList.SetItemText(rowNo, 2, toString(titleCount));//��������

		UpdateData(FALSE);
	}
}


void CAutoXLSDlg::OnBnClickedRemove()
{
	UpdateData(TRUE);

	int nItem = -1;
	POSITION pos;
	while (pos = titleList.GetFirstSelectedItemPosition())
	{

		nItem = -1;
		nItem = titleList.GetNextSelectedItem(pos);
		if (nItem >= 0 && titleList.GetSelectedCount() > 0)
		{
			titleList.DeleteItem(nItem);
		}
	}

	UpdateData(FALSE);
}


void CAutoXLSDlg::OnBnClickedSave()
{
	UpdateData(TRUE);

	TCHAR szFilter[] = _T("Excel�ļ�(*.xls)");
	// ���챣���ļ��Ի���   
	CFileDialog fileDlg(FALSE, _T("xls"), _T("�ɼ�ͳ�Ʊ�"), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, this);
	CString strFilePath;

	// ��ʾ�����ļ��Ի���   
	if (IDOK != fileDlg.DoModal())
	{

		return;
	}
	// ���������ļ��Ի����ϵġ����桱��ť����ѡ����ļ�·����ʾ���༭����   
	strFilePath = fileDlg.GetPathName();

	MatchExcel newExcel;

	MatchExcel::MatchMap inData;
	inData.stuCount = stuCount;
	inData.totalTitle = CA2W((LPCSTR)totalTitle);
	
	inData.isPlusNode = isPlus;
	inData.plusTitle = CA2W((LPCSTR)plusTitle);

	inData.isSum = isSumScore;

	for (int i = 0; i < titleList.GetItemCount(); i++)
	{
		CString name = titleList.GetItemText(i, 1);

		CString titleNum = titleList.GetItemText(i, 2);
		int count = atoi(titleNum);

		std::wstring wStr = CA2W((LPCSTR)name);

		inData.nodeList.push_back(MatchExcel::MatchNode(wStr, count));

	}
	inData.isPlusNode = isPlus;

	newExcel.inputExcel(inData);
	if (newExcel.outputExcel(strFilePath.GetBuffer()))
	{
		MessageBox("����ɹ�");
	}
	else
	{
		MessageBox("�ļ�����ʧ�ܣ�");
	}
}

void CAutoXLSDlg::OnClickedIsplus()
{
	UpdateData(TRUE);

	GetDlgItem(IDC_PlusTitle)->EnableWindow(isPlus);
	GetDlgItem(IDC_IsSum)->EnableWindow(isPlus);

	UpdateData(FALSE);
}


void CAutoXLSDlg::OnNMClickList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);

	POSITION pos = titleList.GetFirstSelectedItemPosition();
	if (pos != NULL)
	{
		//�õ��кţ�ͨ��POSITIONת��
		int rowNo = (int)titleList.GetNextSelectedItem(pos);

		titleName = titleList.GetItemText(rowNo, 1);
		
		CString titleNum = titleList.GetItemText(rowNo, 2);
		titleCount = atoi(titleNum);

		UpdateData(FALSE);
	}

	// TODO: �ڴ���ӿؼ�֪ͨ����������
	*pResult = 0;

}



CString CAutoXLSDlg::toString(int numVal)
{
	char tmpNum[8] = { 0 };
	int len = snprintf(tmpNum, 8, "%d", numVal);
	tmpNum[len] = 0;

	return CString(tmpNum);
}


BOOL CAutoXLSDlg::PreTranslateMessage(MSG* pMsg)
{
	// TODO: �ڴ����ר�ô����/����û���

	if (VK_RETURN == pMsg->wParam)
		return true;

	return CDialogEx::PreTranslateMessage(pMsg);
}

