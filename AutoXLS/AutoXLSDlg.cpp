
// AutoXLSDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "AutoXLS.h"
#include "AutoXLSDlg.h"
#include "afxdialogex.h"

#include "MakeXls_V2.cpp"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAutoXLSDlg 对话框



CAutoXLSDlg::CAutoXLSDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_AUTOXLS_DIALOG, pParent)
	, titleName(_T(""))
	, titleCount(1)
	, isPlus(FALSE)
	, stuCount(1)
	, plusTitle(_T("附加题"))
	, totalTitle(_T("总分"))
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


// CAutoXLSDlg 消息处理程序

BOOL CAutoXLSDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	DWORD dwStyle = titleList.GetExtendedStyle();
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	titleList.SetExtendedStyle(dwStyle);

	titleList.InsertColumn(0, _T("序号"), LVCFMT_LEFT, 40);
	titleList.InsertColumn(1, _T("大题名称"), LVCFMT_LEFT, 120);
	titleList.InsertColumn(2, _T("小题数量"), LVCFMT_LEFT, 60);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CAutoXLSDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CAutoXLSDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CAutoXLSDlg::OnBnClickedAdd()
{
	UpdateData(TRUE);

	int rowNo = titleList.GetItemCount();


	titleList.InsertItem(rowNo, toString(rowNo+1));//插入行
	
	titleList.SetItemText(rowNo, 1, titleName);//设置数据
	titleList.SetItemText(rowNo, 2, toString(titleCount));//设置数据

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
		//得到行号，通过POSITION转化
		int rowNo = (int)titleList.GetNextSelectedItem(pos);
		
		titleList.SetItemText(rowNo, 1, titleName);//设置数据
		titleList.SetItemText(rowNo, 2, toString(titleCount));//设置数据

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
	GetDlgItem(IDC_Save)->SetWindowText("生成中...");
	GetDlgItem(IDC_Save)->EnableWindow(FALSE);

	UpdateData(TRUE);

	TCHAR szFilter[] = _T("Excel文件(*.xls)");
	// 构造保存文件对话框   
	CFileDialog fileDlg(FALSE, _T("xls"), _T("成绩统计表"), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, this);
	CString strFilePath;

	// 显示保存文件对话框   
	if (IDOK != fileDlg.DoModal())
	{

		return;
	}
	// 如果点击了文件对话框上的“保存”按钮，则将选择的文件路径显示到编辑框里   
	strFilePath = fileDlg.GetPathName();

	MatchExcel_V2 newExcel;

	MatchMap inData;
	inData.stuCount = 35;
	inData.isSum = true;

	{
		MatchNodes firstClassNodes;
		firstClassNodes.push_back(MatchNode(L"1", 1));
		firstClassNodes.push_back(MatchNode(L"2", 1));
		firstClassNodes.push_back(MatchNode(L"3", 1));
		firstClassNodes.push_back(MatchNode(L"4", 1));
		firstClassNodes.push_back(MatchNode(L"5", 10));
		firstClassNodes.push_back(MatchNode(L"6", 5));
		firstClassNodes.push_back(MatchNode(L"7", 1));
		firstClassNodes.push_back(MatchNode(L"8", 1));
		firstClassNodes.push_back(MatchNode(L"9", 5));

		MatchClass firstClass(L"A", firstClassNodes);
		inData.nodeList.push_back(firstClass);
	}

	{
		MatchNodes firstClassNodes;
		firstClassNodes.push_back(MatchNode(L"B", 1));

		MatchClass firstClass(L"B");
		inData.nodeList.push_back(firstClass);
	}


	if (inData.nodeList.size() > 0)
	{
		newExcel.inputExcel(inData);
		if (newExcel.outputExcel(strFilePath.GetBuffer()))
		{
			MessageBox("保存成功");
		}
		else
		{
			MessageBox("错误：文件保存失败！");
		}
	}
	else
	{
		MessageBox("错误：题目数不可为空！");
	}

	GetDlgItem(IDC_Save)->SetWindowText("生成文件");
	GetDlgItem(IDC_Save)->EnableWindow(TRUE);

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
		//得到行号，通过POSITION转化
		int rowNo = (int)titleList.GetNextSelectedItem(pos);

		titleName = titleList.GetItemText(rowNo, 1);
		
		CString titleNum = titleList.GetItemText(rowNo, 2);
		titleCount = atoi(titleNum);

		UpdateData(FALSE);
	}

	// TODO: 在此添加控件通知处理程序代码
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
	// TODO: 在此添加专用代码和/或调用基类

	if (VK_RETURN == pMsg->wParam)
		return true;

	return CDialogEx::PreTranslateMessage(pMsg);
}

