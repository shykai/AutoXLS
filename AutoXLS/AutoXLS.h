
// AutoXLS.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CAutoXLSApp: 
// �йش����ʵ�֣������ AutoXLS.cpp
//

class CAutoXLSApp : public CWinApp
{
public:
	CAutoXLSApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CAutoXLSApp theApp;
