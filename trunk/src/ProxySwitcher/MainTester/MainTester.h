
// MainTester.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������

#include "boost\shared_ptr.hpp";

class CEventHandler;
// CMainTesterApp:
// �йش����ʵ�֣������ MainTester.cpp
//

class CMainTesterApp : public CWinApp
{
public:
	CMainTesterApp();

	void registerEventHandler(boost::shared_ptr<CEventHandler> apEventHandler);

// ��д
public:
	virtual BOOL InitInstance();


	afx_msg void DocumentComplete(IDispatch *pDisp, VARIANT *URL);
	//afx_msg void OnQuit();


// ʵ��

	DECLARE_MESSAGE_MAP()

	DECLARE_DISPATCH_MAP()
	
private:

	boost::shared_ptr<CEventHandler> m_eventHandler;
};

extern CMainTesterApp theApp;


class CMainTester
{
public:
	CMainTester();

	~CMainTester();

	bool Init();

	int run();

	bool getTopIE();

	CString	GetWindowClassName(IWebBrowser2* pwb);

	CComQIPtr<IWebBrowser2> m_pWebBrowser2;
};



class CEventHandler
{
public:
	virtual void notify(IDispatch *pDisp, DWORD eventID) = 0;
};

