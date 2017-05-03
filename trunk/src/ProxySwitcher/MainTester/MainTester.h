
// MainTester.h : PROJECT_NAME 应用程序的主头文件
//

#pragma once

#ifndef __AFXWIN_H__
	#error "在包含此文件之前包含“stdafx.h”以生成 PCH 文件"
#endif

#include "resource.h"		// 主符号

#include "boost\shared_ptr.hpp";

class CEventHandler;
// CMainTesterApp:
// 有关此类的实现，请参阅 MainTester.cpp
//

class CMainTesterApp : public CWinApp
{
public:
	CMainTesterApp();

	void registerEventHandler(boost::shared_ptr<CEventHandler> apEventHandler);

// 重写
public:
	virtual BOOL InitInstance();


	afx_msg void DocumentComplete(IDispatch *pDisp, VARIANT *URL);
	//afx_msg void OnQuit();


// 实现

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

