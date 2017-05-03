
// MainTester.cpp : 定义应用程序的类行为。
//

#include "stdafx.h"
#include "MainTester.h"
#include "MainTesterDlg.h"

#include "IEAdapterMFCTest.h"


#include <exdispid.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


//typedef	CComQIPtr<IWebBrowser2> CComQIPtrIWebBrowser2;

// CMainTesterApp

BEGIN_MESSAGE_MAP(CMainTesterApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


BEGIN_DISPATCH_MAP(CMainTesterApp, CWinApp)
	//DISP_FUNCTION_ID(CIEAdapterMFCApp, "OnQuit", DISPID_ONQUIT, OnQuit, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CMainTesterApp, "DocumentComplete", DISPID_DOCUMENTCOMPLETE, 
	                 DocumentComplete, VT_EMPTY, VTS_DISPATCH VTS_PVARIANT)
END_DISPATCH_MAP()


// CMainTesterApp 构造

CMainTesterApp::CMainTesterApp()
{
	// 支持重新启动管理器
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: 在此处添加构造代码，
	// 将所有重要的初始化放置在 InitInstance 中
}


// 唯一的一个 CMainTesterApp 对象

CMainTesterApp theApp;


// CMainTesterApp 初始化

BOOL CMainTesterApp::InitInstance()
{
	// 如果一个运行在 Windows XP 上的应用程序清单指定要
	// 使用 ComCtl32.dll 版本 6 或更高版本来启用可视化方式，
	//则需要 InitCommonControlsEx()。否则，将无法创建窗口。
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// 将它设置为包括所有要在应用程序中使用的
	// 公共控件类。
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();

	EnableAutomation();
	AfxEnableControlContainer();

	// 创建 shell 管理器，以防对话框包含
	// 任何 shell 树视图控件或 shell 列表视图控件。
	CShellManager *pShellManager = new CShellManager;

	// 标准初始化
	// 如果未使用这些功能并希望减小
	// 最终可执行文件的大小，则应移除下列
	// 不需要的特定初始化例程
	// 更改用于存储设置的注册表项
	// TODO: 应适当修改该字符串，
	// 例如修改为公司或组织名
	SetRegistryKey(_T("应用程序向导生成的本地应用程序"));

	CMainTesterDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: 在此放置处理何时用
		//  “确定”来关闭对话框的代码
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: 在此放置处理何时用
		//  “取消”来关闭对话框的代码
	}

	// 删除上面创建的 shell 管理器。
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

	// 由于对话框已关闭，所以将返回 FALSE 以便退出应用程序，
	//  而不是启动应用程序的消息泵。
	return FALSE;
}


void CMainTesterApp::DocumentComplete(IDispatch *pDisp, VARIANT *URL)
{
	//CComQIPtr<IUnknown,&IID_IUnknown> pWBUK(m_pWebBrowser2);
	//CComQIPtr<IUnknown,&IID_IUnknown> pSenderUK(pDisp);

	//if (pWBUK == pSenderUK)
	/*{
		CComQIPtr<IDispatch> pHTMLDocDisp;
		m_pWebBrowser2->get_Document(&pHTMLDocDisp);
		CComQIPtr<IHTMLDocument2> pHTMLDoc(pHTMLDocDisp);

		CComQIPtr<IHTMLElementCollection> ecAll;
		CComPtr<IDispatch> pTagLineDisp;
		if( pHTMLDoc )
		{
			pHTMLDoc->get_all(&ecAll);
		}
		if (ecAll)
		{

		}
	}*/
	if (!m_eventHandler)
	{
		return;
	}

	m_eventHandler->notify(pDisp, DISPID_DOCUMENTCOMPLETE);
}



void CMainTesterApp::registerEventHandler(boost::shared_ptr<CEventHandler> apEventHandler)
{
	m_eventHandler = apEventHandler;
}




/****************    class  CMainTester    *********************/

CMainTester::CMainTester()
{
	CoInitialize (NULL);
}


CMainTester::~CMainTester()
{
	//m_pWebBrowser2=(LPUNKNOWN)NULL;
	CoUninitialize();
}


bool CMainTester::Init()
{
	::ShellExecute(NULL, L"open", L"iexplore.exe", NULL, NULL, SW_SHOW);
	::Sleep(3000);
	getTopIE();
	if (!m_pWebBrowser2)
	{
		/*m_pWebBrowser2.CoCreateInstance(CLSID_InternetExplorer);
		m_pWebBrowser2->put_Visible(VARIANT_TRUE);*/
		return false;
	}

	return true;
}


bool CMainTester::getTopIE()
{
	CComPtr<IShellWindows> psw;
	psw.CoCreateInstance(CLSID_ShellWindows);
	if(psw)
	{
		//array to storage IE window handles and interfaces
		CDWordArray	arHWNDShellWindows;
		arHWNDShellWindows.SetSize(0,10);	//grow by 10 
		CTypedPtrArray<CPtrArray, CComQIPtr<IWebBrowser2>*> arShellWindows;
		arShellWindows.SetSize(0,10);//grow by 10 
		//enumerate the ShellWindow collection 
		long lShellWindowCount=0;
		psw->get_Count(&lShellWindowCount);
		for(long i=0;i<lShellWindowCount;i++)
		{
			CComPtr<IDispatch> pdispShellWindow;
			psw->Item(COleVariant(i),&pdispShellWindow);
			CComQIPtr<IWebBrowser2> pIE(pdispShellWindow);
			if(pIE)//is it a Shell window?
			{
				//is it the right type?
				CString strWindowClass=GetWindowClassName(pIE);
				if(strWindowClass == _T("IEFrame"))
				{
					HWND hWndID=NULL;
					pIE->get_HWND((long*)&hWndID);
					//store its information
					arHWNDShellWindows.Add((DWORD)hWndID);
					arShellWindows.Add(new CComQIPtr<IWebBrowser2>(pIE));
				}
			}	
		}
		if(arHWNDShellWindows.GetSize()>0)//at least one shell window found
		{
			BOOL bFound=FALSE;
			//the first top-level window in zorder
			HWND hwndTest = ::GetWindow((HWND)arHWNDShellWindows[0],GW_HWNDFIRST);
			while( hwndTest&& !bFound)
			{
				for(int i = 0; i < arHWNDShellWindows.GetSize(); i++)
				{
					if(hwndTest == (HWND)arHWNDShellWindows[i])
					{
						//got it, attach to it
						m_pWebBrowser2 = *arShellWindows[i];
						//AdviseSinkIE();
						//NavigateToSamplePage(bIE);
						bFound=TRUE;
						//m_bOwnIE=FALSE;
						break;
					}
				}
				hwndTest = ::GetWindow(hwndTest, GW_HWNDNEXT);
			}
		}
		//cleanup
		for(int i=0;i<arShellWindows.GetSize();i++)
		{
			delete arShellWindows[i];
		}
	}
	return true;
}


CString	CMainTester::GetWindowClassName(IWebBrowser2* pwb)
{
	TCHAR szClassName[_MAX_PATH];
	ZeroMemory(szClassName,_MAX_PATH*sizeof(TCHAR));
	HWND hwnd=NULL;
	if(pwb)
	{
		LONG_PTR lwnd=NULL;
		pwb->get_HWND(&lwnd);
		hwnd=reinterpret_cast<HWND>(lwnd);
		::GetClassName(hwnd,szClassName,_MAX_PATH);
	}
	return szClassName;
}



int CMainTester::run()
{
	int err = 0;	

	if (!getTopIE())
	{
		err = -1;
		return err;
	}

	CIEAdapterMFCTest test;

	test.run(this);

	return err;
}