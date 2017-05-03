
// MainTester.cpp : ����Ӧ�ó��������Ϊ��
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


// CMainTesterApp ����

CMainTesterApp::CMainTesterApp()
{
	// ֧����������������
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: �ڴ˴���ӹ�����룬
	// ��������Ҫ�ĳ�ʼ�������� InitInstance ��
}


// Ψһ��һ�� CMainTesterApp ����

CMainTesterApp theApp;


// CMainTesterApp ��ʼ��

BOOL CMainTesterApp::InitInstance()
{
	// ���һ�������� Windows XP �ϵ�Ӧ�ó����嵥ָ��Ҫ
	// ʹ�� ComCtl32.dll �汾 6 ����߰汾�����ÿ��ӻ���ʽ��
	//����Ҫ InitCommonControlsEx()�����򣬽��޷��������ڡ�
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// ��������Ϊ��������Ҫ��Ӧ�ó�����ʹ�õ�
	// �����ؼ��ࡣ
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();

	EnableAutomation();
	AfxEnableControlContainer();

	// ���� shell ���������Է��Ի������
	// �κ� shell ����ͼ�ؼ��� shell �б���ͼ�ؼ���
	CShellManager *pShellManager = new CShellManager;

	// ��׼��ʼ��
	// ���δʹ����Щ���ܲ�ϣ����С
	// ���տ�ִ���ļ��Ĵ�С����Ӧ�Ƴ�����
	// ����Ҫ���ض���ʼ������
	// �������ڴ洢���õ�ע�����
	// TODO: Ӧ�ʵ��޸ĸ��ַ�����
	// �����޸�Ϊ��˾����֯��
	SetRegistryKey(_T("Ӧ�ó��������ɵı���Ӧ�ó���"));

	CMainTesterDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: �ڴ˷��ô����ʱ��
		//  ��ȷ�������رնԻ���Ĵ���
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: �ڴ˷��ô����ʱ��
		//  ��ȡ�������رնԻ���Ĵ���
	}

	// ɾ�����洴���� shell ��������
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

	// ���ڶԻ����ѹرգ����Խ����� FALSE �Ա��˳�Ӧ�ó���
	//  ����������Ӧ�ó������Ϣ�á�
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