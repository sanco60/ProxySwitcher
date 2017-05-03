#include "stdafx.h"
#include "IEAdapterMFCTest.h"

#include "IEAdapterMFC.h"

CIEAdapterMFCTest::CIEAdapterMFCTest()
{}


CIEAdapterMFCTest::~CIEAdapterMFCTest()
{}


int CIEAdapterMFCTest::run(CMainTester* tester)
{
	int err = 0;

	if (NULL == tester || !tester->m_pWebBrowser2)
	{
		err = -1;
		return err;
	}

	CIEAdapterMFC ieadapter(tester->m_pWebBrowser2);

	//ieadapter.registerEvents(theApp.GetIDispatch(FALSE));

	//ÇÐ»»´úÀí
	/*UINT rtExec = ::WinExec("ProxySet.exe 10.10.10.111 90", SW_NORMAL);
	::Sleep(1000);*/

	std::string szUrl("http://www.sse.com.cn/disclosure/listedinfo/announcement/");
	ieadapter.navigate(szUrl);
	::Sleep(3000);

	std::vector<CComQIPtr<IHTMLElement>> elemVec;
	std::string szTag("a");
	std::vector<std::string> attrs;
	attrs.push_back(std::string("href"));
	attrs.push_back(std::string("title"));

	ieadapter.searchItems(szTag, attrs, elemVec);

	//ieadapter.unregisterEvents(theApp.GetIDispatch(FALSE));

	return err;
}


