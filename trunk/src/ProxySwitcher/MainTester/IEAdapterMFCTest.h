#pragma once

#include "stdafx.h"

#include "IEAdapterMFC.h"
#include "MainTester.h"


class CIEAdapterMFCTest
{
public:
	CIEAdapterMFCTest();

	~CIEAdapterMFCTest();

	int run(CMainTester* tester);

	//CComQIPtr<IWebBrowser2> m_pWebBrowser2;
};


