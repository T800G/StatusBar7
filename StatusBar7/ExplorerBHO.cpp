#include "stdafx.h"
#include <Commctrl.h>
#pragma comment(lib,"Comctl32.lib")
#include "StatusBar7.h"
#include "ExplorerBHO.h"
#include "helpers.h"

#define SBTXTBUFLEN 256//1024
#define SBDELAY 50
wchar_t g_SBText[SBTXTBUFLEN];


HRESULT CExplorerBHO::FinalConstruct()
{
	ATLTRACE("CExplorerBHO::FinalConstruct\n");
	//don't load BHO in IE7/8
	TCHAR szExePath[MAX_PATH];
	if (!GetModuleFileName(NULL, szExePath, MAX_PATH)) return HRESULT_FROM_WIN32(GetLastError());
	if (StrCmpI(PathFindFileName(szExePath), _T("iexplore.exe"))==0) return E_ABORT;

return S_OK;
}
void CExplorerBHO::FinalRelease()
{
	ATLTRACE("CExplorerBHO::FinalRelease\n");
}

//https://msdn.microsoft.com/en-us/library/windows/desktop/bb773183(v=vs.85).aspx
//?static CWindowSubClass::WindowSubClassProc   //?pure virtual LRESULT WndSubClassProc()=0
LRESULT CALLBACK WindowSubClassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR pThis)
{
	return  ((CExplorerBHO*)pThis)->WndSubClassProc(hWnd,uMsg,wParam,lParam,uIdSubclass);
}

LRESULT CExplorerBHO::WndSubClassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass)
{
	if (uMsg==SB_SETTEXT && LOBYTE(LOWORD(wParam))!=255) //255 /shell context menu is active (shows menuitem descriptions in StatusBar)
	{
		//ATLTRACE(L"SB_SETTEXT hWnd=%x HIBYTE=%d LOBYTE=%d lParam=%ws\n",hWnd,HIBYTE(LOWORD(wParam)),LOBYTE(LOWORD(wParam)),lParam);
		if (HIBYTE(LOWORD(wParam))!=SBT_OWNERDRAW)
		{
			if SUCCEEDED(StringCchCopy(g_SBText, SBTXTBUFLEN, (LPCTSTR)lParam))
				SetTimer(m_hwStatusBar, (UINT_PTR)m_hwStatusBar, SBDELAY, NULL);
		}
	}

	if (uMsg==WM_TIMER && wParam==(WPARAM)m_hwStatusBar)
	{
		KillTimer(hWnd,wParam);
		__int64 size=-1;
		CComPtr<IShellView> pShellView;
		if SUCCEEDED(m_pShellBrowser->QueryActiveShellView(&pShellView))
		{
			CComQIPtr<IFolderView> pFolderView=pShellView;
			CComPtr<IPersistFolder2> pPersistFolder2;
			if (pFolderView && SUCCEEDED(pFolderView->GetFolder(IID_IPersistFolder2,(void**)&pPersistFolder2)))
			{
				CComQIPtr<IShellFolder2> pShellFolder2=pPersistFolder2;
				int count;
				if (pShellFolder2 && SUCCEEDED(pFolderView->ItemCount(SVGIO_SELECTION,&count)) && count)
				{
					//ATLTRACE("selection count=%d\n",count);
					CComPtr<IEnumIDList> pEnum;
					if (SUCCEEDED(pFolderView->Items(SVGIO_SELECTION,IID_IEnumIDList,(void**)&pEnum)) && pEnum)
					{
						PITEMID_CHILD pidc;
						SHCOLUMNID scid={PSGUID_STORAGE,PID_STG_SIZE};
						while (pEnum->Next(1,&pidc,NULL)==S_OK)
						{
							CComVariant cvt;
							if (SUCCEEDED(pShellFolder2->GetDetailsEx(pidc,&scid,&cvt)) && cvt.vt==VT_UI8)
							{
								if (size<0) size=cvt.ullVal;
								else size+=cvt.ullVal;
							}
							ILFree(pidc);
						}
					}
				}
			}
		}

		ATLTRACE("selection size=%d\n", size);
		if (size>=0)
		{
			// format the file size as KB, MB, etc
			wchar_t pbuf[SBTXTBUFLEN];
			if (StrFormatByteSizeW(size,pbuf,SBTXTBUFLEN))
			{
				//ATLTRACE(L"(%ws)\n",pbuf);
				wchar_t pNew[SBTXTBUFLEN];
				if SUCCEEDED(StringCchPrintfW(pNew, SBTXTBUFLEN, _T("%s  (%s)"), g_SBText, pbuf))
					SendMessage(m_hwStatusBar, WM_SETTEXT, NULL, (LPARAM)((wchar_t*)pNew));
			}
		}
	}
return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

HRESULT CExplorerBHO::SetSite(IUnknown *pUnkSite)
{
	IObjectWithSiteImpl<CExplorerBHO>::SetSite(pUnkSite);
	if (pUnkSite)
	{
		CComPtr<IServiceProvider> psp;
		if SUCCEEDED(pUnkSite->QueryInterface(IID_IServiceProvider,(void**)&psp))
		{
			if (SUCCEEDED(psp->QueryService( SID_SWebBrowserApp, IID_IWebBrowser2, (void**)&m_pWebBrowser2)) &&
				SUCCEEDED(psp->QueryService(SID_SShellBrowser, IID_IShellBrowser, (void**)&m_pShellBrowser)))
			{
				if (m_dwEventCookie==0xFEFEFEFE) DispEventAdvise( m_pWebBrowser2, &IID_IDispatch);
				ATLASSERT(m_hwStatusBar==NULL);
				if SUCCEEDED(m_pShellBrowser->GetControlWindow(FCW_STATUS, &m_hwStatusBar))
				{
					ATLTRACE("msctls_statusbar32=%x\n", m_hwStatusBar);
					if (!SetWindowSubclass(m_hwStatusBar, (SUBCLASSPROC)WindowSubClassProc, (UINT_PTR)m_hwStatusBar, (DWORD_PTR)this)) m_hwStatusBar=NULL;
				}

			}
		}
	}
	else
	{
		ATLTRACE("SetSite:: NULL\n");
		if (IsWindow(m_hwStatusBar))
		{
			ATLTRACE("msctls_statusbar32=%x\n", m_hwStatusBar);	
			if (RemoveWindowSubclass(m_hwStatusBar, (SUBCLASSPROC)WindowSubClassProc, (UINT_PTR)m_hwStatusBar)) m_hwStatusBar=NULL;
//?handle error/unsubclass on WM_DESTROY inside pfnSubclass
		}

		if ((m_dwEventCookie != 0xFEFEFEFE) && m_pWebBrowser2) DispEventUnadvise(m_pWebBrowser2, &IID_IDispatch);
		m_pWebBrowser2.Release();
		m_pShellBrowser.Release();
	}
return S_OK;//SetSite allways returns S_OK
}


HRESULT CExplorerBHO::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags,
					DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
{
	HRESULT hr=IDispatchImpl<IExplorerBHO, &IID_IExplorerBHO,
								&LIBID_STATUSBAR7Lib,1>::Invoke(dispidMember,
													riid, lcid,wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr);

	if (dispidMember==DISPID_ONQUIT)
		if((m_dwEventCookie!=0xFEFEFEFE) && m_pWebBrowser2)
			return DispEventUnadvise(m_pWebBrowser2, &IID_IDispatch);

return hr;
}
