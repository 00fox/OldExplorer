#include "stdafx.h"
#include "OldExplorer.h"
#include "WebBrowser.h"

WebBrowser::WebBrowser(HWND _hWndParent)
{
	OleInitialize(NULL);

	iComRefCount = 0;
	::SetRect(&rObject, -300, -300, 300, 300);
	hWndParent = _hWndParent;

	if (CreateBrowser() == FALSE)
		return;

	setupOle();
	AddRef();

	ShowWindow(GetControlWindow(), SW_SHOW);

	//this->webBrowser2->put_FullScreen(VARIANT_TRUE);
	this->webBrowser2->put_Silent(VARIANT_TRUE);
	this->webBrowser2->put_MenuBar(VARIANT_FALSE);
	this->webBrowser2->put_ToolBar(VARIANT_FALSE);
	this->webBrowser2->put_AddressBar(VARIANT_FALSE);
	this->webBrowser2->put_StatusBar(VARIANT_FALSE);
	this->Navigate(_T("about:blank"));
}

WebBrowser::~WebBrowser()
{
	this->webBrowser2->Quit();
	this->webBrowser2->Release();
	OleUninitialize();
}

bool WebBrowser::CreateBrowser(bool showScrollbars)
{
	hasScrollbars = showScrollbars;
	HRESULT hr;
	hr = ::OleCreate(CLSID_WebBrowser, IID_IOleObject, OLERENDER_DRAW, 0, this, this, (void**)&oleObject);

	if (FAILED(hr))
		return FALSE;

	hr = oleObject->SetClientSite(this);
	hr = OleSetContainedObject(oleObject, TRUE);

	RECT posRect;
	::SetRect(&posRect, -300, -300, 300, 300);
	hr = oleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, this, -1, hWndParent, &posRect);
	if (FAILED(hr))
		return FALSE;

	hr = oleObject->QueryInterface(&this->webBrowser2);
	if (FAILED(hr))
		return FALSE;

	return TRUE;
}

void WebBrowser::setupOle()
{
	RECT rc;
	GetClientRect(hWndControl, &rc);

	HRESULT hr;
	IOleObject* oleObject = 0;
	hr = CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_INPROC_SERVER, IID_IOleObject, (void**)&oleObject);
	if (oleObject == 0)
		return;

	hr = oleObject->SetClientSite(this);
	if (hr != S_OK)
	{
		oleObject->Release();
		return;
	}

	hr = oleObject->SetHostNames(L"MyHost", L"MyDoc");
	if (hr != S_OK)
	{
		oleObject->Release();
		return;
	}

	hr = OleSetContainedObject(oleObject, TRUE);
	if (hr != S_OK)
	{
		oleObject->Release();
		return;
	}

	hr = oleObject->DoVerb(OLEIVERB_SHOW, 0, this, 0, hWndControl, &rc);
	if (hr != S_OK)
	{
		oleObject->Release();
		return;
	}

	bool connected = false;
	IConnectionPointContainer* cpc = 0;
	oleObject->QueryInterface(IID_IConnectionPointContainer, (void**)&cpc);
	if (cpc != 0)
	{
		IConnectionPoint* cp = 0;
		cpc->FindConnectionPoint(DIID_DWebBrowserEvents2, &cp);

		if (cp != 0)
		{
			cp->Advise(this->webBrowser2, &cookie);
			cp->Release();
			connected = true;
		}

		cpc->Release();
	}
	if (!connected)
	{
		oleObject->Release();
		return;
	}

	//oleObject->QueryInterface(IID_IWebBrowser2, (void**)webBrowser2);

	oleObject->Release();
}

RECT WebBrowser::PixelToHiMetric(const RECT& _rc)
{
	static bool s_initialized = false;
	static int s_pixelsPerInchX, s_pixelsPerInchY;
	if(!s_initialized)
	{
		HDC hdc = ::GetDC(0);
		s_pixelsPerInchX = ::GetDeviceCaps(hdc, LOGPIXELSX);
		s_pixelsPerInchY = ::GetDeviceCaps(hdc, LOGPIXELSY);
		::ReleaseDC(0, hdc);
		s_initialized = true;
	}

	RECT rc;
	rc.left = MulDiv(2540, _rc.left, s_pixelsPerInchX);
	rc.top = MulDiv(2540, _rc.top, s_pixelsPerInchY);
	rc.right = MulDiv(2540, _rc.right, s_pixelsPerInchX);
	rc.bottom = MulDiv(2540, _rc.bottom, s_pixelsPerInchY);
	return rc;
}

void WebBrowser::SetRect(const RECT& _rc)
{
	rObject = _rc;

	{
		RECT hiMetricRect = PixelToHiMetric(rObject);
		SIZEL sz;
		sz.cx = hiMetricRect.right - hiMetricRect.left;
		sz.cy = hiMetricRect.bottom - hiMetricRect.top;
		oleObject->SetExtent(DVASPECT_CONTENT, &sz);
	}

	RECT rect;
	GetClientRect(hWndParent, &rect);
	this->webBrowser2->put_Left(0);
	this->webBrowser2->put_Top(0);
	this->webBrowser2->put_Width(rect.right);
	this->webBrowser2->put_Height(rect.bottom);
	//this->webBrowser2->put_Width(1920);
	//this->webBrowser2->put_Height(1080);
	//this->webBrowser2->Release();

	if(oleInPlaceObject != 0)
		oleInPlaceObject->SetObjectRects(&rObject, &rObject);
}

// ----- Control methods -----

void WebBrowser::GoBack()
{
	this->webBrowser2->GoBack();
}

void WebBrowser::GoForward()
{
	this->webBrowser2->GoForward();
}

void WebBrowser::Navigate(wstring szUrl)
{
	bstr_t url(szUrl.c_str());
	variant_t flags(navNoHistory | navAllowAutosearch | navOpenNewForegroundTab);
						//navBrowserBar | navHyperlink | navNewWindowsManaged | navUntrustedForDownload
    this->webBrowser2->Navigate(url, &flags, 0, 0, 0);
}

void WebBrowser::Refresh(bool zoomreset)
{
	this->webBrowser2->Refresh();
	if (zoomreset)
		ZoomReset();
	else
	{
		SetScrollPos(SB_HORZ, 0);
		SetScrollPos(SB_VERT, 0);
	}
}

void WebBrowser::Stop()
{
	this->webBrowser2->Stop();
}

void WebBrowser::Home()
{
	this->webBrowser2->GoHome();
}

void WebBrowser::Search()
{
	this->webBrowser2->GoSearch();
}

void WebBrowser::Zoom(int zoomto)
{
	zoom = max(1, zoomto);
	CComVariant varZoom(zoom);
	this->webBrowser2->ExecWB(OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, &varZoom, NULL);
}

void WebBrowser::ZoomPlus(int zoomin)
{
	zoom = zoom + max(1, zoomin * (zoom / 10));
	CComVariant varZoom(zoom);
	this->webBrowser2->ExecWB(OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, &varZoom, NULL);
}

void WebBrowser::ZoomMinus(int zoomout)
{
	zoom = max(1, (zoom - max(1, zoomout * (zoom / 10))));
	CComVariant varZoom(zoom);
	this->webBrowser2->ExecWB(OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, &varZoom, NULL);
}

void WebBrowser::ZoomReset(int zoomto)
{
	zoom = max(1, zoomto);
	CComVariant varZoom(zoom);
	this->webBrowser2->ExecWB(OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, &varZoom, NULL);
	SetPosition(0, 0);
	SetScrollPos(SB_HORZ);
	SetScrollPos(SB_VERT);
}

void WebBrowser::FullScreen()
{
	fullscreen = !fullscreen;
	if (fullscreen)
		this->webBrowser2->put_FullScreen(VARIANT_TRUE);
	else
		this->webBrowser2->put_FullScreen(VARIANT_FALSE);
}

void WebBrowser::SetPosition(int x, int y)
{
	this->webBrowser2->put_Left(x);
	this->webBrowser2->put_Top(y);
}

void WebBrowser::SetScrollPos(unsigned char bar, long pos)
{
	IHTMLDocument2* doc = GetDoc();
	if (doc == NULL)
		return;

	CComPtr<IHTMLElement> body;
	HRESULT hr = doc->get_body(&body);
	if (body)
	{
		IHTMLElement2* pElement = NULL;
		hr = body->QueryInterface(IID_IHTMLElement2, (void**)&pElement);
		if (!SUCCEEDED(hr))
			return;

		//long scroll_height;
		//long scroll_width;
		//pElement->get_scrollHeight(&scroll_height);
		//pElement->get_scrollWidth(&scroll_width);

		if (bar == SB_HORZ)
			pElement->put_scrollLeft(pos);
		else if (bar == SB_VERT)
			pElement->put_scrollTop(pos);

		CComPtr<IHTMLBodyElement> bodyElement;
		hr = body->QueryInterface(IID_IHTMLBodyElement, (void**)&bodyElement);
		if (bodyElement)
			bodyElement->put_scroll(CComBSTR(L"no")); // - masquer la barre de défilement
	}
}

void WebBrowser::BeforeNavigate(std::string url, bool* cancel)
{
	static const std::string exconst = "EXECUTE://";
	*cancel = false;

	bool execute = true;
	size_t i = 0;
	while (execute && i < exconst.length())
	{
		if (std::toupper(url[i], std::locale()) != exconst[i])
			execute = false;
		++i;
	}

	if (execute)
	{
		std::string command = url.substr(exconst.length());
		if (command[command.length() - 1] == '/')
			command = command.substr(0, command.length() - 1);
		*cancel = true;
	}
}

void WebBrowser::NavigateComplete(std::string url, WebBrowser* webForm)
{
	static JSObject* jsobj = new JSObject();
	jsobj->AddRef();
	webForm->AddCustomObject(jsobj, "JSObject");

	CComPtr<IHTMLDocument2> doc = GetDoc();
	CComPtr<IHTMLElement> pElement;
	HRESULT hr = doc->get_body(&pElement);
	if (pElement)
	{
		CComPtr<IHTMLBodyElement> pBodyElement;
		hr = pElement->QueryInterface(IID_IHTMLBodyElement, (void**)&pBodyElement);
		if (pBodyElement)
			pBodyElement->put_scroll(CComBSTR(L"no")); // - masquer la barre de défilement
	}
}

void WebBrowser::DocumentComplete(const wchar_t*)
{
	isnaving &= ~2;

	if (isnaving & 4)
		return;

	WPARAM w = (GetWindowLong(hWndControl, GWL_ID) & 0xFFFF) | ((WEBFN_LOADED & 0xFFFF) << 16);
	PostMessage(hWndParent, WM_COMMAND, w, (LPARAM)hWndControl);
}

void WebBrowser::AddCustomObject(IDispatch* custObj, std::string name)
{
	IHTMLDocument2* doc = GetDoc();

	if (doc == NULL)
		return;

	IHTMLWindow2* win = NULL;
	doc->get_parentWindow(&win);
	doc->Release();

	if (win == NULL)
		return;

	IDispatchEx* winEx;
	win->QueryInterface(&winEx);
	win->Release();

	if (winEx == NULL)
		return;

	int lenW = MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name.c_str(), -1, NULL, 0);
	BSTR objName = SysAllocStringLen(0, lenW);
	MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name.c_str(), -1, objName, lenW);

	DISPID dispid;
	if (objName == NULL)
		return;

	HRESULT hr = winEx->GetDispID(objName, fdexNameEnsure, &dispid);

	SysFreeString(objName);

	if (FAILED(hr))
		return;

	DISPID namedArgs[] = { DISPID_PROPERTYPUT };
	DISPPARAMS params;
	params.rgvarg = new VARIANT[1];
	params.rgvarg[0].pdispVal = custObj;
	params.rgvarg[0].vt = VT_DISPATCH;
	params.rgdispidNamedArgs = namedArgs;
	params.cArgs = 1;
	params.cNamedArgs = 1;

	hr = winEx->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, NULL, NULL, NULL);
	winEx->Release();

	if (FAILED(hr))
		return;
}

IHTMLDocument2* WebBrowser::GetDoc()
{
	IDispatch* dispatch = 0;
	this->webBrowser2->get_Document(&dispatch);

	if (dispatch == NULL)
		return NULL;

	IHTMLDocument2* doc;
	dispatch->QueryInterface(IID_IHTMLDocument2, (void**)&doc);
	dispatch->Release();

	return doc;
}

// ----- IUnknown -----

HRESULT STDMETHODCALLTYPE WebBrowser::QueryInterface(REFIID riid, void**ppvObject)
{
	if (riid == __uuidof(IUnknown))
		(*ppvObject) = static_cast<IOleClientSite*>(this);
	else if (riid == __uuidof(IOleInPlaceSite))
		(*ppvObject) = static_cast<IOleInPlaceSite*>(this);
	else
		return E_NOINTERFACE;

	AddRef();
	return S_OK;
}

ULONG STDMETHODCALLTYPE WebBrowser::AddRef(void)
{
	iComRefCount++; 

	return iComRefCount;
}

ULONG STDMETHODCALLTYPE WebBrowser::Release(void)
{
	iComRefCount--; 

	return iComRefCount;
}

// ---------- IDispatch ----------

HRESULT STDMETHODCALLTYPE WebBrowser::Invoke(DISPID dispIdMember, REFIID riid,
	LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult,
	EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	// DWebBrowserEvents2
	switch (dispIdMember)
	{
	case DISPID_BEFORENAVIGATE2:
	{
		BSTR bstrUrl = pDispParams->rgvarg[5].pvarVal->bstrVal;
		char* lpstrUrl = NULL;

		lpstrUrl = BSTRToLPSTR(bstrUrl, lpstrUrl);
		if (lpstrUrl == NULL)
			break;

		std::string url = lpstrUrl;
		delete[] lpstrUrl;

		bool cancel = false;

		BeforeNavigate(url, &cancel);

		// Set Cancel parameter to TRUE to cancel the current event
		*(((*pDispParams).rgvarg)[0].pboolVal) = cancel ? TRUE : FALSE;

		break;
	}
	case DISPID_DOCUMENTCOMPLETE:
		DocumentComplete(pDispParams->rgvarg[0].pvarVal->bstrVal);
		break;
	case DISPID_NAVIGATECOMPLETE2:
	{
		BSTR bstrUrl = pDispParams->rgvarg[0].pvarVal->bstrVal;
		char* lpstrUrl = NULL;

		lpstrUrl = BSTRToLPSTR(bstrUrl, lpstrUrl);
		if (lpstrUrl == NULL)
			break;

		std::string url = lpstrUrl;
		delete[] lpstrUrl;

		NavigateComplete(url, this);

		break;
	}
	case DISPID_AMBIENT_DLCONTROL:
		pVarResult->vt = VT_I4;
		pVarResult->lVal = DLCTL_DLIMAGES | DLCTL_VIDEOS | DLCTL_BGSOUNDS | DLCTL_SILENT;
		break;
	default:
		return DISP_E_MEMBERNOTFOUND;
	}

	return S_OK;
}

// ---------- IDocHostUIHandler ----------

HRESULT STDMETHODCALLTYPE WebBrowser::GetHostInfo(DOCHOSTUIINFO* pInfo)
{
	pInfo->dwFlags = (hasScrollbars ? 0 : DOCHOSTUIFLAG_SCROLL_NO) | DOCHOSTUIFLAG_NO3DOUTERBORDER;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetExternal(IDispatch** ppDispatch)
{
	*ppDispatch = static_cast<IDispatch*>(this->webBrowser2);

	return S_OK;
}

// ---------- IOleWindow ----------

HRESULT STDMETHODCALLTYPE WebBrowser::GetWindow(__RPC__deref_out_opt HWND *phwnd)
{
	(*phwnd) = hWndParent;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

// ---------- IOleInPlaceSite ----------

HRESULT STDMETHODCALLTYPE WebBrowser::CanInPlaceActivate(void)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnInPlaceActivate(void)
{
	OleLockRunning(oleObject, TRUE, FALSE);
	oleObject->QueryInterface(&oleInPlaceObject);
	oleInPlaceObject->SetObjectRects(&rObject, &rObject);

	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnUIActivate(void)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetWindowContext( 
	__RPC__deref_out_opt IOleInPlaceFrame **ppFrame,
	__RPC__deref_out_opt IOleInPlaceUIWindow **ppDoc,
	__RPC__out LPRECT lprcPosRect,
	__RPC__out LPRECT lprcClipRect,
	__RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	HWND hwnd = hWndParent;

	//*ppFrame = static_cast<IOleInPlaceFrame*>(this);
	AddRef();

	(*ppFrame) = NULL;
	(*ppDoc) = NULL;
	(*lprcPosRect).left = rObject.left;
	(*lprcPosRect).top = rObject.top;
	(*lprcPosRect).right = rObject.right;
	(*lprcPosRect).bottom = rObject.bottom;
	*lprcClipRect = *lprcPosRect;

	lpFrameInfo->fMDIApp = false;
	lpFrameInfo->hwndFrame = hwnd;
	lpFrameInfo->haccel = NULL;
	lpFrameInfo->cAccelEntries = 0;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Scroll(SIZE scrollExtant)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnUIDeactivate(BOOL fUndoable)
{
	return S_OK;
}

HWND WebBrowser::GetControlWindow()
{
	if(hWndControl != 0)
		return hWndControl;

	if(oleInPlaceObject == 0)
		return 0;

	oleInPlaceObject->GetWindow(&hWndControl);

	return hWndControl;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnInPlaceDeactivate(void)
{
	hWndControl = 0;
	oleInPlaceObject = 0;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DiscardUndoState(void)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DeactivateAndUndo(void)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnPosRectChange(LPCRECT lprcPosRect)
{
	IOleInPlaceObject* iole = NULL;
	this->webBrowser2->QueryInterface(IID_IOleInPlaceObject, (void**)&iole);

	if (iole != NULL)
	{
		iole->SetObjectRects(lprcPosRect, lprcPosRect);
		iole->Release();
	}

	return S_OK;
}

// ---------- IOleClientSite ----------

HRESULT STDMETHODCALLTYPE WebBrowser::SaveObject(void)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, __RPC__deref_out_opt IMoniker **ppmk)
{
	if((dwAssign == OLEGETMONIKER_ONLYIFTHERE) && (dwWhichMoniker == OLEWHICHMK_CONTAINER))
		return E_FAIL;

	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetContainer(__RPC__deref_out_opt IOleContainer **ppContainer)
{
	// Retrieves a pointer to the object's container
	*ppContainer = NULL;
	return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ShowObject(void)
{
	// This method is called when the browser need thr container to display its object
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnShowWindow(BOOL fShow)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::RequestNewObjectLayout(void)
{
	return E_NOTIMPL;
}

// ----- IStorage -----

HRESULT STDMETHODCALLTYPE WebBrowser::CreateStream( 
	__RPC__in_string const OLECHAR *pwcsName,
	DWORD grfMode,
	DWORD reserved1,
	DWORD reserved2,
	__RPC__deref_out_opt IStream **ppstm)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OpenStream( 
	const OLECHAR *pwcsName,
	void *reserved1,
	DWORD grfMode,
	DWORD reserved2,
	IStream **ppstm)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::CreateStorage( 
	__RPC__in_string const OLECHAR *pwcsName,
	DWORD grfMode,
	DWORD reserved1,
	DWORD reserved2,
	__RPC__deref_out_opt IStorage **ppstg)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OpenStorage( 
	__RPC__in_opt_string const OLECHAR *pwcsName,
	__RPC__in_opt IStorage *pstgPriority,
	DWORD grfMode,
	__RPC__deref_opt_in_opt SNB snbExclude,
	DWORD reserved,
	__RPC__deref_out_opt IStorage **ppstg)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::CopyTo( 
	DWORD ciidExclude,
	const IID *rgiidExclude,
	__RPC__in_opt  SNB snbExclude,
	IStorage *pstgDest)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::MoveElementTo( 
	__RPC__in_string const OLECHAR *pwcsName,
	__RPC__in_opt IStorage *pstgDest,
	__RPC__in_string const OLECHAR *pwcsNewName,
	DWORD grfFlags)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Commit(DWORD grfCommitFlags)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Revert(void)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::EnumElements( 
	DWORD reserved1,
	void *reserved2,
	DWORD reserved3,
	IEnumSTATSTG **ppenum)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DestroyElement(__RPC__in_string const OLECHAR *pwcsName)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::RenameElement( 
	__RPC__in_string const OLECHAR *pwcsOldName,
	__RPC__in_string const OLECHAR *pwcsNewName)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetElementTimes( 
	__RPC__in_opt_string const OLECHAR *pwcsName,
	__RPC__in_opt const FILETIME *pctime,
	__RPC__in_opt const FILETIME *patime,
	__RPC__in_opt const FILETIME *pmtime)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetClass(__RPC__in REFCLSID clsid)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetStateBits(DWORD grfStateBits, DWORD grfMask)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Stat(__RPC__out STATSTG *pstatstg, DWORD grfStatFlag)
{
	return E_NOTIMPL;
}

void WebBrowser::TCharToWide(const char* src, wchar_t* dst, int dst_size_in_wchars)
{
	MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, src, -1, dst, dst_size_in_wchars);
}

void WebBrowser::TCharToWide(const wchar_t* src, wchar_t* dst, int dst_size_in_wchars)
{
#pragma warning(suppress:4996)
	wcscpy(dst, src);
}

char* WebBrowser::BSTRToLPSTR(BSTR bStr, LPSTR lpstr)
{
	int lenW = SysStringLen(bStr);

	if (lenW)
	{
		int lenA = WideCharToMultiByte(CP_ACP, 0, bStr, lenW, 0, 0, NULL, NULL);

		if (lenA > 0)
		{
			lpstr = new char[lenA + 1];
			WideCharToMultiByte(CP_ACP, 0, bStr, lenW, lpstr, lenA, NULL, NULL);
			lpstr[lenA] = '\0';
		}
		else
			lpstr = NULL;
	}
	else
		lpstr = NULL;

	return lpstr;
}

JSObject::JSObject() : ref(0)
{
	idMap.insert(std::make_pair(L"execute", DISPID_USER_EXECUTE));
	idMap.insert(std::make_pair(L"writefile", DISPID_USER_WRITEFILE));
	idMap.insert(std::make_pair(L"readfile", DISPID_USER_READFILE));
	idMap.insert(std::make_pair(L"getevar", DISPID_USER_GETVAL));
	idMap.insert(std::make_pair(L"setevar", DISPID_USER_SETVAL));
}

HRESULT STDMETHODCALLTYPE JSObject::QueryInterface(REFIID riid, void** ppv)
{
	*ppv = NULL;

	if (riid == IID_IUnknown || riid == IID_IDispatch)
		*ppv = static_cast<IDispatch*>(this);

	if (*ppv != NULL)
	{
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE JSObject::AddRef()
{
	return InterlockedIncrement(&ref);
}

ULONG STDMETHODCALLTYPE JSObject::Release()
{
	int tmp = InterlockedDecrement(&ref);

	if (tmp == 0)
	{
		OutputDebugString(L"JSObject::Release(): delete this");
		delete this;
	}

	return tmp;
}

HRESULT STDMETHODCALLTYPE JSObject::GetTypeInfoCount(UINT* pctinfo)
{
	*pctinfo = 0;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE JSObject::GetTypeInfo(UINT iTInfo, LCID lcid,
	ITypeInfo** ppTInfo)
{
	return E_FAIL;
}

HRESULT STDMETHODCALLTYPE JSObject::GetIDsOfNames(REFIID riid,
	LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	HRESULT hr = S_OK;

	for (UINT i = 0; i < cNames; i++)
	{
		std::map<std::wstring, DISPID>::iterator iter = idMap.find(rgszNames[i]);
		if (iter != idMap.end())
			rgDispId[i] = iter->second;
		else
		{
			rgDispId[i] = DISPID_UNKNOWN;
			hr = DISP_E_UNKNOWNNAME;
		}
	}

	return hr;
}

HRESULT STDMETHODCALLTYPE JSObject::Invoke(DISPID dispIdMember, REFIID riid,
	LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult,
	EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	if (wFlags & DISPATCH_METHOD)
	{
		HRESULT hr = S_OK;

		std::string* args = new std::string[pDispParams->cArgs];
		for (size_t i = 0; i < pDispParams->cArgs; ++i)
		{
			BSTR bstrArg = pDispParams->rgvarg[i].bstrVal;
			LPSTR arg = NULL;
			arg = BSTRToLPSTR(bstrArg, arg);
			//Re-reverse order of arguments
			args[pDispParams->cArgs - 1 - i] = arg;
			delete[] arg;
		}

		switch (dispIdMember)
		{
		case DISPID_USER_EXECUTE:
		{
			//LSExecute(NULL, args[0].c_str(), SW_NORMAL);
			break;
		}
		case DISPID_USER_WRITEFILE:
		{
			std::ofstream outfile;
			std::ios_base::openmode mode = std::ios_base::out;

			if (args[1] == "overwrite")
				mode |= std::ios_base::trunc;
			else if (args[1] == "append")
				mode |= std::ios_base::app;

			outfile.open(args[0].c_str());
			outfile << args[2];
			outfile.close();
			break;
		}
		case DISPID_USER_READFILE:
		{
			std::string buffer;
			std::string line;
			std::ifstream infile;
			infile.open(args[0].c_str());

			while (std::getline(infile, line))
			{
				buffer += line;
				buffer += "\n";
			}

			int lenW = MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, buffer.c_str(), -1, NULL, 0);
			BSTR bstrRet = SysAllocStringLen(0, lenW - 1);
			MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, buffer.c_str(), -1, bstrRet, lenW);

			pVarResult->vt = VT_BSTR;
			pVarResult->bstrVal = bstrRet;

			break;
		}
		case DISPID_USER_GETVAL:
		{
			char* buf = new char[256];

#pragma warning( disable : 4996 )
			strncpy(buf, values[args[0]].c_str(), 256);

			int lenW = MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, buf, -1, NULL, 0);
			BSTR bstrRet = SysAllocStringLen(0, lenW - 1);
			MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, buf, -1, bstrRet, lenW);

			pVarResult->vt = VT_BSTR;
			pVarResult->bstrVal = bstrRet;

			break;
		}
		case DISPID_USER_SETVAL:
		{
			std::map<std::string, std::string>::iterator itr = values.find(args[0]);
			if (itr == values.end())
				values.insert(std::make_pair(args[0], args[1]));
			else
				values[args[0]] = args[1];

			break;
		}
		default:
			hr = DISP_E_MEMBERNOTFOUND;
		}

		delete[] args;

		return hr;
	}

	return E_FAIL;
}

char* JSObject::BSTRToLPSTR(BSTR bStr, LPSTR lpstr)
{
	int lenW = SysStringLen(bStr);
	if (lenW)
	{
		int lenA = WideCharToMultiByte(CP_ACP, 0, bStr, lenW, 0, 0, NULL, NULL);

		if (lenA > 0)
		{
			lpstr = new char[lenA + 1];
			WideCharToMultiByte(CP_ACP, 0, bStr, lenW, lpstr, lenA, NULL, NULL);
			lpstr[lenA] = '\0';
		}
		else
			lpstr = NULL;
	}
	else
		lpstr = NULL;

	return lpstr;
}
