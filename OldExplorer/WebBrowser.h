#include <comdef.h>
#include <mshtml.h>
#include <mshtmhst.h>
#include <map>
#include <exdisp.h>

#define WEBFN_LOADED 3
// Notification sent via WM_COMMAND when you have called WebformGo(hWebF, url)
// Indicates that the page has finished loading

using namespace std;

#pragma warning(push)
#pragma warning(disable:4584)
class WebBrowser :
	public IOleClientSite,
	public IOleInPlaceSite,
	public IStorage
{
public:
	enum BrowserNavConstants {
		navOpenInNewWindow			= 0x01u,		// Open the resource or file in a new window
		navNoHistory				= 0x02u,		// Do not add the resource or file to the history list
													// The new page replaces the current page in the list.
		navNoReadFromCache			= 0x04u,		// Not implemented
		navNoWriteToCache			= 0x08u,		// Not implemented
		navAllowAutosearch			= 0x010u,		// If the navigation fails, the autosearch functionality attempts to navigate common root domains (.com, .edu, and so on)
													// If this also fails, the URL is passed to a search engine
		navBrowserBar				= 0x020u,		// Causes the current Explorer Bar to navigate to the given item, if possible
		navHyperlink				= 0x040u,		// If the navigation fails when a hyperlink is being followed
													// This constant specifies that the resource should then be bound to the parser
		navEnforceRestricted		= 0x080u,		// Force the URL into the restricted zone
		navNewWindowsManaged		= 0x0100u,		// Use the default Popup Manager to block pop-up windows
		navUntrustedForDownload		= 0x0200u,		// Block files that normally trigger a file download dialog box
		navTrustedForActiveX		= 0x0400u,		// Prompt for the installation of Microsoft ActiveX controls
		navOpenInNewTab				= 0x0800u,		// Open the resource or file in a new tab
		navOpenInBackgroundTab		= 0x01000u,		// Open the resource or file in a new background tab
		navKeepWordWheelText		= 0x02000u,		// Maintain state for dynamic navigation based on the filter string entered in the search band text box (wordwheel)
		navVirtualTab				= 0x04000u,		// Open the resource as a replacement for the current or target tab
													// The existing tab is closed while the new tab takes its place in the tab bar and replaces it in the tab group
		navBlockRedirectsXDomain	= 0x08000u,		// Block cross-domain redirect requests
		navOpenNewForegroundTab		= 0x010000u,	// The foreground tab is replaced by the new opened tab
	};

	WebBrowser(HWND hWndParent);
	~WebBrowser();

	bool				CreateBrowser(bool showScrollbars = true);
	void				setupOle();
	RECT				PixelToHiMetric(const RECT& _rc);
	virtual void		SetRect(const RECT& _rc);

	// ----- Control methods -----
	void				GoBack();
	void				GoForward();
	void				Navigate(wstring szUrl);
	void				Refresh(bool zoomreset = false);
	void				Stop();
	void				Home();
	void				Search();
	void				Zoom(int zoomto);
	void				ZoomPlus(int zoomin = 5);
	void				ZoomMinus(int zoomout = 5);
	void				ZoomReset(int zoomto = 47);
	void				FullScreen();
	void				SetPosition(int x, int y);
	void				SetScrollPos(unsigned char bar = SB_VERT, long pos = 0);
	void				BeforeNavigate(std::string url, bool* cancel);
	void				NavigateComplete(std::string url, WebBrowser* webForm);
	void				DocumentComplete(const wchar_t* url);
	void				AddCustomObject(IDispatch* custObj, std::string name);

	IHTMLDocument2*		GetDoc();
	IWebBrowser2*		webBrowser2;

	// ----- IUnknown -----
	virtual HRESULT		STDMETHODCALLTYPE QueryInterface(REFIID riid, void**ppvObject) override;
	virtual ULONG		STDMETHODCALLTYPE AddRef(void);
	virtual ULONG		STDMETHODCALLTYPE Release(void);

	// ---------- IDispatch ----------
	HRESULT				STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

	// ---------- IDocHostUIHandler ----------
	virtual HRESULT		STDMETHODCALLTYPE GetHostInfo(DOCHOSTUIINFO* pInfo);
	virtual HRESULT		STDMETHODCALLTYPE GetExternal(IDispatch** ppDispatch);

	// ---------- IOleWindow ----------
	virtual HRESULT		STDMETHODCALLTYPE GetWindow(__RPC__deref_out_opt HWND *phwnd) override;
	virtual HRESULT		STDMETHODCALLTYPE ContextSensitiveHelp(BOOL fEnterMode) override;

	// ---------- IOleInPlaceSite ----------
	virtual HRESULT		STDMETHODCALLTYPE CanInPlaceActivate(void) override;
	virtual HRESULT		STDMETHODCALLTYPE OnInPlaceActivate(void) override;
	virtual HRESULT		STDMETHODCALLTYPE OnUIActivate(void) override;
	virtual HRESULT		STDMETHODCALLTYPE GetWindowContext(__RPC__deref_out_opt IOleInPlaceFrame **ppFrame, __RPC__deref_out_opt IOleInPlaceUIWindow **ppDoc, __RPC__out LPRECT lprcPosRect, __RPC__out LPRECT lprcClipRect, __RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo) override;
	virtual HRESULT		STDMETHODCALLTYPE Scroll(SIZE scrollExtant) override;
	virtual HRESULT		STDMETHODCALLTYPE OnUIDeactivate(BOOL fUndoable) override;
	virtual HWND		GetControlWindow();
	virtual HRESULT		STDMETHODCALLTYPE OnInPlaceDeactivate(void) override;
	virtual HRESULT		STDMETHODCALLTYPE DiscardUndoState(void) override;
	virtual HRESULT		STDMETHODCALLTYPE DeactivateAndUndo(void) override;
	virtual HRESULT		STDMETHODCALLTYPE OnPosRectChange(LPCRECT lprcPosRect) override;

	// ---------- IOleClientSite ----------
	virtual HRESULT		STDMETHODCALLTYPE SaveObject(void) override;
	virtual HRESULT		STDMETHODCALLTYPE GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, __RPC__deref_out_opt IMoniker **ppmk) override;
	virtual HRESULT		STDMETHODCALLTYPE GetContainer(__RPC__deref_out_opt IOleContainer **ppContainer) override;
	virtual HRESULT		STDMETHODCALLTYPE ShowObject( void) override;
	virtual HRESULT		STDMETHODCALLTYPE OnShowWindow(BOOL fShow) override;
	virtual HRESULT		STDMETHODCALLTYPE RequestNewObjectLayout( void) override;

	// ----- IStorage -----
	virtual HRESULT		STDMETHODCALLTYPE CreateStream(__RPC__in_string const OLECHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, __RPC__deref_out_opt IStream **ppstm) override;
	virtual HRESULT		STDMETHODCALLTYPE OpenStream(const OLECHAR *pwcsName, void *reserved1, DWORD grfMode, DWORD reserved2, IStream **ppstm) override;
	virtual HRESULT		STDMETHODCALLTYPE CreateStorage(__RPC__in_string const OLECHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, __RPC__deref_out_opt IStorage **ppstg) override;
	virtual HRESULT		STDMETHODCALLTYPE OpenStorage(__RPC__in_opt_string const OLECHAR *pwcsName, __RPC__in_opt IStorage *pstgPriority, DWORD grfMode, __RPC__deref_opt_in_opt SNB snbExclude, DWORD reserved, __RPC__deref_out_opt IStorage **ppstg) override;
	virtual HRESULT		STDMETHODCALLTYPE CopyTo(DWORD ciidExclude, const IID *rgiidExclude, __RPC__in_opt  SNB snbExclude, IStorage *pstgDest) override;
	virtual HRESULT		STDMETHODCALLTYPE MoveElementTo(__RPC__in_string const OLECHAR *pwcsName, __RPC__in_opt IStorage *pstgDest, __RPC__in_string const OLECHAR *pwcsNewName, DWORD grfFlags) override;
	virtual HRESULT		STDMETHODCALLTYPE Commit(DWORD grfCommitFlags) override;
	virtual HRESULT		STDMETHODCALLTYPE Revert( void) override;
	virtual HRESULT		STDMETHODCALLTYPE EnumElements(DWORD reserved1, void *reserved2, DWORD reserved3, IEnumSTATSTG **ppenum) override;
	virtual HRESULT		STDMETHODCALLTYPE DestroyElement(__RPC__in_string const OLECHAR *pwcsName) override;
	virtual HRESULT		STDMETHODCALLTYPE RenameElement(__RPC__in_string const OLECHAR *pwcsOldName, __RPC__in_string const OLECHAR *pwcsNewName) override;
	virtual HRESULT		STDMETHODCALLTYPE SetElementTimes(__RPC__in_opt_string const OLECHAR *pwcsName, __RPC__in_opt const FILETIME *pctime, __RPC__in_opt const FILETIME *patime, __RPC__in_opt const FILETIME *pmtime) override;
	virtual HRESULT		STDMETHODCALLTYPE SetClass(__RPC__in REFCLSID clsid) override;
	virtual HRESULT		STDMETHODCALLTYPE SetStateBits(DWORD grfStateBits, DWORD grfMask) override;
	virtual HRESULT		STDMETHODCALLTYPE Stat(__RPC__out STATSTG *pstatstg, DWORD grfStatFlag) override;

	void				TCharToWide(const char* src, wchar_t* dst, int dst_size_in_wchars);
	void				TCharToWide(const wchar_t* src, wchar_t* dst, int dst_size_in_wchars);
	char*				BSTRToLPSTR(BSTR bStr, LPSTR lpstr);

private:
	int					zoom = 47;
	bool				fullscreen = true;
	bool				hasScrollbars;			// Read from WS_VSCROLL|WS_HSCROLL at WM_CREATE
	unsigned int		isnaving;				// bitmask:
												//1: haven't yet received BeforeNavigate
												//2: haven't yet received DocumentComplete
												//4: haven't yet finished Navigate call
	DWORD				cookie;

protected:
	HWND hWndParent;
	HWND hWndControl;

	IOleInPlaceObject*	oleInPlaceObject;
	IOleObject*			oleObject;
	LONG				iComRefCount;
	RECT				rObject;
};

class JSObject : public IDispatch
{
private:
	std::map<std::wstring, DISPID> idMap;
	std::map<std::string, std::string> values;
	static const		DISPID DISPID_USER_EXECUTE = DISPID_VALUE + 1;
	static const		DISPID DISPID_USER_WRITEFILE = DISPID_VALUE + 2;
	static const		DISPID DISPID_USER_READFILE = DISPID_VALUE + 3;
	static const		DISPID DISPID_USER_GETVAL = DISPID_VALUE + 4;
	static const		DISPID DISPID_USER_SETVAL = DISPID_VALUE + 5;
	long				ref;
	char*				BSTRToLPSTR(BSTR bStr, LPSTR lpstr);

public:
	JSObject();

	// IUnknown
	virtual HRESULT		STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv);
	virtual ULONG		STDMETHODCALLTYPE AddRef();
	virtual ULONG		STDMETHODCALLTYPE Release();

	// IDispatch
	virtual HRESULT		STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo);
	virtual HRESULT		STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
	virtual HRESULT		STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
	virtual HRESULT		STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);
};
