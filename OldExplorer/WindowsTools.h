#pragma once

//-----------------------------------------------------------------------------
inline WCHAR* WCHARI(std::wstring init = L"")
{
	size_t tocopy = init.length();
	if (tocopy)
	{
		WCHAR* ret = new WCHAR[tocopy + 1];
		wcscpy_s(ret, tocopy + 1, init.c_str());
		return ret;
	}
	else
	{
		WCHAR* ret = new WCHAR;
		ret[0] = 0;
		return ret;
	}
}
template <typename M, typename F, typename... Args>
inline WCHAR* WCHARI(M maxsize, F format, Args... args)
{
	WCHAR* ret = new WCHAR[maxsize];
	ret[0] = 0;
	swprintf_s(ret, maxsize, format, args...);
	return ret;
}
inline WCHAR* WCHARI(const WCHAR* init)
{
	size_t tocopy = wcslen(init);
	if (tocopy)
	{
		WCHAR* ret = new WCHAR[tocopy + 1];
		wcscpy_s(ret, tocopy + 1, init);
		return ret;
	}
	else
	{
		WCHAR* ret = new WCHAR;
		ret[0] = 0;
		return ret;
	}
	//WCHAR* ret = (WCHAR*)&init[0];
}
inline WCHAR* WCHARI(int ressourceid, const WCHAR* init = L"")
{
	WCHAR buffer[1024];
	if (LoadStringW(GetModuleHandle(NULL), ressourceid, buffer, 1024))
		return WCHARI(buffer);
	else
		return WCHARI(init);
}

//-----------------------------------------------------------------------------
inline std::wstring rs2ws(char* s)
{
	size_t len = strlen(s) + 1;

	if (len > 1)
	{
		wchar_t* ws = new wchar_t[len];

		size_t outlen;
		mbstowcs_s(&outlen, ws, len, s, len - 1);
		return ws;
	}
	else
		return L"";
}

//-----------------------------------------------------------------------------
inline std::wstring rs2ws(std::string s)
{
	return rs2ws((char*)s.c_str());
}

//-----------------------------------------------------------------------------
inline std::string rws2s(wchar_t* ws)
{
	size_t len = wcslen(ws) + 1;

	if (len > 1)
	{
		char* s = new char[len];

		size_t outlen;
		wcstombs_s(&outlen, s, len, ws, len - 1);
		return s;
	}
	else
		return "";
}

//-----------------------------------------------------------------------------
inline std::string rws2s(std::wstring ws)
{
	return rws2s((wchar_t*)ws.c_str());
}


//-----------------------------------------------------------------------------
inline BOOL ClientArea(RECT* rect, bool points = false)	//x, y, w, h, true x, y, x', y'
{
	RECT desktop;
	GetClientRect(GetDesktopWindow(), &desktop);

	HWND taskbarhwnd = FindWindow(L"Shell_traywnd", NULL);

	if (taskbarhwnd)
	{
		RECT taskbar;
		GetWindowRect(taskbarhwnd, &taskbar);

		if (taskbar.top)							//Bottom position
		{
			rect->left = 0;
			rect->top = 0;
			rect->right = desktop.right;
			rect->bottom = taskbar.top;
		}
		else if (taskbar.left)						//Right position
		{
			rect->left = 0;
			rect->top = 0;
			rect->right = taskbar.left;
			rect->bottom = desktop.bottom;
		}
		else if (taskbar.right < desktop.right)		//Left position
		{
			rect->left = taskbar.right;
			rect->top = 0;
			if (points)
				rect->right = desktop.right;
			else
				rect->right = desktop.right - taskbar.right;
			rect->bottom = desktop.bottom;
		}
		else										//Top position
		{
			rect->left = 0;
			rect->top = taskbar.bottom;
			rect->right = desktop.right;
			if (points)
				rect->bottom = desktop.bottom;
			else
				rect->bottom = desktop.bottom - taskbar.bottom;
		}

		return TRUE;
	}
	else
		*rect = desktop;

	return FALSE;
}

//-----------------------------------------------------------------------------
inline void SetWindowTransparent(HWND hwnd, bool bTransparent, int nTransparency)
{
	int lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE);

	if (bTransparent)
	{
		HMODULE hMod = LoadLibrary(_T("User32"));
		if (hMod != nullptr)
		{
			typedef BOOL(WINAPI* pSetLayeredWindowAttributes)(HWND hwnd, COLORREF crKey, BYTE bAlpha, DWORD dwFlags);
			pSetLayeredWindowAttributes p = (pSetLayeredWindowAttributes)GetProcAddress(hMod, "SetLayeredWindowAttributes");
			if (p)
			{
				(void)SetWindowLong(hwnd, GWL_EXSTYLE, lExStyle | WS_EX_LAYERED);
				(void)p(hwnd, 0, (BYTE)(nTransparency * 2.55), LWA_ALPHA);
			}
			(void)FreeLibrary(hMod);
		}
	}
	else
	{
		(void)SetWindowLong(hwnd, GWL_EXSTYLE, lExStyle & ~WS_EX_LAYERED);
		(void)RedrawWindow(hwnd, NULL, NULL, RDW_ERASE | RDW_INVALIDATE | RDW_FRAME | RDW_ALLCHILDREN);
	}
}