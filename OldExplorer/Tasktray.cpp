#include "stdafx.h"
#include "Tasktray.h"

Tasktray::Tasktray()
{
}

Tasktray::~Tasktray()
{
	DestroyMenu(m_menu);
}

void Tasktray::Init(HWND hWnd)
{
	m_hWnd = hWnd;
	m_nid.cbSize = sizeof(NOTIFYICONDATA);
	m_nid.hWnd = hWnd;
	m_nid.uID = 1;
	m_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	m_nid.uCallbackMessage = WM_TASKTRAY;
	m_nid.hIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_OLDEXPLORER_ICON32));
	swprintf_s(m_nid.szTip, 128, L"OldExplorer");
	CreateMenu();
}

void Tasktray::CreateMenu()
{
	DestroyMenu(m_menu);
	m_menu = CreatePopupMenu();
	AppendMenu(m_menu, MF_POPUP, IDM_MENU_ABOUT, L"About");
	AppendMenu(m_menu, MF_SEPARATOR, NULL, NULL);
	AppendMenu(m_menu, MF_POPUP, IDM_MENU_EXIT, L"Exit");
}

void Tasktray::Show()
{
	if (!m_flag)
		Shell_NotifyIcon(NIM_ADD, &m_nid);
	m_flag = true;
}

void Tasktray::Hide()
{
	if (m_flag)
		Shell_NotifyIcon(NIM_DELETE, &m_nid);
	m_flag = false;
}

void Tasktray::Tip(WCHAR* buf)
{
	if (!m_flag)
		return;
	wcscpy_s(m_nid.szTip, wcslen(buf) + 1, buf);
	Shell_NotifyIcon(NIM_MODIFY, &m_nid);
}

void Tasktray::Message(WPARAM wParam, LPARAM lParam)
{
	if (wParam == 1)
	{
		switch (lParam)
		{
		case WM_LBUTTONDBLCLK:
		{
			ShowWindow(m_hWnd, SW_SHOWNORMAL);
			SetForegroundWindow(m_hWnd);
			PostMessage(m_hWnd, WM_SYSCOMMAND, SC_RESTORE, 0);
			break;
		}
		case WM_RBUTTONDOWN:
		{
			SetForegroundWindow(m_hWnd);
			struct MenuPosition mp;
			GetMenuPosition(&mp);
			TrackPopupMenu(m_menu, mp.flags, mp.x, mp.y, 0, m_hWnd, NULL);
			break;
		}
		}
	}
}

	UINT Tasktray::GetTaskBarLocation()
	{
		RECT rectangle;

		return GetTaskBarLocation(&rectangle);
	}

	UINT Tasktray::GetTaskBarLocation(RECT* rectangle)
	{
		UINT location = ABE_BOTTOM;
		APPBARDATA pData;
		pData.cbSize = sizeof(APPBARDATA);
		pData.hWnd = m_hWnd;

		if (SHAppBarMessage(ABM_GETTASKBARPOS, &pData))
		{
			*rectangle = pData.rc;
			return pData.uEdge;
		}

		return location;
	}

	void Tasktray::GetMenuPosition(struct MenuPosition* menuposition)
	{
		POINT p;
		if (!GetCursorPos(&p))
		{
			p.x = GetSystemMetrics(SM_CXSCREEN) / 2;
			p.y = GetSystemMetrics(SM_CYSCREEN) / 2;
		}

		switch (GetTaskBarLocation())
		{
		case ABE_BOTTOM: { *menuposition = { TPM_CENTERALIGN | TPM_BOTTOMALIGN, p.x, p.y - 10 }; break; }
		case ABE_LEFT: { *menuposition = { TPM_LEFTALIGN | TPM_BOTTOMALIGN, p.x + 10, p.y }; break; }
		case ABE_RIGHT: { *menuposition = { TPM_RIGHTALIGN | TPM_BOTTOMALIGN, p.x - 10, p.y }; break; }
		case ABE_TOP: { *menuposition = { TPM_CENTERALIGN | TPM_TOPALIGN, p.x, p.y + 10 }; break; }
		}
	}
