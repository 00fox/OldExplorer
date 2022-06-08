#include "stdafx.h"
#include "OldExplorer.h"
#include "Explorer.h"

Explorer::Explorer()
{
	HKEY hkey = NULL;
	long openStatus = RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION"), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, &hkey);
	if (openStatus == ERROR_SUCCESS)
	{
		DWORD navigatorver = 0x2ee1;
		long status = RegSetValueEx(hkey, L"OldExplorer.exe", 0, REG_DWORD, (const BYTE*)&navigatorver, sizeof(navigatorver));
		RegCloseKey(hkey);
	}
}

Explorer::~Explorer()
{
	HKEY hkey = NULL;
	long openStatus = RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION"), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, &hkey);
	if (openStatus == ERROR_SUCCESS)
	{
		long status = RegDeleteValue(hkey, L"OldExplorer.exe");
		RegCloseKey(hkey);
	}
	delete explorer;
	explorer = NULL;
}

void Explorer::Init(HWND hWnd)
{
	m_hDlg = CreateDialogParam(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_EXPLORER), hWnd, (DLGPROC)Proc, (LPARAM)this);
	m_hWnd = hWnd;

	explorer = new WebBrowser(m_hDlg);
	SetTimer(m_hDlg, 8, 500, NULL);

	SendDlgItemMessage(m_hDlg, IDC_EDIT_URL, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_GO, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_BACK, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_REFRESH, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FORWARD, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_AUTOREFRESH, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_ZOOM_PLUS, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_ZOOM_MINUS, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_0, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_1, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_2, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_3, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_4, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_5, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_6, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_7, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_8, WM_SETFONT, WPARAM(hLegend2), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_FAV_9, WM_SETFONT, WPARAM(hLegend3), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_TOP, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_BOTTOM, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_LEFT, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));
	SendDlgItemMessage(m_hDlg, IDC_BUTTON_RIGHT, WM_SETFONT, WPARAM(hLegend), MAKELPARAM(TRUE, 0));

	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_GO), L"»");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_BACK), L"‹");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_REFRESH), L"⭯");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FORWARD), L"›");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_AUTOREFRESH), L"□");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_ZOOM_PLUS), L"⁺");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_ZOOM_MINUS), L"⁻");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_0), L"¹");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_1), L"²");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_2), L"³");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_3), L"⁴");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_4), L"⁵");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_5), L"⁶");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_6), L"⁷");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_7), L"⁸");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_8), L"⁹");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_FAV_9), L"¹⁰");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_TOP), L"⭡");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_BOTTOM), L"⭣");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_LEFT), L"⭠");
	SetWindowText(GetDlgItem(m_hDlg, IDC_BUTTON_RIGHT), L"⭢");
	
	::SetWindowTheme(GetDlgItem(m_hDlg, IDC_BUTTON_AUTOREFRESH), L"", L"");

	Hide();
}

INT_PTR CALLBACK Explorer::Proc(HWND hWnd, UINT uMsg, WPARAM wparam, LPARAM lparam)
{
	Explorer* dlg;

	if (uMsg == WM_INITDIALOG)
	{
		dlg = reinterpret_cast<Explorer*>(lparam);
		SetWindowLongPtrW(hWnd, DWLP_USER, lparam);
	}
	else
		dlg = reinterpret_cast<Explorer*>(GetWindowLongPtrW(hWnd, DWLP_USER));
	if (dlg)
	{
		INT_PTR result;
		result = dlg->_proc(hWnd, uMsg, wparam, lparam);
		return result;
	}
	return DefWindowProcW(hWnd, uMsg, wparam, lparam);
}

INT_PTR Explorer::_proc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_PAINT:
	{
		if (!IsIconic(hWnd))
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(hWnd, &ps);
			RECT rect;
			GetClientRect(hWnd, &rect);
			HBRUSH hB_web = CreateSolidBrush(RGB(255, 255, 255));
			FillRect(hDC, &rect, hB_web);
			DeleteObject(hB_web);
			::ReleaseDC(hWnd, hDC);
			EndPaint(hWnd, &ps);
		}
		return FALSE;
	}
	case WM_NOTIFY:
	{
		switch (((LPNMLISTVIEW)lParam)->hdr.code)
		{
		case BCN_HOTITEMCHANGE:
		{
			switch (((NMBCHOTITEM*)lParam)->dwFlags)
				// HICF_OTHER			= 0x00000000
				// HICF_MOUSE			= 0x00000001	Triggered by mouse
				// HICF_ARROWKEYS		= 0x00000002	Triggered by arrow keys
				// HICF_ACCELERATOR		= 0x00000004	Triggered by accelerator
				// HICF_DUPACCEL		= 0x00000008	This accelerator is not unique
				// HICF_ENTERING		= 0x00000010	idOld is invalid
				// HICF_LEAVING			= 0x00000020	idNew is invalid
				// HICF_RESELECT		= 0x00000040	hot item reselected
				// HICF_LMOUSE			= 0x00000080	left mouse button selected
				// HICF_TOGGLEDROPDOWN	= 0x00000100	Toggle button's dropdown state
			{
			case (HICF_ENTERING | HICF_MOUSE):
			{
				switch (((LPNMHDR)lParam)->idFrom)
				{
				case IDC_BUTTON_TOP:SetTimer(m_hDlg, 2, 25, NULL); break;
				case IDC_BUTTON_BOTTOM:SetTimer(m_hDlg, 3, 25, NULL); break;
				case IDC_BUTTON_LEFT:SetTimer(m_hDlg, 4, 25, NULL); break;
				case IDC_BUTTON_RIGHT:SetTimer(m_hDlg, 5, 25, NULL); break;
				case IDC_BUTTON_ZOOM_PLUS:SetTimer(m_hDlg, 6, 70, NULL); break;
				case IDC_BUTTON_ZOOM_MINUS:SetTimer(m_hDlg, 7, 70, NULL); break;
				default:
					return FALSE;
				}
				break;
			}
			case (HICF_LEAVING | HICF_MOUSE):
			{
				switch (((LPNMHDR)lParam)->idFrom)
				{
				case IDC_BUTTON_TOP:KillTimer(m_hDlg, 2); break;
				case IDC_BUTTON_BOTTOM:KillTimer(m_hDlg, 3); break;
				case IDC_BUTTON_LEFT:KillTimer(m_hDlg, 4); break;
				case IDC_BUTTON_RIGHT:KillTimer(m_hDlg, 5); break;
				case IDC_BUTTON_ZOOM_PLUS:KillTimer(m_hDlg, 6); break;
				case IDC_BUTTON_ZOOM_MINUS:KillTimer(m_hDlg, 7); break;
				default:
					return FALSE;
				}
				break;
			}
			default:
			{
				return FALSE;
			}
			}
			break;
		}
		}
		break;
	}
	case WM_TIMER:
	{
		if (IsIconic(hWnd) || !IsWindowVisible(m_hDlg))
			break;

		switch (wParam)
		{
		case 1:
		{
			explorer->Refresh();
			break;
		}
		case 2:
		{
			if (GetAsyncKeyState(VK_LBUTTON) & 0x8000)
			{
				posy += 10;
				explorer->SetPosition(posx, posy);
			}
			break;
		}
		case 3:
		{
			if (GetAsyncKeyState(VK_LBUTTON) & 0x8000)
			{
				posy -= 10;
				explorer->SetPosition(posx, posy);
			}
			break;
		}
		case 4:
		{
			if (GetAsyncKeyState(VK_LBUTTON) & 0x8000)
			{
				posx += 30;
				explorer->SetPosition(posx, posy);
			}
			break;
		}
		case 5:
		{
			if (GetAsyncKeyState(VK_LBUTTON) & 0x8000)
			{
				posx -= 30;
				explorer->SetPosition(posx, posy);
			}
			break;
		}
		case 6:
		{
			if (GetAsyncKeyState(VK_LBUTTON) & 0x8000)
			{
				explorer->ZoomPlus(1);
			}
			else if (GetAsyncKeyState(VK_RBUTTON) & 0x8000)
			{
				explorer->ZoomReset(Defaultzoom);
			}
			break;
		}
		case 7:
		{
			if (GetAsyncKeyState(VK_LBUTTON) & 0x8000)
			{
				explorer->ZoomMinus(1);
			}
			else if (GetAsyncKeyState(VK_RBUTTON) & 0x8000)
			{
				explorer->ZoomReset(Defaultzoom);
			}
			break;
		}
		case 8:
		{
			GetCursorPos(&mousepoint);
			RECT win;
			GetWindowRect(hWnd, &win);
			if (mousepoint.x > win.left && mousepoint.x < win.right && mousepoint.y > win.top && mousepoint.y < win.bottom)
			{
				if (!mousein)
				{
					mousein = true;
					ShowWindow(GetDlgItem(hWnd, IDC_EDIT_URL), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_GO), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_BACK), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_REFRESH), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FORWARD), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_AUTOREFRESH), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_ZOOM_PLUS), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_ZOOM_MINUS), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_0), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_1), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_2), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_3), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_4), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_5), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_6), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_7), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_8), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_9), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_TOP), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_BOTTOM), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_LEFT), SW_SHOW);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_RIGHT), SW_SHOW);
				}
			}
			else
			{
				if (mousein)
				{
					mousein = false;
					ShowWindow(GetDlgItem(hWnd, IDC_EDIT_URL), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_GO), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_BACK), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_REFRESH), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FORWARD), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_AUTOREFRESH), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_ZOOM_PLUS), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_ZOOM_MINUS), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_0), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_1), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_2), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_3), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_4), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_5), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_6), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_7), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_8), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_9), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_TOP), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_BOTTOM), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_LEFT), SW_HIDE);
					ShowWindow(GetDlgItem(hWnd, IDC_BUTTON_RIGHT), SW_HIDE);
				}
			}
			break;
		}
		case 9:
		{
			KillTimer(hWnd, 9);
			explorer->ZoomReset(Defaultzoom);
			break;
		}
		}
		break;
	}
	case WM_SHOWWINDOW:
	{
		if (wParam == TRUE)
		{
			if (!firsttime)
			{
				BSTR bstrURL = SysAllocString(WebURL0);
				if (sizeof(bstrURL))
					explorer->Navigate(bstrURL);
				SetTimer(hWnd, 9, 250, NULL);
				firsttime = true;
			}

			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_0), (wcslen(WebURL0)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_1), (wcslen(WebURL1)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_2), (wcslen(WebURL2)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_3), (wcslen(WebURL3)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_4), (wcslen(WebURL4)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_5), (wcslen(WebURL5)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_6), (wcslen(WebURL6)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_7), (wcslen(WebURL7)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_8), (wcslen(WebURL8)) ? TRUE : FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_BUTTON_FAV_9), (wcslen(WebURL9)) ? TRUE : FALSE);
		}
		break;
	}
	case WM_INITDIALOG:
	{
		return FALSE;
	}
	case WM_COMMAND:
	{
		switch (LOWORD(wParam))
		{
			case IDC_BUTTON_BACK:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					explorer->GoBack();
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_REFRESH:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					explorer->Refresh();
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_AUTOREFRESH:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					autorefresh = !autorefresh;
					if (autorefresh)
					{
						explorer->Refresh();
						SetTimer(m_hDlg, 1, 1000 * WebRefreshTime, NULL);
					}
					else
					{
						KillTimer(m_hDlg, 1);
					}
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FORWARD:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					explorer->GoForward();
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_GO:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					TCHAR* szUrl = new TCHAR[1024];
					GetWindowText(GetDlgItem(hWnd, IDC_EDIT_URL), szUrl, 1024);
					explorer->Navigate(szUrl);
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_STOP:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					explorer->Stop();
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_HOME:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					explorer->Home();
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_SEARCH:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					explorer->Search();
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_ZOOM_RESET:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					explorer->ZoomReset();
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_0:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL0);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 0;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_1:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL1);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 1;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_2:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL2);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 2;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_3:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL3);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 3;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_4:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL4);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 4;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_5:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL5);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 5;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_6:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL6);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 6;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_7:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL7);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 7;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_8:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL8);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 8;
					break;
				}
				}
				break;
			}
			case IDC_BUTTON_FAV_9:
			{
				switch (HIWORD(wParam))
				{
				case BN_CLICKED:
				{
					BSTR bstrURL = SysAllocString(WebURL9);
					explorer->Navigate(bstrURL);
					explorer->ZoomReset();
					LastLinkUsed = 9;
					break;
				}
				}
				break;
			}
		}
		break;
	}
	default:
		return FALSE;
	}
	return TRUE;
}

void Explorer::Show()
{
	ShowWindow(m_hDlg, SW_SHOW);
}

void Explorer::Hide()
{
	ShowWindow(m_hDlg, SW_HIDE);
}

BOOL Explorer::MoveWindow(int x, int y, int w, int h, BOOL r)
{
	RECT rect;
	rect.left = x-3;
	rect.top = y-21;
	//rect.right = x+w-5;		//Scrollbars
	//rect.bottom = y+h-23;
	rect.right = x + w + 11;	//No Scrollbar
	rect.bottom = y + h - 7;
	explorer->SetRect(rect);
	return ::MoveWindow(m_hDlg, x, y, w, h, r);
}
