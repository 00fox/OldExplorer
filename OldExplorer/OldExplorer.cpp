#include "stdafx.h"
#include "OldExplorer.h"
#include "Explorer.h"
#include "Tasktray.h"

static HWND			ProghWnd;
static WCHAR		szTitle[100];
static WCHAR		szWindowClass[100];

	   ATOM             RegisterWndClass(HINSTANCE hInstance);
	   BOOL             InitInstance(HINSTANCE, int);
static int              RunMessagePump();
static LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
	   INT_PTR CALLBACK About(HWND, UINT, WPARAM, LPARAM);

DPI_AWARENESS_CONTEXT dpiAwarenessContext = DPI_AWARENESS_CONTEXT_UNAWARE;

int APIENTRY wWinMain(_In_     HINSTANCE hInstance,
					  _In_opt_ HINSTANCE hPrevInstance,
					  _In_     LPWSTR    lpCmdLine,
					  _In_     int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

    // Test if this code runs on the correct machine
	SYSTEM_INFO info;
	GetNativeSystemInfo(&info);
	if (info.wProcessorArchitecture != PROCESSOR_ARCHITECTURE_AMD64)	//PROCESSOR_ARCHITECTURE_AMD64
	{
		MessageBox(NULL, WCHARI(IDS_FOR_X64), WCHARI(IDS_APP_TITLE), MB_OK | MB_ICONERROR);
		return -1;
	};

	LoadStringW(hInstance, IDS_OLDEXPLORER, szWindowClass, 100);
	LoadStringW(hInstance, IDS_APP_TITLE, szTitle, 100);
	SetCurrentProcessExplicitAppUserModelID(szTitle);
	RegisterWndClass(hInstance);

	if (!InitInstance (hInstance, nCmdShow))
		return FALSE;

	if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE | COINIT_SPEED_OVER_MEMORY)))
		return FALSE;

	SetProcessDpiAwarenessContext(dpiAwarenessContext);
	InitCommonControls();

	int retVal = RunMessagePump();

	CoUninitialize();

	return retVal;
}
ATOM RegisterWndClass(HINSTANCE hInstance)
{
	WNDCLASSEXW wcex;

	wcex.cbSize			= sizeof(WNDCLASSEX);
	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_OLDEXPLORER_ICON256));
	wcex.hCursor		= LoadCursor(nullptr, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW + 1);
	wcex.lpszMenuName	= 0;
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_OLDEXPLORER_ICON32));

	return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	RECT desk;
	ClientArea(&desk);
	DWORD dwStyle = WS_VISIBLE | WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX/* | WS_MAXIMIZEBOX | WS_THICKFRAME*/;
	ProghWnd = CreateWindowExW(WS_EX_CONTROLPARENT, szWindowClass, szTitle, dwStyle, (desk.right - desk.left) / 2 - 246, desk.top + 200, 492, 327, nullptr, nullptr, hInstance, nullptr);

	if (!ProghWnd)
		return FALSE;

   return TRUE;
}

// Run the message pump for one thread.
static int RunMessagePump()
{
	HACCEL hAccelTable = LoadAccelerators(GetModuleHandle(NULL), MAKEINTRESOURCE(IDR_OLDEXPLORER));
	MSG msg;

	while (GetMessage(&msg, nullptr, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			if (!IsDialogMessage(GetAncestor(msg.hwnd, GA_ROOT), &msg))
			{
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
	}

	return (int)msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	static bool		first_WM_CREATE = false;
	static bool		done_WM_CREATE = false;
	static Tasktray	tasktray;
	static Explorer	explorerDlg;
	static POINT	mousepoint;
	static short	W;							//Logical (scaling)
	static short	H;
	static short	Wp;							//Physical
	static short	Hp;
	static short	w;							//Logical / 2
	static short	h;
	static double	Hscale;						//Scaling
	static double	Vscale;
	static bool		Transparency = false;
	static BYTE		Opacity = 60;
	static bool		TopMost = true;
	static bool		FirstMove = false;
	static short	x = 0;
	static short	y = 0;
	static RECT		FirstMoveRect = { (0, 0, 0, 0) };
	static unsigned long long m_flag_drag = 0;

	switch (message)
	{
	case WM_CREATE:
	{
		if (first_WM_CREATE)
			break;

		first_WM_CREATE = true;

		ProghWnd = hWnd;
		GetCursorPos(&mousepoint);
		tasktray.Init(hWnd);
		tasktray.Show();
		SendMessage(hWnd, WM_TRANSPARENCY, 0, 0);
		SendMessage(hWnd, WM_DISPLAYCHANGE, 0, 0);
		::SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_DEFERERASE | SWP_NOMOVE | SWP_NOSIZE);
		
		explorerDlg.Init(hWnd);
		explorerDlg.Show();

		SetTimer(hWnd, 1, 10, NULL);	//Moving window

		done_WM_CREATE = true;
		break;
	}
	case WM_TIMER:
	{
		if (wParam == 1)
		{
			if (!GetCursorPos(&mousepoint))
				break;

			if (m_flag_drag)
			{
				if (!(GetAsyncKeyState(VK_LBUTTON) & 0x8000))
				{
					PostMessage(hWnd, WM_NCLBUTTONUP, 0, 0);
				}
				else
				{
					m_flag_drag++;
					if (m_flag_drag == 4)
					{
						FirstMove = true;
						SendMessage(hWnd, WM_MOVE, 0, 0);
					}
					if (m_flag_drag > 3 && !FirstMove)
					{
						::SetWindowPos(hWnd, HWND_NOTOPMOST, mousepoint.x - x, mousepoint.y - y, 492, 327, SWP_NOOWNERZORDER | SWP_NOZORDER | SWP_DEFERERASE);
					}
				}
			}
		}
		break;
	}
	case WM_DPICHANGED:
	case WM_DISPLAYCHANGE:
	{
		auto activeWindow = GetActiveWindow();
		HMONITOR monitor = MonitorFromWindow(activeWindow, MONITOR_DEFAULTTONEAREST);

		// Get the logical width and height of the monitor
		MONITORINFOEX monitorInfoEx;
		monitorInfoEx.cbSize = sizeof(monitorInfoEx);
		GetMonitorInfo(monitor, &monitorInfoEx);
		W = monitorInfoEx.rcMonitor.right - monitorInfoEx.rcMonitor.left;
		H = monitorInfoEx.rcMonitor.bottom - monitorInfoEx.rcMonitor.top;
		w = (short)(W / 2);
		h = (short)(H / 2);

		// Get the physical width and height of the monitor
		DEVMODE devMode;
		devMode.dmSize = sizeof(devMode);
		devMode.dmDriverExtra = 0;
		EnumDisplaySettings(monitorInfoEx.szDevice, ENUM_CURRENT_SETTINGS, &devMode);
		Wp = devMode.dmPelsWidth;
		Hp = devMode.dmPelsHeight;

		// Calculate the scaling factor
		Hscale = double(Wp) / double(W);
		Vscale = double(Hp) / double(H);
		
		SendMessage(hWnd, WM_SIZE, 0, 0);
		break;
	}
	case WM_NCRBUTTONDOWN:
	case WM_TRANSPARENCY:
	{
		if (lParam)
			Transparency = !Transparency;
		if (Transparency)
			SetWindowTransparent(hWnd, true, Opacity);
		else
			SetWindowTransparent(hWnd, false, NULL);
		break;
	}
	case WM_NCLBUTTONDOWN:
	{
		if (!done_WM_CREATE)
			break;

		GetCursorPos(&mousepoint);
		GetWindowRect(hWnd, &FirstMoveRect);
		x = (short)(mousepoint.x - FirstMoveRect.left);
		y = (short)(mousepoint.y - FirstMoveRect.top);
		m_flag_drag = true;
		break;
	}
	case WM_NCLBUTTONUP:
	{
		if (!done_WM_CREATE)
			break;

		short m_flag_drag_tmp = m_flag_drag;
		m_flag_drag = false;
		GetCursorPos(&mousepoint);
		RECT win;
		GetWindowRect(hWnd, &win);
		x = (short)(mousepoint.x - win.left);
		y = (short)(mousepoint.y - win.top);
		bool minimize = false;
		bool close = false;
		if (m_flag_drag_tmp < 20)
			if (y >= 0 && y <= 30)
			{
				short z = x;
				if (z > 345 && z < 392)
					minimize = true;
				else if (z > 438 && z < 485)
					close = true;
			}
		if (minimize)
		{
			if (m_flag_drag_tmp > 3)
				::SetWindowPos(hWnd, HWND_TOPMOST, FirstMoveRect.left, FirstMoveRect.top, 0, 0, SWP_NOOWNERZORDER | SWP_NOSIZE | SWP_NOZORDER | SWP_DEFERERASE);
			PostMessage(hWnd, WM_SYSCOMMAND, SC_MINIMIZE, 0); break;
		}
		else if (close)
		{
			if (m_flag_drag_tmp > 3)
				::SetWindowPos(hWnd, HWND_TOPMOST, FirstMoveRect.left, FirstMoveRect.top, 0, 0, SWP_NOOWNERZORDER | SWP_NOSIZE | SWP_NOZORDER | SWP_DEFERERASE);
			PostMessage(hWnd, WM_SYSCOMMAND, SC_CLOSE, 0); break;
		}
		if (m_flag_drag_tmp > 3 && !FirstMove)
			PostMessage(hWnd, WM_SIZE, 0, 0);
		break;
	}
	case WM_MOVE:
	{
		if (!done_WM_CREATE)
			break;

		if (FirstMove)
		{
			RECT win;
			GetWindowRect(hWnd, &win);
			SetFocus(hWnd);
			PostMessage(hWnd, WM_SIZE, 0, 0);
			::SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
			::SetWindowPos(hWnd, HWND_TOPMOST, win.left, win.top, 0, 0, SWP_NOOWNERZORDER | SWP_NOSIZE | SWP_NOZORDER | SWP_DEFERERASE);
		}
		FirstMove = false;
		PostMessage(hWnd, WM_SIZE, 0, 0);
		break;
	}
	case WM_EXITSIZEMOVE:
	{
		PostMessage(hWnd, WM_SIZE, 0, 0);
		break;
	}
	case WM_SIZE:
	{
		RECT rect;
		GetClientRect(hWnd, &rect);

		if (done_WM_CREATE)
			explorerDlg.MoveWindow(rect.left, rect.top, rect.right, rect.bottom, FALSE);

		InvalidateRect(hWnd, NULL, TRUE);
		break;
	}
	case WM_COMMAND:
	{
		switch (LOWORD(wParam))
		{
		case IDCANCEL: //Trap VK_ESCAPE
		{
			break;
		}
		case IDM_MENU_EXIT:
		{
			PostMessage(hWnd, WM_DESTROY, 0, 0);
			break;
		}
		case IDM_MENU_ABOUT:
		{
			DialogBox(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		}
		default:
		{
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		}
		break;
	}
	case WM_TASKTRAY:
	{
		tasktray.Message(wParam, lParam);
		break;
	}
	case WM_SYSCOMMAND:
	{
		switch (wParam)
		{
		case SC_CLOSE: { PostMessage(hWnd, WM_DESTROY, 0, 0); break; }
		case SC_MINIMIZE:
		{
			if (done_WM_CREATE)
			{
				::SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
				TopMost = false;
				ShowWindow(hWnd, SW_HIDE);
			}
			return FALSE;
		}
		case SC_MAXIMIZE:
		{
			break;
		}
		case SC_RESTORE:
		{
			break;
		}
		}
		break;
	}
	case WM_CLOSE:
	{
		if (GetKeyState(VK_ESCAPE) < 0 || GetKeyState(VK_CANCEL) < 0)
			return 0;
		else
			return DefWindowProc(hWnd, message, wParam, lParam);
	}
	case WM_DESTROY:
	{
		static bool DestructionLaunched = false;
		if (done_WM_CREATE && !DestructionLaunched)
		{
			DestructionLaunched = true;
			KillTimer(hWnd, 1);
			DeleteObject(hAbout);
			DeleteObject(hLegend);
			DeleteObject(hLegend2);
			DeleteObject(hLegend3);
			tasktray.Hide();
			PostQuitMessage(0);
		}
		break;
	}
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

INT_PTR CALLBACK About(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
	{
		RECT desktop;
		ClientArea(&desktop);
		RECT dialog;
		GetClientRect(hWnd, &dialog);
		SetWindowPos(hWnd, HWND_TOP,
			desktop.left + ((desktop.right - desktop.left - dialog.right) / 2),
			desktop.top - 55 + ((desktop.bottom - desktop.top - dialog.bottom) / 2),
			0, 0, SWP_NOSIZE);
		SendDlgItemMessage(hWnd, IDC_ABOUT_1, WM_SETFONT, WPARAM(hAbout), MAKELPARAM(TRUE, 0));
		SendDlgItemMessage(hWnd, IDC_ABOUT_1, WM_SETFONT, WPARAM(hAbout), MAKELPARAM(TRUE, 0));
		SetWindowText(GetDlgItem(hWnd, IDC_ABOUT_1), L"OldExplorer");
		SetWindowText(GetDlgItem(hWnd, IDC_ABOUT_OK), L"Ok");
		return (INT_PTR)TRUE;
	}
	case WM_COMMAND:
	{
		if (LOWORD(wParam) == IDC_ABOUT_OK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hWnd, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	}
	return (INT_PTR)FALSE;
}
