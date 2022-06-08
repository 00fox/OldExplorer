#pragma once
#include "WebBrowser.h"

class Explorer
{
	//DECLARE_DYNAMIC(Explorer)

public:
	Explorer();
	~Explorer();

	void			Init(HWND);
	void			Show();
	void			Hide();
	BOOL			MoveWindow(int, int, int, int, BOOL);

	WebBrowser*		explorer;

private:
	static INT_PTR CALLBACK Proc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	INT_PTR CALLBACK _proc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	HWND			m_hWnd = NULL;
	HWND			m_hDlg = NULL;
	BYTE			LastLinkUsed = 0;
	POINT			mousepoint = { (0, 0) };

	bool			firsttime = false;
	bool			autorefresh = false;
	int				posx = 0;
	int				posy = 0;
	bool			mousein = false;
};

extern				Explorer eDlg;