#pragma once

class Tasktray
{
	struct MenuPosition
	{
		unsigned int flags;
		int x;
		int y;
	};

public:
	Tasktray();
	~Tasktray();

	void				Init(HWND hWnd);
	void				CreateMenu();
	void				Show();
	void				Hide();
	void				Tip(WCHAR* buf);
	void				Message(WPARAM wPAram, LPARAM lParam);
	UINT				GetTaskBarLocation();
	UINT				GetTaskBarLocation(RECT* rectangle);
	void				GetMenuPosition(struct MenuPosition* test);

	NOTIFYICONDATA		m_nid = { 0 };

private:
	HMENU				m_menu = NULL;
	HWND				m_hWnd = NULL;

	bool				m_flag = false;
};

	extern Tasktray		tasktray;
