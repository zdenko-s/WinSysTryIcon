#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <filesystem>

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comctl32.lib")

#define WM_TRAYICON (WM_USER + 1)
#define TRAY_ID 1001
#define ID_MENU_BASE 2000

HINSTANCE hInst;
HWND hwnd;
NOTIFYICONDATA nid;

std::vector<std::wstring> shortcutPaths;

std::wstring shortcutDir = L"C:\\Util\\Util"; //Set your folder here

// Forward declarations
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void ShowShortcutMenu(HWND hwnd);
void LoadShortcuts();
void ExecuteShortcut(const std::wstring& path);
HICON GetIconFromShortcut(const std::wstring& path);

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int) {
	hInst = hInstance;

	WNDCLASS wc = {};
	wc.lpfnWndProc = WndProc;
	wc.hInstance = hInstance;
	wc.lpszClassName = L"MyTrayShortcutApp";
	RegisterClass(&wc);

	hwnd = CreateWindow(wc.lpszClassName, L"TrayShortcuts", WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, CW_USEDEFAULT, 300, 200,
		NULL, NULL, hInstance, NULL);

	LoadShortcuts();

	nid = {};
	nid.cbSize = sizeof(nid);
	nid.hWnd = hwnd;
	nid.uID = 1;
	nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	nid.uCallbackMessage = WM_TRAYICON;
	nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wcscpy_s(nid.szTip, L"Shortcut Menu");
	Shell_NotifyIcon(NIM_ADD, &nid);

	MSG msg;
	while (GetMessage(&msg, nullptr, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	Shell_NotifyIcon(NIM_DELETE, &nid);
	return 0;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_TRAYICON && LOWORD(lParam) == WM_RBUTTONUP) {
		ShowShortcutMenu(hwnd);
	}
	else if (msg == WM_COMMAND) {
		int id = LOWORD(wParam);
		if (id >= ID_MENU_BASE) {
			int index = id - ID_MENU_BASE;
			if (index >= 0 && index < shortcutPaths.size()) {
				ExecuteShortcut(shortcutPaths[index]);
			}
		}
	}
	else if (msg == WM_DESTROY) {
		PostQuitMessage(0);
	}

	return DefWindowProc(hwnd, msg, wParam, lParam);
}

void LoadShortcuts() {
	shortcutPaths.clear();
	for (const auto& entry : std::filesystem::directory_iterator(shortcutDir)) {
		if (entry.path().extension() == L".lnk") {
			shortcutPaths.push_back(entry.path().wstring());
		}
	}
}

void ShowShortcutMenu(HWND hwnd) {
	POINT pt;
	GetCursorPos(&pt);

	HMENU hMenu = CreatePopupMenu();

	for (int i = 0; i < shortcutPaths.size(); ++i) {
		std::wstring fileName = std::filesystem::path(shortcutPaths[i]).stem();
		HICON hIcon = GetIconFromShortcut(shortcutPaths[i]);

		MENUITEMINFO mii = {};
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_ID | MIIM_STRING | MIIM_FTYPE | MIIM_BITMAP;
		mii.fType = MFT_STRING;
		mii.wID = ID_MENU_BASE + i;
		mii.dwTypeData = const_cast<LPWSTR>(fileName.c_str());

		if (hIcon) {
			HBITMAP hBmp = NULL;
			ICONINFO iconInfo;
			if (GetIconInfo(hIcon, &iconInfo)) {
				hBmp = iconInfo.hbmColor;
				mii.hbmpItem = hBmp;
			}
		}

		InsertMenuItem(hMenu, i, TRUE, &mii);
	}

	SetForegroundWindow(hwnd);
	TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
	DestroyMenu(hMenu);
}

void ExecuteShortcut(const std::wstring& path) {
	ShellExecute(NULL, L"open", path.c_str(), NULL, NULL, SW_SHOWNORMAL);
}

HICON GetIconFromShortcut(const std::wstring& lnkPath) {
	CoInitialize(NULL);
	IShellLink* psl;
	WCHAR target[MAX_PATH];

	HICON hIcon = NULL;
	if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
		IID_IShellLink, (LPVOID*)&psl))) {
		IPersistFile* ppf;
		if (SUCCEEDED(psl->QueryInterface(IID_IPersistFile, (void**)&ppf))) {
			if (SUCCEEDED(ppf->Load(lnkPath.c_str(), STGM_READ))) {
				WIN32_FIND_DATA wfd;
				if (SUCCEEDED(psl->GetPath(target, MAX_PATH, &wfd, SLGP_RAWPATH))) {
					SHFILEINFO sfi;
					if (SHGetFileInfo(target, 0, &sfi, sizeof(sfi), SHGFI_ICON)) {
						hIcon = sfi.hIcon;
					}
				}
			}
			ppf->Release();
		}
		psl->Release();
	}
	CoUninitialize();
	return hIcon;
}
