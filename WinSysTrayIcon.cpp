#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <map>

#include <filesystem>
namespace fs = std::filesystem;

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comctl32.lib")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 9000
#define ID_TRAY_BASE 2000

std::map<int, std::wstring> menuCommandMap;
int currentMenuID = ID_TRAY_BASE;

// Shortcut directory. Read from ini file
std::wstring rootShortcutDir;

// Forward declarations
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void ShowContextMenu(HWND hwnd);
void PopulateMenuFromFolder(HMENU hMenu, const std::wstring& folder, bool recurse);
void ExecuteShortcut(const std::wstring& path);
HICON GetSmallIcon(const std::wstring& filePath);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR    lpCmdLine,
	_In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);
	// Read ini file
	wchar_t iniFileName[MAX_PATH];
	if (GetModuleFileName(nullptr, iniFileName, MAX_PATH))
	{
		fs::path iniFile = fs::path{ iniFileName };
		iniFile.replace_extension("ini");
		wchar_t readBuffer[MAX_PATH];
		*readBuffer = 0;
		GetPrivateProfileString(L"startup", L"Folder", L"", readBuffer, MAX_PATH, iniFile.c_str());
		if (*readBuffer)
			rootShortcutDir.assign(readBuffer);
	}

	auto hr = CoInitialize(NULL);

	WNDCLASS wc = {};
	wc.lpfnWndProc = WndProc;
	wc.hInstance = hInstance;
	wc.lpszClassName = L"ShortcutTrayAppClass";
	RegisterClass(&wc);

	HWND hwnd = CreateWindow(wc.lpszClassName, L"TrayApp", WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, CW_USEDEFAULT, 300, 200,
		NULL, NULL, hInstance, NULL);

	// Add tray icon
	NOTIFYICONDATA nid = {};
	nid.cbSize = sizeof(nid);
	nid.hWnd = hwnd;
	nid.uID = 1;
	nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	nid.uCallbackMessage = WM_TRAYICON;
	nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wcscpy_s(nid.szTip, L"Shortcut Tray");

	DWORD rc = 0;
	wchar_t buffer[MAX_PATH];
	rc = GetModuleFileName(nullptr, buffer, MAX_PATH);
	if (rc)
	{
		nid.hIcon = GetSmallIcon(buffer);
	}

	Shell_NotifyIcon(NIM_ADD, &nid);

	MSG msg;
	while (GetMessage(&msg, nullptr, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	Shell_NotifyIcon(NIM_DELETE, &nid);
	CoUninitialize();
	return 0;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_TRAYICON) {
		if (LOWORD(lParam) == WM_LBUTTONUP || LOWORD(lParam) == WM_RBUTTONUP) {
			ShowContextMenu(hwnd);
		}
	}
	else if (msg == WM_COMMAND) {
		if (LOWORD(wParam) == ID_TRAY_EXIT) {
			PostQuitMessage(0);
		}
		else {
			int id = LOWORD(wParam);
			if (menuCommandMap.count(id)) {
				ExecuteShortcut(menuCommandMap[id]);
			}
		}
	}
	else if (msg == WM_DESTROY) {
		PostQuitMessage(0);
	}

	return DefWindowProc(hwnd, msg, wParam, lParam);
}

void ShowContextMenu(HWND hwnd) {
	POINT pt;
	GetCursorPos(&pt);

	HMENU hMenu = CreatePopupMenu();
	currentMenuID = ID_TRAY_BASE;
	menuCommandMap.clear();

	PopulateMenuFromFolder(hMenu, rootShortcutDir, true);

	AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
	AppendMenu(hMenu, MF_STRING, ID_TRAY_EXIT, L"Exit");

	SetForegroundWindow(hwnd);
	TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
	DestroyMenu(hMenu);
}

void PopulateMenuFromFolder(HMENU hMenu, const std::wstring& folder, bool recurse) {
	WIN32_FIND_DATA findData;
	std::wstring searchPath = folder + L"\\*";
	HANDLE hFind = FindFirstFile(searchPath.c_str(), &findData);

	if (hFind == INVALID_HANDLE_VALUE)
		return;

	do {
		std::wstring name = findData.cFileName;
		if (name == L"." || name == L"..")
			continue;

		std::wstring fullPath = folder + L"\\" + name;

		if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY && recurse) {
			HMENU subMenu = CreatePopupMenu();
			PopulateMenuFromFolder(subMenu, fullPath, recurse);
			InsertMenu(hMenu, -1, MF_BYPOSITION | MF_POPUP, (UINT_PTR)subMenu, name.c_str());
		}
		else if (PathMatchSpec(name.c_str(), L"*.lnk")) {
			HICON hIcon = GetSmallIcon(fullPath);
			MENUITEMINFO mii = { sizeof(MENUITEMINFO) };
			mii.fMask = MIIM_ID | MIIM_STRING | MIIM_FTYPE | MIIM_BITMAP;
			mii.fType = MFT_STRING;
			mii.wID = currentMenuID;
			// Remove extension
			std::wstring base_name = findData.cFileName;
			auto menuitem = base_name.substr(0, base_name.length() - 4);

			mii.dwTypeData = (LPWSTR)menuitem.c_str();

			HBITMAP hBmp = NULL;
			if (hIcon) {
				ICONINFO iconInfo;
				GetIconInfo(hIcon, &iconInfo);
				hBmp = iconInfo.hbmColor;
				mii.hbmpItem = hBmp;
			}

			InsertMenuItem(hMenu, -1, TRUE, &mii);
			menuCommandMap[currentMenuID++] = fullPath;
		}
	} while (FindNextFile(hFind, &findData));

	FindClose(hFind);
}

HICON GetSmallIcon(const std::wstring& filePath) {
	SHFILEINFO sfi{ 0 };
	SHGetFileInfo(filePath.c_str(), 0, &sfi, sizeof(sfi),
		SHGFI_ICON | SHGFI_SMALLICON);
	return sfi.hIcon;
}

void ExecuteShortcut(const std::wstring& path) {
	IShellLink* psl = NULL;
	HRESULT hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
		IID_IShellLink, (LPVOID*)&psl);
	if (SUCCEEDED(hr)) {
		IPersistFile* ppf = NULL;
		hr = psl->QueryInterface(IID_IPersistFile, (LPVOID*)&ppf);
		if (SUCCEEDED(hr)) {
			hr = ppf->Load(path.c_str(), STGM_READ);
			if (SUCCEEDED(hr)) {
				psl->Resolve(NULL, SLR_NO_UI);
				WIN32_FIND_DATA wfd{ 0 };
				WCHAR szTarget[MAX_PATH];
				psl->GetPath(szTarget, MAX_PATH, &wfd, SLGP_UNCPRIORITY);
				ShellExecute(NULL, L"open", szTarget, NULL, NULL, SW_SHOWNORMAL);
			}
			ppf->Release();
		}
		psl->Release();
	}
}
