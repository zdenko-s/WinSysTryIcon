#include <windows.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <vector>
#include <string>
#include <filesystem>

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_BASE 2000

HINSTANCE hInst;
HWND hwnd;
NOTIFYICONDATA nid;
std::wstring exeFolder = L"C:\\Util\\putty.74";  // Change this to your folder

struct MenuItemInfo {
    std::wstring name;
    std::wstring path;
    HICON icon;
};
std::vector<MenuItemInfo> menuItems;

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void ShowContextMenu(HWND hwnd);
void LoadMenuItems();

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int) {
    hInst = hInstance;

    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"TrayAppWithIcons";
    RegisterClass(&wc);

    hwnd = CreateWindow(wc.lpszClassName, L"TrayApp", WS_OVERLAPPEDWINDOW,
        0, 0, 300, 200, NULL, NULL, hInstance, NULL);

    // Add tray icon
    nid = {};
    nid.cbSize = sizeof(nid);
    nid.hWnd = hwnd;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wcscpy_s(nid.szTip, L"Executable Menu");
    Shell_NotifyIcon(NIM_ADD, &nid);

    LoadMenuItems();

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // Clean up icons
    for (auto& item : menuItems)
        if (item.icon) DestroyIcon(item.icon);

    Shell_NotifyIcon(NIM_DELETE, &nid);
    return 0;
}

void LoadMenuItems() {
    using namespace std::filesystem;
    menuItems.clear();
    for (const auto& entry : directory_iterator(exeFolder)) {
        if (entry.is_regular_file() && entry.path().extension() == L".EXE") {
            MenuItemInfo info;
            info.path = entry.path().wstring();
            info.name = entry.path().filename().wstring();

            SHFILEINFO sfi = {};
            SHGetFileInfo(info.path.c_str(), 0, &sfi, sizeof(sfi), SHGFI_ICON | SHGFI_SMALLICON);
            info.icon = sfi.hIcon;

            menuItems.push_back(info);
        }
    }
}

void ShowContextMenu(HWND hwnd) {
    POINT pt;
    GetCursorPos(&pt);
    HMENU hMenu = CreatePopupMenu();

    int id = ID_TRAY_BASE;
    for (const auto& item : menuItems) {
        MENUITEMINFO mii = { sizeof(MENUITEMINFO) };
        mii.fMask = MIIM_STRING | MIIM_ID | MIIM_FTYPE | MIIM_BITMAP;
        mii.fType = MFT_STRING;
        mii.wID = id++;
        mii.dwTypeData = const_cast<LPWSTR>(item.name.c_str());

        if (item.icon) {
            HBITMAP hBmp = nullptr;
            ICONINFO ii;
            GetIconInfo(item.icon, &ii);
            hBmp = ii.hbmColor;

            mii.fMask |= MIIM_BITMAP;
            mii.hbmpItem = HBMMENU_CALLBACK;  // let system draw icon automatically
        }

        InsertMenuItem(hMenu, -1, TRUE, &mii);
    }

    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING, 9999, L"Exit");

    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_TRAYICON && LOWORD(lParam) == WM_RBUTTONUP) {
        ShowContextMenu(hwnd);
    }
    else if (msg == WM_COMMAND) {
        int cmd = LOWORD(wParam);
        if (cmd == 9999) {
            PostQuitMessage(0);
        }
        else if (cmd >= ID_TRAY_BASE && cmd < ID_TRAY_BASE + 1000) {
            int index = cmd - ID_TRAY_BASE;
            if (index >= 0 && index < menuItems.size()) {
                ShellExecute(NULL, L"open", menuItems[index].path.c_str(), NULL, NULL, SW_SHOWNORMAL);
            }
        }
    }
    else if (msg == WM_DESTROY) {
        PostQuitMessage(0);
    }

    return DefWindowProc(hwnd, msg, wParam, lParam);
}
