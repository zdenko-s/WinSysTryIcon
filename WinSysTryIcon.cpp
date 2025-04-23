#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <map>
#include <filesystem>

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Comctl32.lib")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 9000
#define TRAY_ICON_UID 1

HINSTANCE hInst;
HWND hwnd;
NOTIFYICONDATA nid;
HMENU hTrayMenu;
std::wstring shortcutsDir = L"C:\\Util\\Util";  // <--- CHANGE ME

int commandIdCounter = 1000;
std::map<int, std::wstring> commandMap;

// Forward declarations
void ShowContextMenu(HWND hwnd);
void AddShortcutsToMenu(HMENU menu, const std::filesystem::path& dir);
void RunShortcut(const std::wstring& path);

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_TRAYICON:
        if (LOWORD(lParam) == WM_RBUTTONUP) {
            ShowContextMenu(hwnd);
        }
        break;
    case WM_COMMAND:
        if (LOWORD(wParam) == ID_TRAY_EXIT) {
            PostQuitMessage(0);
        }
        else if (commandMap.count(LOWORD(wParam))) {
            RunShortcut(commandMap[LOWORD(wParam)]);
        }
        break;
    case WM_DESTROY:
        Shell_NotifyIcon(NIM_DELETE, &nid);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void AddShortcutsToMenu(HMENU menu, const std::filesystem::path& dir) {
    for (const auto& entry : std::filesystem::directory_iterator(dir)) {
        if (entry.is_directory()) {
            HMENU subMenu = CreatePopupMenu();
            AddShortcutsToMenu(subMenu, entry.path());
            AppendMenu(menu, MF_POPUP, (UINT_PTR)subMenu, entry.path().filename().c_str());
        }
        else if (entry.path().extension() == L".lnk") {
            int id = commandIdCounter++;
            commandMap[id] = entry.path().wstring();

            SHFILEINFO sfi = {};
            SHGetFileInfo(entry.path().c_str(), 0, &sfi, sizeof(sfi),
                SHGFI_ICON | SHGFI_SMALLICON | SHGFI_DISPLAYNAME);

            AppendMenu(menu, MF_STRING, id, sfi.szDisplayName);

            if (sfi.hIcon) {
                MENUITEMINFO mii = {};
                mii.cbSize = sizeof(mii);
                mii.fMask = MIIM_BITMAP;
                mii.hbmpItem = CreateMappedBitmap(hInst, (INT_PTR)sfi.hIcon, 0, NULL, 0);
                SetMenuItemInfo(menu, id, FALSE, &mii);
                DestroyIcon(sfi.hIcon);
            }
        }
    }
}

void ShowContextMenu(HWND hwnd) {
    POINT pt;
    GetCursorPos(&pt);

    if (hTrayMenu) DestroyMenu(hTrayMenu);
    hTrayMenu = CreatePopupMenu();

    commandMap.clear();
    commandIdCounter = 1000;

    AddShortcutsToMenu(hTrayMenu, shortcutsDir);
    AppendMenu(hTrayMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hTrayMenu, MF_STRING, ID_TRAY_EXIT, L"Exit");

    SetForegroundWindow(hwnd);
    TrackPopupMenu(hTrayMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
}

void RunShortcut(const std::wstring& path) {
    CoInitialize(NULL);
    IShellLink* pShellLink = nullptr;
    if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
        IID_IShellLink, (void**)&pShellLink))) {
        IPersistFile* pPersistFile = nullptr;
        if (SUCCEEDED(pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile))) {
            if (SUCCEEDED(pPersistFile->Load(path.c_str(), STGM_READ))) {
                pShellLink->Resolve(hwnd, SLR_NO_UI);
                WCHAR target[MAX_PATH];
                if (SUCCEEDED(pShellLink->GetPath(target, MAX_PATH, NULL, 0))) {
                    ShellExecute(NULL, L"open", target, NULL, NULL, SW_SHOWNORMAL);
                }
            }
            pPersistFile->Release();
        }
        pShellLink->Release();
    }
    CoUninitialize();
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int) {
    hInst = hInstance;

    WNDCLASS wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"TrayAppClass";
    RegisterClass(&wc);

    hwnd = CreateWindow(wc.lpszClassName, L"TrayApp", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 300, 200,
        NULL, NULL, hInstance, NULL);

    nid = {};
    nid.cbSize = sizeof(nid);
    nid.hWnd = hwnd;
    nid.uID = TRAY_ICON_UID;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wcscpy_s(nid.szTip, L"Shortcut Menu");
    Shell_NotifyIcon(NIM_ADD, &nid);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}
