#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <shlwapi.h>
#include <string>
#include <vector>
#include <cwctype>
#include "resource.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "Shlwapi.lib")

/// =============================================================

const int MAX_TABS = 50; // Maximum number of tabs

int g_tabCount = 0;
int g_buttonRows = 3;
int g_buttonCols = 8;
int g_buttonCountPerTab =
g_buttonRows * g_buttonCols; // Maximum number of buttons per tab

LPCWSTR g_lpWindowName = L"MultiTab Launcher";
const HICON g_hDefaultIcon = LoadIcon(NULL, IDI_APPLICATION);

std::wstring g_exeDirPath;
std::wstring g_iniFilePath;

HWND g_hWnd = NULL;
HWND g_hTabControl = NULL;
std::vector<std::wstring> g_tabNames;

struct ButtonInfo {
    HWND hButton{ NULL };
    std::wstring name{ L"" };
    std::wstring path{ L"" };
    std::wstring parameters{ L"" };
    bool adminMode{ false };
    HICON hIcon{ NULL };
};
std::vector<std::vector<ButtonInfo>> g_buttonInfos;

HBRUSH g_hBackgroundBrush = NULL;
HBRUSH g_hTabBrush = NULL;
HBRUSH g_hButtonBrush = NULL;
HBRUSH g_hButtonHoverBrush = NULL;

int g_currentTab = 0;

/// =============================================================
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
void CreateTabControl(HWND hwnd);
void CreateTabButtons(HWND hwnd, int tabIndex);
void ResizeControls(HWND hwnd);
void InitializeColors();
void SaveWindowPlacement(HWND hwnd);
void LoadWindowPlacement(HWND hwnd);
bool ExecuteProcess(const std::wstring& filePath,
    const std::wstring& parameters, bool asAdmin);
void CreateIni();
bool CreateIniWithDefaults();
bool CreateUTF16LEBOMFile(const wchar_t* filename);
bool SaveFileAsUTF16LE(const wchar_t* filename, const std::wstring& text);
std::wstring GetConfigDefault();
void LoadConfigFromIni();
bool SaveButtonInfoToIni(int tabIndex, int buttonIndex, const ButtonInfo& info);

HICON GetFileIcon(const std::wstring& filePath);
std::wstring FindExePath(const wchar_t* targetFile);
void onPressButton(int tabIndex, int buttonIndex);
std::wstring ExpandEnvVars(const std::wstring& str);
inline void trim(std::wstring& s);
inline void rtrim(std::wstring& s);
inline void ltrim(std::wstring& s);

// Button Dialog
INT_PTR CALLBACK ButtonInfoDialogProc(HWND hDlg, UINT message, WPARAM wParam,
    LPARAM lParam);
LRESULT CALLBACK EditSubclassProc(HWND hEdit, UINT msg, WPARAM wParam,
    LPARAM lParam, UINT_PTR, DWORD_PTR);
int ShowButtonInfoDialog(int tabIdx, int btnIdx);

/// =============================================================

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow) {
    INITCOMMONCONTROLSEX icc = { sizeof(INITCOMMONCONTROLSEX), ICC_TAB_CLASSES };
    InitCommonControlsEx(&icc);

    // 실행파일 경로
    wchar_t buffer[MAX_PATH] = { 0 };

    // 1. 현재 실행 파일이 있는 디렉토리, inifile 경로 확인
    DWORD modulePathLength = GetModuleFileNameW(NULL, buffer, MAX_PATH);
    if (modulePathLength > 0) {
        std::filesystem::path exePath(buffer);
        std::filesystem::path exeDir = exePath.parent_path();
        g_exeDirPath = exePath.parent_path().wstring();
        SetCurrentDirectoryW(g_exeDirPath.c_str());
        g_iniFilePath = g_exeDirPath + L"\\MultiTabLauncher.ini";
    }
    else {
        // 오류 처리
        MessageBox(NULL, L"Failed to get the executable path.", L"Error",
            MB_OK | MB_ICONERROR);
        return 1;
    }

    LoadConfigFromIni();
    InitializeColors();

    WNDCLASSEX wc = { sizeof(WNDCLASSEX) };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = g_hBackgroundBrush;
    wc.lpszClassName = g_lpWindowName;
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));
    wc.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));

    RegisterClassEx(&wc);

    g_hWnd =
        CreateWindowEx(WS_EX_CLIENTEDGE, g_lpWindowName, L"MultiTab Launcher",
            WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 800,
            600, NULL, NULL, hInstance, NULL);

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    DeleteObject(g_hBackgroundBrush);
    DeleteObject(g_hTabBrush);
    DeleteObject(g_hButtonBrush);
    DeleteObject(g_hButtonHoverBrush);
    return (int)msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
    {
        LoadWindowPlacement(hwnd);
        CreateTabControl(hwnd);
        for (int i = 0; i < g_buttonCountPerTab; i++)
            ShowWindow(g_buttonInfos[0][i].hButton, SW_SHOW);
    }


    break;
    case WM_SIZE:
        ResizeControls(hwnd);
        break;
    case WM_NOTIFY: {
        LPNMHDR nmhdr = (LPNMHDR)lParam;
        if (nmhdr->hwndFrom == g_hTabControl && nmhdr->code == TCN_SELCHANGE) {
            int newTab = TabCtrl_GetCurSel(g_hTabControl);
            for (int i = 0; i < g_buttonCountPerTab; i++)
                ShowWindow(g_buttonInfos[newTab][i].hButton, SW_SHOW);
            for (int i = 0; i < g_buttonCountPerTab; i++)
                ShowWindow(g_buttonInfos[g_currentTab][i].hButton, SW_HIDE);

            g_currentTab = newTab;
            InvalidateRect(hwnd, NULL, FALSE);
        }
        break;
    }
    case WM_COMMAND: {
        int buttonId = LOWORD(wParam);
        if (buttonId >= 1000 && buttonId < 1000 + (g_tabCount * 100)) {
            int tabIndex = (buttonId - 1000) / 100;
            int buttonIndex = (buttonId - 1000) % 100;

            onPressButton(tabIndex, buttonIndex);
        }
        break;
    }

    case WM_CONTEXTMENU: {
        HWND hCtrl = (HWND)wParam;

        // Right click on the button.
        for (int tab = 0; tab < g_tabCount; ++tab) {
            for (int buttonIndex = 0; buttonIndex < g_buttonCountPerTab;
                ++buttonIndex) {
                if (hCtrl == g_buttonInfos[tab][buttonIndex].hButton) {
                    if (ShowButtonInfoDialog(tab, buttonIndex) == IDOK) {
                        SetWindowTextW(g_buttonInfos[tab][buttonIndex].hButton,
                            g_buttonInfos[tab][buttonIndex].name.c_str());
                        if (g_buttonInfos[tab][buttonIndex].path.empty() == false) {
                            g_buttonInfos[tab][buttonIndex].hIcon =
                                GetFileIcon(g_buttonInfos[tab][buttonIndex].path);
                        }
                        else {
                            g_buttonInfos[tab][buttonIndex].hIcon = NULL;
                        }
                        InvalidateRect(hCtrl, NULL, TRUE);
                        SaveButtonInfoToIni(tab, buttonIndex,
                            g_buttonInfos[tab][buttonIndex]);
                        SetCurrentDirectoryW(g_exeDirPath.c_str());
                    }
                }
            }
        }
        break;
    }

    case WM_CTLCOLORBTN:
        SetTextColor((HDC)wParam, RGB(255, 255, 255));
        SetBkColor((HDC)wParam, RGB(45, 45, 45));
        return (LRESULT)g_hButtonBrush;

    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT pDIS = (LPDRAWITEMSTRUCT)lParam;
        if (pDIS->CtlType == ODT_BUTTON) {
            // Background color
            HBRUSH brush = (pDIS->itemState & ODS_SELECTED)
                ? CreateSolidBrush(RGB(9, 71, 113))
                : g_hButtonBrush;
            FillRect(pDIS->hDC, &pDIS->rcItem, brush);
            if (pDIS->itemState & ODS_SELECTED)
                DeleteObject(brush);

            // Border color
            HPEN pen = CreatePen(PS_SOLID, 1, RGB(50, 50, 50));
            HPEN oldPen = (HPEN)SelectObject(pDIS->hDC, pen);
            HBRUSH oldBrush =
                (HBRUSH)SelectObject(pDIS->hDC, GetStockObject(NULL_BRUSH));
            Rectangle(pDIS->hDC, pDIS->rcItem.left, pDIS->rcItem.top,
                pDIS->rcItem.right, pDIS->rcItem.bottom);
            SelectObject(pDIS->hDC, oldBrush);
            SelectObject(pDIS->hDC, oldPen);
            DeleteObject(pen);

            // icon
            int iconSize = 32;
            int iconPaddingTop = 10;
            int spaceBetween = 8;

            int buttonId = GetDlgCtrlID(pDIS->hwndItem);
            int tabIndex = (buttonId - 1000) / 100;
            int btnIndex = (buttonId - 1000) % 100;
            const ButtonInfo& btnInfo = g_buttonInfos[tabIndex][btnIndex];

            int iconY = pDIS->rcItem.top + iconPaddingTop;
            int iconX = pDIS->rcItem.left +
                (pDIS->rcItem.right - pDIS->rcItem.left - iconSize) / 2;

            int totalHeight = iconSize + spaceBetween;

            // text
            WCHAR text[256];
            GetWindowText(pDIS->hwndItem, text, 256);

            SIZE textSize{};
            GetTextExtentPoint32(pDIS->hDC, text, lstrlenW(text), &textSize);
            totalHeight += textSize.cy;

            // Center Align
            int availableHeight = pDIS->rcItem.bottom - pDIS->rcItem.top;
            int startY = pDIS->rcItem.top + (availableHeight - totalHeight) / 2;

            if (btnInfo.hIcon) {
                DrawIconEx(pDIS->hDC, iconX, startY, btnInfo.hIcon, iconSize, iconSize,
                    0, NULL, DI_NORMAL);
            }

            SetTextColor(pDIS->hDC, RGB(204, 204, 204));
            SetBkMode(pDIS->hDC, TRANSPARENT);

            RECT rcText = pDIS->rcItem;
            rcText.top = startY + iconSize + spaceBetween;
            rcText.bottom = rcText.top + textSize.cy;

            DrawText(pDIS->hDC, text, -1, &rcText,
                DT_CENTER | DT_TOP | DT_SINGLELINE | DT_END_ELLIPSIS);

            return TRUE;
        }
        break;
    }

    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        HDC memDC = CreateCompatibleDC(hdc);
        RECT rc;
        GetClientRect(hwnd, &rc);
        HBITMAP bmp = CreateCompatibleBitmap(hdc, rc.right, rc.bottom);
        HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, bmp);
        FillRect(memDC, &rc, g_hBackgroundBrush);
        BitBlt(hdc, 0, 0, rc.right, rc.bottom, memDC, 0, 0, SRCCOPY);
        SelectObject(memDC, oldBmp);
        DeleteObject(bmp);
        DeleteDC(memDC);
        EndPaint(hwnd, &ps);
        break;
    }
    case WM_DESTROY:
        SaveWindowPlacement(hwnd);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}


/// =============================================================

void InitializeColors() {
    g_hBackgroundBrush = CreateSolidBrush(RGB(30, 30, 30));
    g_hTabBrush = CreateSolidBrush(RGB(37, 37, 38));
    g_hButtonBrush = CreateSolidBrush(RGB(60, 60, 60));
    g_hButtonHoverBrush = CreateSolidBrush(RGB(78, 78, 78));
}

void CreateTabControl(HWND hwnd) {
    g_hTabControl =
        CreateWindowEx(0, WC_TABCONTROL, L"",
            WS_CHILD | WS_VISIBLE | TCS_FLATBUTTONS | TCS_BUTTONS, 0,
            0, 0, 0, hwnd, NULL, GetModuleHandle(NULL), NULL);

    // Font
    LOGFONT lf = {};
    lf.lfHeight = 24;
    lf.lfWeight = FW_BOLD;
    HFONT hFont = CreateFontIndirect(&lf);
    SendMessage(g_hTabControl, WM_SETFONT, (WPARAM)hFont, TRUE);
    TCITEM tie = { TCIF_TEXT };
    for (int i = 0; i < g_tabCount; ++i) {
        tie.pszText = (LPWSTR)g_tabNames[i].c_str();
        TabCtrl_InsertItem(g_hTabControl, i, &tie);
        CreateTabButtons(hwnd, i);
    }
}

void CreateTabButtons(HWND hwnd, int tabIndex) {
    RECT rcTab;
    GetClientRect(g_hTabControl, &rcTab);
    TabCtrl_AdjustRect(g_hTabControl, FALSE, &rcTab);
    int bw = (rcTab.right - rcTab.left) / g_buttonCols;
    int bh = (rcTab.bottom - rcTab.top) / g_buttonRows;

    for (int row = 0; row < g_buttonRows; ++row) {
        for (int col = 0; col < g_buttonCols; ++col) {
            int i = row * g_buttonCols + col;
            int id = 1000 + tabIndex * 100 + i;
            g_buttonInfos[tabIndex][i].hButton =
                CreateWindowEx(0, L"BUTTON", g_buttonInfos[tabIndex][i].name.c_str(),
                    WS_CHILD | BS_PUSHBUTTON | BS_OWNERDRAW,
                    rcTab.left + col * bw, rcTab.top + row * bh, bw, bh,
                    hwnd, (HMENU)(INT_PTR)id, GetModuleHandle(NULL), NULL);
        }
    }
}

void ResizeControls(HWND hwnd) {
    RECT rcClient;
    GetClientRect(hwnd, &rcClient);
    MoveWindow(g_hTabControl, 0, 0, rcClient.right, rcClient.bottom, TRUE);
    RECT rcTab;
    GetClientRect(g_hTabControl, &rcTab);
    TabCtrl_AdjustRect(g_hTabControl, FALSE, &rcTab);
    int bw = (rcTab.right - rcTab.left) / g_buttonCols;
    int bh = (rcTab.bottom - rcTab.top) / g_buttonRows;

    for (int tab = 0; tab < g_tabCount; ++tab) {
        for (int i = 0; i < g_buttonCountPerTab; ++i) {
            if (g_buttonInfos[tab][i].hButton) {
                int row = i / g_buttonCols;
                int col = i % g_buttonCols;
                MoveWindow(g_buttonInfos[tab][i].hButton, rcTab.left + col * bw,
                    rcTab.top + row * bh, bw, bh, TRUE);
            }
        }
    }
    InvalidateRect(hwnd, NULL, FALSE);
}


// Create default INI
bool CreateIniWithDefaults() {
    return SaveFileAsUTF16LE(g_iniFilePath.c_str(), GetConfigDefault());
}

// Not in use
void CreateIni() {

    if (CreateUTF16LEBOMFile(g_iniFilePath.c_str()) == false) {
        MessageBox(NULL, L"Failed to create config.ini file.", L"Error",
            MB_OK | MB_ICONERROR);
        std::exit(0);
        return;
    }

    const int tabCount = 10;
    const wchar_t* tabNames[tabCount] = {
        L"Home",  L"Document", L"Program",     L"System", L"Network",
        L"Media", L"Utility",  L"Development", L"Tab9",   L"Tab10" };

    WritePrivateProfileStringW(L"Tabs", L"Count",
        std::to_wstring(tabCount).c_str(),
        g_iniFilePath.c_str());

    for (int i = 0; i < tabCount; i++) {
        wchar_t key[20];
        swprintf_s(key, 20, L"Tab%d", i);
        WritePrivateProfileStringW(L"Tabs", key, tabNames[i],
            g_iniFilePath.c_str());

        for (int j = 0; j < g_buttonCountPerTab; j++) {
            ButtonInfo info;
            info.name = L"";
            info.path = L"";
            info.parameters = L"";
            info.adminMode = false;
            SaveButtonInfoToIni(i, j, info);
        }
    }
}

void LoadConfigFromIni() {
    if (!PathFileExists(g_iniFilePath.c_str())) {
        CreateIniWithDefaults();
    }

    for (int i = 0; i < 3; i++) {
        if (i >= 2) {
            MessageBox(NULL, L"Failed to create config.ini file.", L"Error",
                MB_OK | MB_ICONERROR);
            std::exit(0);
            return;
        }

        g_tabCount =
            GetPrivateProfileInt(L"Tabs", L"Count", 0, g_iniFilePath.c_str());
        if (g_tabCount <= 0) {
            CreateIniWithDefaults();
            continue;
        }
        else if (g_tabCount > MAX_TABS) {
            g_tabCount = MAX_TABS;
        }


        g_buttonRows =
            GetPrivateProfileInt(L"Tabs", L"ButtonRows", 1, g_iniFilePath.c_str());

        g_buttonCols =
            GetPrivateProfileInt(L"Tabs", L"ButtonCols", 1, g_iniFilePath.c_str());

        g_buttonCountPerTab = g_buttonRows * g_buttonCols;

        break;
    }

    g_tabNames.resize(g_tabCount);
    g_buttonInfos.resize(g_tabCount,
        std::vector<ButtonInfo>(g_buttonCountPerTab));

    // tab Names
    for (int i = 0; i < g_tabCount; i++) {
        WCHAR buffer[256] = { 0 };
        GetPrivateProfileString(L"Tabs", (L"Tab" + std::to_wstring(i)).c_str(),
            g_tabNames[i].c_str(), buffer, 256,
            g_iniFilePath.c_str());

        if (lstrlen(buffer) > 0 && lstrlen(buffer) < 30) {
            g_tabNames[i] = buffer;
        }
        else {
            g_tabNames[i] = L"Tab" + std::to_wstring(i + 1);
        }
    }

    // Button
    for (int tab = 0; tab < g_tabCount; tab++) {

        std::wstring section = L"Tab" + std::to_wstring(tab);

        for (int btn = 0; btn < g_buttonCountPerTab; btn++) {
            std::wstring key = L"Button" + std::to_wstring(btn);
            WCHAR buffer[512] = { 0 };

            // Button Names
            GetPrivateProfileStringW(section.c_str(), (key + L"_Name").c_str(),
                g_buttonInfos[tab][btn].name.c_str(), buffer,
                512, g_iniFilePath.c_str());
            g_buttonInfos[tab][btn].name = buffer;
            trim(g_buttonInfos[tab][btn].name);

            // path
            {
                GetPrivateProfileString(section.c_str(), (key + L"_Path").c_str(), L"",
                    buffer, 512, g_iniFilePath.c_str());

                if (lstrlen(buffer) > 0) {
                    g_buttonInfos[tab][btn].path = buffer;
                    g_buttonInfos[tab][btn].hIcon =
                        GetFileIcon(g_buttonInfos[tab][btn].path);
                    trim(g_buttonInfos[tab][btn].path);

                }
                else {
                    g_buttonInfos[tab][btn].path = L"";
                    g_buttonInfos[tab][btn].hIcon = NULL;
                }
            }

            // parameters
            {
                GetPrivateProfileString(section.c_str(), (key + L"_Params").c_str(),
                    L"", buffer, 512, g_iniFilePath.c_str());

                if (lstrlen(buffer) > 0) {
                    // %USERPROFILE% Environment variable substitution.
                    g_buttonInfos[tab][btn].parameters = buffer;
                    trim(g_buttonInfos[tab][btn].parameters);
                }
                else {
                    g_buttonInfos[tab][btn].parameters = L"";
                }
            }

            // adminMode
            g_buttonInfos[tab][btn].adminMode =
                GetPrivateProfileInt(section.c_str(), (key + L"_Admin").c_str(), 0,
                    g_iniFilePath.c_str()) != 0;
        }
    }
}

bool SaveButtonInfoToIni(int tabIndex, int buttonIndex,
    const ButtonInfo& info) {
    std::wstring section = L"Tab" + std::to_wstring(tabIndex);
    std::wstring btnKey = L"Button" + std::to_wstring(buttonIndex);

    BOOL ok = TRUE;
    ok &= WritePrivateProfileStringW(section.c_str(), (btnKey + L"_Name").c_str(),
        info.name.c_str(), g_iniFilePath.c_str());
    ok &= WritePrivateProfileStringW(section.c_str(), (btnKey + L"_Path").c_str(),
        info.path.c_str(), g_iniFilePath.c_str());
    ok &= WritePrivateProfileStringW(
        section.c_str(), (btnKey + L"_Params").c_str(), info.parameters.c_str(),
        g_iniFilePath.c_str());
    ok &= WritePrivateProfileStringW(
        section.c_str(), (btnKey + L"_Admin").c_str(),
        info.adminMode ? L"1" : L"0", g_iniFilePath.c_str());

    return ok != 0;
}


/// =============================================================

bool ExecuteProcess(const std::wstring& filePath,
    const std::wstring& parameters, bool asAdmin) {
    // asAdmin
    std::wstring operation = asAdmin ? L"runas" : L"open";

    HINSTANCE result = ShellExecuteW(
        NULL,
        operation.c_str(),
        filePath.c_str(),
        parameters.empty()
        ? NULL
        : ExpandEnvVars(parameters.c_str()).c_str(),
        NULL,
        SW_SHOWNORMAL
    );

    if ((INT_PTR)result > 32) {
        return true;
    }

    std::wstring absPath = FindExePath(filePath.c_str());
    if (!absPath.empty() && absPath != filePath) {
        result = ShellExecuteW(
            NULL,
            operation.c_str(),
            absPath.c_str(),
            parameters.empty() ? NULL
            : ExpandEnvVars(parameters.c_str())
            .c_str(),
            NULL,
            SW_SHOWNORMAL
        );
        if ((INT_PTR)result > 32) {
            return true;
        }
    }

    std::wstring msg = L"ShellExecuteW failed.\n";
    msg += L"File: " + filePath + L"\n";
    msg += L"Error code: " + std::to_wstring((INT_PTR)result) + L"\n";

    switch ((INT_PTR)result) {
    case ERROR_SUCCESS:
        msg += L"Description: The operating system is out of memory or resources.";
        break;
    case ERROR_FILE_NOT_FOUND:
        msg += L"Description: The specified file was not found.";
        break;
    case ERROR_PATH_NOT_FOUND:
        msg += L"Description: The specified path was not found.";
        break;
    case ERROR_ACCESS_DENIED:
        msg += L"Description: Access is denied.";
        break;
    case ERROR_NOT_ENOUGH_MEMORY:
        msg += L"Description: Not enough memory to complete the operation.";
        break;
    case ERROR_GEN_FAILURE:
        msg += L"Description: There is no application associated with the given "
            L"file extension.";
        break;
    default:
        msg += L"Description: Unknown error.";
        break;
    }

    MessageBoxW(NULL, msg.c_str(), L"Execution Error", MB_OK | MB_ICONERROR);

    return false;
}


void SaveWindowPlacement(HWND hwnd) {
    WINDOWPLACEMENT wp = { sizeof(WINDOWPLACEMENT) };
    if (GetWindowPlacement(hwnd, &wp)) {
        HKEY hKey;
        if (RegCreateKeyEx(HKEY_CURRENT_USER, L"Software\\MultiTabLauncher", 0,
            NULL, 0, KEY_WRITE, NULL, &hKey,
            NULL) == ERROR_SUCCESS) {
            RegSetValueEx(hKey, L"Placement", 0, REG_BINARY, (BYTE*)&wp, sizeof(wp));
            RegCloseKey(hKey);
        }
    }
}

void LoadWindowPlacement(HWND hwnd) {
    WINDOWPLACEMENT wp = { sizeof(WINDOWPLACEMENT) };
    HKEY hKey;
    DWORD size = sizeof(wp);
    if (RegOpenKeyEx(HKEY_CURRENT_USER, L"Software\\MultiTabLauncher", 0,
        KEY_READ, &hKey) == ERROR_SUCCESS) {
        if (RegQueryValueEx(hKey, L"Placement", NULL, NULL, (LPBYTE)&wp, &size) ==
            ERROR_SUCCESS) {
            SetWindowPlacement(hwnd, &wp);
        }
        RegCloseKey(hKey);
    }
}

HICON GetFileIcon(const std::wstring& filePath) {
    std::wstring absPath = FindExePath(filePath.c_str());
    SHFILEINFO sfi = {};
    DWORD attribs = GetFileAttributesW(absPath.c_str());
    if (attribs != INVALID_FILE_ATTRIBUTES) {
        if (SHGetFileInfo(absPath.c_str(), 0, &sfi, sizeof(sfi), SHGFI_ICON)) {
            return sfi.hIcon;
        }
    }

    return g_hDefaultIcon;
}

std::wstring FindExePath(const wchar_t* targetFile) {
    wchar_t buffer[MAX_PATH] = { 0 };

    // Search in current executable file directory
    if (g_exeDirPath.empty() == false) {
        std::filesystem::path exeDir(g_exeDirPath);
        std::filesystem::path targetPath = exeDir / targetFile;

        if (std::filesystem::exists(targetPath)) {
            return targetPath.wstring();
        }
    }

    // Search in system PATH
    DWORD pathLength = SearchPathW(NULL,
        targetFile,
        NULL,
        MAX_PATH,
        buffer,
        NULL
    );

    if (pathLength > 0 && pathLength < MAX_PATH) {
        return std::wstring(buffer);
    }

    // Search in Windows directory
    wchar_t winDir[MAX_PATH] = { 0 };
    if (GetWindowsDirectoryW(winDir, MAX_PATH) > 0) {
        std::filesystem::path windowsPath(winDir);
        std::filesystem::path system32Path = windowsPath / L"System32" / targetFile;

        if (std::filesystem::exists(system32Path)) {
            return system32Path.wstring();
        }

        std::filesystem::path windowsNotepadPath = windowsPath / targetFile;
        if (std::filesystem::exists(windowsNotepadPath)) {
            return windowsNotepadPath.wstring();
        }
    }

    return L"";
}

void onPressButton(int tabIndex, int buttonIndex) {
    const ButtonInfo& buttonInfo = g_buttonInfos[tabIndex][buttonIndex];
    if (!buttonInfo.path.empty()) {
        ExecuteProcess(buttonInfo.path, buttonInfo.parameters,
            buttonInfo.adminMode);
        SetCurrentDirectoryW(g_exeDirPath.c_str());
    }
}

bool CreateUTF16LEBOMFile(const wchar_t* filename) {
    FILE* fp = nullptr;
    if (_wfopen_s(&fp, filename, L"wb") != 0 || !fp)
        return false;

    // UTF-16 LE BOM
    unsigned char bom[] = { 0xFF, 0xFE };
    size_t written = fwrite(bom, 1, sizeof(bom), fp);
    fclose(fp);

    return (written == sizeof(bom));
}

std::wstring ExpandEnvVars(const std::wstring& str) {
    DWORD len = ExpandEnvironmentStringsW(str.c_str(), NULL, 0);
    if (len == 0)
        return str;
    std::vector<wchar_t> buf(len);
    ExpandEnvironmentStringsW(str.c_str(), buf.data(), len);
    return std::wstring(buf.data());
}

bool SaveFileAsUTF16LE(const wchar_t* filename, const std::wstring& text) {
    std::ofstream file(filename, std::ios::binary);
    if (!file.is_open())
        return false;

    // UTF-16 LE BOM
    unsigned char bom[] = { 0xFF, 0xFE };
    file.write(reinterpret_cast<const char*>(bom), sizeof(bom));

    if (false == text.empty()) {
        file.write(reinterpret_cast<const char*>(text.c_str()),
            text.size() * sizeof(wchar_t));
    }

    file.close();
    return true;
}

inline void ltrim(std::wstring& s) {
    s.erase(s.begin(), std::find_if(s.begin(), s.end(), [](wchar_t ch) {
        return !std::iswspace(ch);
        }));
}

inline void rtrim(std::wstring& s) {
    s.erase(std::find_if(s.rbegin(), s.rend(), [](wchar_t ch) {
        return !std::iswspace(ch);
        }).base(), s.end());
}

inline void trim(std::wstring& s) {
    ltrim(s);
    rtrim(s);
}

/// =============================================================
// ButtonInfo Dialog
int ShowButtonInfoDialog(int tabIdx, int btnIdx) {

    return (int)DialogBoxParam(
        GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_BUTTONINFO), g_hWnd,
        ButtonInfoDialogProc, (LPARAM)&g_buttonInfos[tabIdx][btnIdx]);
}

INT_PTR CALLBACK ButtonInfoDialogProc(HWND hDlg, UINT message, WPARAM wParam,
    LPARAM lParam) {
    ButtonInfo* pInfo = (ButtonInfo*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
    switch (message) {
    case WM_INITDIALOG:
        SetWindowLongPtr(hDlg, GWLP_USERDATA, lParam);
        if (lParam) {
            ButtonInfo* info = (ButtonInfo*)lParam;
            SetDlgItemTextW(hDlg, IDC_EDIT_NAME, info->name.c_str());
            SetDlgItemTextW(hDlg, IDC_EDIT_PATH, info->path.c_str());
            SetDlgItemTextW(hDlg, IDC_EDIT_PARAMS, info->parameters.c_str());
            CheckDlgButton(hDlg, IDC_CHECK_ADMIN,
                info->adminMode ? BST_CHECKED : BST_UNCHECKED);
        }

        SetWindowSubclass(GetDlgItem(hDlg, IDC_EDIT_NAME), EditSubclassProc, 1, 0);
        SetWindowSubclass(GetDlgItem(hDlg, IDC_EDIT_PATH), EditSubclassProc, 1, 0);
        SetWindowSubclass(GetDlgItem(hDlg, IDC_EDIT_PARAMS), EditSubclassProc, 1,
            0);

        return TRUE;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_BUTTON_BROWSE: {
            wchar_t fileName[MAX_PATH] = { 0 };
            OPENFILENAMEW ofn = { sizeof(ofn) };
            ofn.hwndOwner = hDlg;
            ofn.lpstrFilter = L"Executable Files (*.exe)\0*.exe\0All Files\0*.*\0";
            ofn.lpstrFile = fileName;
            ofn.nMaxFile = MAX_PATH;
            if (GetOpenFileNameW(&ofn)) {
                SetDlgItemTextW(hDlg, IDC_EDIT_PATH, fileName);
            }
            break;
        }
        case IDC_BUTTON_URL:
            ShellExecuteW(NULL, L"open", L"https://github.com/edgarp9/MultiTabLauncher", NULL, NULL, SW_SHOWNORMAL);
            break;
        case IDOK:
            if (pInfo) {
                wchar_t buf[512];
                GetDlgItemTextW(hDlg, IDC_EDIT_NAME, buf, 511);
                pInfo->name = buf;
                trim(pInfo->name);

                GetDlgItemTextW(hDlg, IDC_EDIT_PATH, buf, 511);
                pInfo->path = buf;
                trim(pInfo->path);

                GetDlgItemTextW(hDlg, IDC_EDIT_PARAMS, buf, 511);
                pInfo->parameters = buf;
                trim(pInfo->parameters);

                pInfo->adminMode =
                    (IsDlgButtonChecked(hDlg, IDC_CHECK_ADMIN) == BST_CHECKED);
            }
            EndDialog(hDlg, IDOK);
            break;
        case IDCANCEL:
            EndDialog(hDlg, IDCANCEL);
            break;
        }
        break;
    }
    return FALSE;
}

LRESULT CALLBACK EditSubclassProc(HWND hEdit, UINT msg, WPARAM wParam,
    LPARAM lParam, UINT_PTR, DWORD_PTR) {
    if (msg == WM_KEYDOWN) {
        if ((wParam == 'A' || wParam == 'a') &&
            (GetKeyState(VK_CONTROL) & 0x8000)) {
            SendMessageW(hEdit, EM_SETSEL, 0, -1);
            return 0;
        }
    }
    return DefSubclassProc(hEdit, msg, wParam, lParam);
}

/// =============================================================

std::wstring GetConfigDefault() {
    return LR"(

[Tabs]
Count=10
ButtonRows=3
ButtonCols=8
Tab0=Home
Tab1=System
Tab2=Dev
Tab3=Tab4
Tab4=Tab5
Tab5=Tab6
Tab6=Tab7
Tab7=Tab8
Tab8=Tab9
Tab9=Tab10
[Tab0]
Button0_Name=explorer
Button0_Path=explorer.exe
Button0_Params=
Button0_Admin=0
Button1_Name=notepad
Button1_Path=notepad.exe
Button1_Params=
Button1_Admin=0
Button2_Name=snippingtool
Button2_Path=snippingtool
Button2_Params=
Button2_Admin=0
Button3_Name=mspaint
Button3_Path=mspaint.exe
Button3_Params=
Button3_Admin=0
Button4_Name=calc
Button4_Path=calc.exe
Button4_Params=
Button4_Admin=0
Button5_Name=chrome
Button5_Path=C:\Program Files\Google\Chrome\Application\chrome.exe 
Button5_Params=
Button5_Admin=0
Button6_Name=
Button6_Path=
Button6_Params=
Button6_Admin=0
Button7_Name=
Button7_Path=
Button7_Params=
Button7_Admin=0
Button8_Name=Taskmgr
Button8_Path=Taskmgr.exe
Button8_Params=
Button8_Admin=0
Button9_Name=cmd
Button9_Path=cmd.exe
Button9_Params=
Button9_Admin=0
Button10_Name=mstsc
Button10_Path=mstsc
Button10_Params=
Button10_Admin=0
Button11_Name=Sandbox
Button11_Path=WindowsSandbox
Button11_Params=
Button11_Admin=0
Button12_Name=
Button12_Path=
Button12_Params=
Button12_Admin=0
Button13_Name=Desktop
Button13_Path=explorer.exe
Button13_Params=%USERPROFILE%\Desktop
Button13_Admin=0
Button14_Name=Downloads
Button14_Path=explorer.exe
Button14_Params=%USERPROFILE%\Downloads
Button14_Admin=0
Button15_Name=Run
Button15_Path=explorer.exe
Button15_Params=\_NDrv\rima\_run
Button15_Admin=0
[Tab1]
Button0_Name=hdwwiz
Button0_Path=control.exe
Button0_Params=hdwwiz.cpl
Button0_Admin=0
Button1_Name=control
Button1_Path=control
Button1_Params=
Button1_Admin=0
Button2_Name=diskmgmt
Button2_Path=diskmgmt.msc
Button2_Params=
Button2_Admin=0
Button3_Name=compmgmt
Button3_Path=compmgmt.msc
Button3_Params=
Button3_Admin=0
Button4_Name=appwiz
Button4_Path=control.exe
Button4_Params=appwiz.cpl
Button4_Admin=0
Button5_Name=msinfo32
Button5_Path=msinfo32
Button5_Params=
Button5_Admin=0
[Tab2]
Button0_Name=vs code
Button0_Path=code
Button0_Params=
Button0_Admin=0
Button1_Name=vs2022
Button1_Path=devenv.exe
Button1_Params=
Button1_Admin=0
Button2_Name=vs_installer
Button2_Path=C:\Program Files (x86)\Microsoft Visual Studio\Installer\setup.exe
Button2_Params=
Button2_Admin=0

)";
}
