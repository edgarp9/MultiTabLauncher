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
#include <chrono>
#include <thread>

#include "resource.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "Shlwapi.lib")

// =============================================================
//                   Global Constants and Variables
// =============================================================

const int MAX_TABS = 50;
const int BUTTON_ID_BASE = 1000;

// --- Application State ---
int g_tabCount = 0;
int g_currentTab = 0;
int g_buttonRows = 3;
int g_buttonCols = 8;
int g_buttonCountPerTab = g_buttonRows * g_buttonCols;

// --- Window and Path Information ---
LPCWSTR g_windowClassName = L"MultiTab Launcher";
const HICON g_hDefaultIcon = LoadIcon(NULL, IDI_APPLICATION);
std::wstring g_executableDirectory;
std::wstring g_configFilePath;

// --- Handles ---
HWND g_hMainWindow = NULL;
HWND g_hTabControl = NULL;

// --- Data Structures ---
struct ButtonInfo
{
    enum class ActionType { Launch, SendKeys };

    HWND hButton{ NULL };
    std::wstring name{ L"" };
    std::wstring path{ L"" };
    std::wstring parameters{ L"" };
    bool adminMode{ false };
    bool autoEnter{ false };
    HICON hIcon{ NULL };
    ActionType type{ ActionType::Launch };
};
std::vector<std::vector<ButtonInfo>> g_tabButtonData;
std::vector<std::wstring> g_tabNames;

// --- GDI Resources ---
HBRUSH g_hBackgroundBrush = NULL;
HBRUSH g_hTabBrush = NULL;
HBRUSH g_hButtonBrush = NULL;
HPEN g_hBorderPen = NULL;
HFONT g_hTabFont = NULL;


// =============================================================
//                   Function Prototypes
// =============================================================

// --- Window Procedures ---
LRESULT CALLBACK MainWindowProcedure(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK ButtonSettingsDialogProcedure(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK SelectAllEditSubclassProcedure(HWND hEdit, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR);

// --- Control Management ---
void InitializeTabControl(HWND hwnd);
void CreateButtonsForTab(HWND hwnd, int tabIndex);
void UpdateLayoutOnResize(HWND hwnd);

// --- GDI Resource Management ---
void InitializeGdiResources();
void ReleaseGdiResources();

// --- Window State Persistence ---
void SaveWindowPosition(HWND hwnd);
void RestoreWindowPosition(HWND hwnd);

// --- Configuration (INI File) Handling ---
void LoadConfigurationFromFile();
bool SaveButtonConfigurationToFile(int tabIndex, int buttonIndex, const ButtonInfo& info);
bool GenerateDefaultConfigFile();
std::wstring GetDefaultConfigString();
std::wstring ReadStringFromIni(LPCWSTR section, LPCWSTR key, LPCWSTR defaultValue, LPCWSTR filePath);
bool WriteUtf16LeFile(const wchar_t* filename, const std::wstring& text);

// --- Core Application Logic ---
bool LaunchApplication(const std::wstring& filePath, const std::wstring& parameters, bool asAdmin);
void OnLaunchButtonClick(int tabIndex, int buttonIndex);
bool SendTextToWindow(const std::wstring& command, bool isAutoEnter);
int DisplayButtonSettingsDialog(int tabIdx, int btnIdx);

// --- Utility Functions ---
HICON ExtractIconFromFile(const std::wstring& filePath);
std::wstring ResolveExecutablePath(const wchar_t* targetFile);
std::wstring ExpandEnvironmentVariables(const std::wstring& str);
std::wstring GetTextFromDialogControl(HWND hDlg, int nCtlId);
inline void trim(std::wstring& s);
inline void rtrim(std::wstring& s);
inline void ltrim(std::wstring& s);


// =============================================================
//                   WinMain - Entry Point
// =============================================================

// =============================================================
//                   WinMain - Entry Point
// =============================================================

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow)
{
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr))
    {
    }

    // Initialize common controls (for the tab control)
    INITCOMMONCONTROLSEX icc = { sizeof(INITCOMMONCONTROLSEX), ICC_TAB_CLASSES };
    InitCommonControlsEx(&icc);

    // Get the directory of the executable to resolve relative paths
    std::wstring modulePath(MAX_PATH, L'\0');
    DWORD modulePathLength = 0;
    do {
        modulePathLength = GetModuleFileNameW(NULL, modulePath.data(), static_cast<DWORD>(modulePath.length()));
        if (modulePathLength == 0)
        {
            MessageBox(NULL, L"Failed to get the executable path.", L"Error", MB_OK | MB_ICONERROR);
            CoUninitialize(); 
            return 1;
        }
        if (modulePathLength < modulePath.length())
        {
            modulePath.resize(modulePathLength);
            break;
        }
        modulePath.resize(modulePath.length() * 2);
    } while (true);

    // Set current directory to the executable's directory and define INI file path
    std::filesystem::path exePath(modulePath);
    g_executableDirectory = exePath.parent_path().wstring();
    SetCurrentDirectoryW(g_executableDirectory.c_str());
    g_configFilePath = g_executableDirectory + L"\\MultiTabLauncher.ini";

    // Load configuration from INI and initialize GDI resources
    LoadConfigurationFromFile();
    InitializeGdiResources();

    // Register the window class
    WNDCLASSEX wc = { sizeof(WNDCLASSEX) };
    wc.lpfnWndProc = MainWindowProcedure;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = g_hBackgroundBrush;
    wc.lpszClassName = g_windowClassName;
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));
    wc.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));

    if (!RegisterClassEx(&wc))
    {
        MessageBox(NULL, L"Window Registration Failed!", L"Error", MB_OK | MB_ICONERROR);
        CoUninitialize();
        return 1;
    }

    // Create the main window
    g_hMainWindow = CreateWindowEx(
        WS_EX_CLIENTEDGE, g_windowClassName, L"MultiTab Launcher v3.1",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 800,
        600, NULL, NULL, hInstance, NULL
    );

    if (g_hMainWindow == NULL)
    {
        MessageBox(NULL, L"Window Creation Failed!", L"Error", MB_OK | MB_ICONERROR);
        CoUninitialize();
        return 1;
    }

    ShowWindow(g_hMainWindow, nCmdShow);
    UpdateWindow(g_hMainWindow);

    // Main message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    CoUninitialize();

    return (int)msg.wParam;
}

// =============================================================
//               MainWindowProcedure - Main Window Procedure
// =============================================================

LRESULT CALLBACK MainWindowProcedure(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
    {
        RestoreWindowPosition(hwnd);
        InitializeTabControl(hwnd);
        // Show buttons for the initially selected tab
        for (int i = 0; i < g_buttonCountPerTab; i++)
        {
            ShowWindow(g_tabButtonData[0][i].hButton, SW_SHOW);
        }
        break;
    }

    case WM_ACTIVATE:
    {
        if (LOWORD(wParam) != WA_INACTIVE)
        {
            InvalidateRect(hwnd, NULL, FALSE);
        }
        return 0;
    }

    case WM_SIZE:
    {
        UpdateLayoutOnResize(hwnd);
        break;
    }

    case WM_NOTIFY:
    {
        LPNMHDR nmhdr = (LPNMHDR)lParam;
        if (nmhdr->hwndFrom == g_hTabControl && nmhdr->code == TCN_SELCHANGE)
        {
            int newTab = TabCtrl_GetCurSel(g_hTabControl);
            if (newTab != g_currentTab)
            {
                // Hide buttons of the old tab
                for (int i = 0; i < g_buttonCountPerTab; i++)
                {
                    ShowWindow(g_tabButtonData[g_currentTab][i].hButton, SW_HIDE);
                }
                // Show buttons of the new tab
                for (int i = 0; i < g_buttonCountPerTab; i++)
                {
                    ShowWindow(g_tabButtonData[newTab][i].hButton, SW_SHOW);
                }
                g_currentTab = newTab;
                InvalidateRect(hwnd, NULL, FALSE); // Redraw window
            }
        }
        break;
    }

    case WM_COMMAND:
    {
        int buttonId = LOWORD(wParam);
        if (buttonId >= BUTTON_ID_BASE && buttonId < BUTTON_ID_BASE + (g_tabCount * g_buttonCountPerTab))
        {
            int tabIndex = (buttonId - BUTTON_ID_BASE) / g_buttonCountPerTab;
            int buttonIndex = (buttonId - BUTTON_ID_BASE) % g_buttonCountPerTab;
            OnLaunchButtonClick(tabIndex, buttonIndex);
        }
        break;
    }

    case WM_CONTEXTMENU:
    {
        // Handle right-click on a button to open the settings dialog
        HWND hCtrl = (HWND)wParam;
        for (int tab = 0; tab < g_tabCount; ++tab)
        {
            for (int btn = 0; btn < g_buttonCountPerTab; ++btn)
            {
                if (hCtrl == g_tabButtonData[tab][btn].hButton)
                {
                    if (DisplayButtonSettingsDialog(tab, btn) == IDOK)
                    {
                        // Update button text and icon after dialog closes
                        ButtonInfo& info = g_tabButtonData[tab][btn];
                        SetWindowTextW(info.hButton, info.name.c_str());

                        HICON hOldIcon = info.hIcon;
                        if (info.type == ButtonInfo::ActionType::Launch) {
                            info.hIcon = info.path.empty() ? NULL : ExtractIconFromFile(info.path);
                        }
                        else {
                            info.hIcon = NULL;
                        }
                        if (hOldIcon && hOldIcon != g_hDefaultIcon)
                        {
                            DestroyIcon(hOldIcon);
                        }

                        InvalidateRect(hCtrl, NULL, TRUE);
                        SaveButtonConfigurationToFile(tab, btn, info);
                        SetCurrentDirectoryW(g_executableDirectory.c_str());
                    }
                }
            }
        }
        break;
    }

    case WM_CTLCOLORBTN:
    {
        // Set custom colors for buttons (used as a base for owner-draw)
        SetTextColor((HDC)wParam, RGB(255, 255, 255));
        SetBkColor((HDC)wParam, RGB(45, 45, 45));
        return (LRESULT)g_hButtonBrush;
    }

    case WM_DRAWITEM:
    {
        // Custom drawing for buttons (BS_OWNERDRAW style)
        LPDRAWITEMSTRUCT pDIS = (LPDRAWITEMSTRUCT)lParam;
        if (pDIS->CtlType == ODT_BUTTON)
        {
            // Determine button state and set background color
            bool isSelected = pDIS->itemState & ODS_SELECTED;
            HBRUSH brush = isSelected ? CreateSolidBrush(RGB(9, 71, 113)) : g_hButtonBrush;
            FillRect(pDIS->hDC, &pDIS->rcItem, brush);
            if (isSelected) DeleteObject(brush);

            // Draw border
            SelectObject(pDIS->hDC, g_hBorderPen);
            SelectObject(pDIS->hDC, GetStockObject(NULL_BRUSH));
            Rectangle(pDIS->hDC, pDIS->rcItem.left, pDIS->rcItem.top, pDIS->rcItem.right, pDIS->rcItem.bottom);

            // Get button info
            int buttonId = GetDlgCtrlID(pDIS->hwndItem);
            int tabIndex = (buttonId - BUTTON_ID_BASE) / g_buttonCountPerTab;
            int btnIndex = (buttonId - BUTTON_ID_BASE) % g_buttonCountPerTab;
            const ButtonInfo& btnInfo = g_tabButtonData[tabIndex][btnIndex];

            // Get button text
            WCHAR text[256];
            GetWindowText(pDIS->hwndItem, text, 256);
            SIZE textSize{};
            GetTextExtentPoint32(pDIS->hDC, text, lstrlenW(text), &textSize);

            // Calculate vertical alignment for icon and text
            const int iconSize = 32;
            const int spaceBetweenIconAndText = 8;
            int totalHeight = (btnInfo.hIcon ? iconSize + spaceBetweenIconAndText : 0) + textSize.cy;
            int startY = pDIS->rcItem.top + (pDIS->rcItem.bottom - pDIS->rcItem.top - totalHeight) / 2;

            // Draw icon
            if (btnInfo.hIcon)
            {
                int iconX = pDIS->rcItem.left + (pDIS->rcItem.right - pDIS->rcItem.left - iconSize) / 2;
                DrawIconEx(pDIS->hDC, iconX, startY, btnInfo.hIcon, iconSize, iconSize, 0, NULL, DI_NORMAL);
            }

            // Draw text
            SetTextColor(pDIS->hDC, RGB(204, 204, 204));
            SetBkMode(pDIS->hDC, TRANSPARENT);
            RECT rcText = pDIS->rcItem;
            rcText.top = startY + (btnInfo.hIcon ? iconSize + spaceBetweenIconAndText : 0);
            rcText.bottom = rcText.top + textSize.cy;
            DrawText(pDIS->hDC, text, -1, &rcText, DT_CENTER | DT_TOP | DT_SINGLELINE | DT_END_ELLIPSIS);

            return TRUE;
        }
        break;
    }

    case WM_ERASEBKGND:
    {
        // Prevent background flicker by handling all painting in WM_PAINT
        return 1;
    }

    case WM_PAINT:
    {
        // Use double-buffering to prevent flicker during resize or redraw
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);

        RECT rc;
        GetClientRect(hwnd, &rc);

        HDC memDC = CreateCompatibleDC(hdc);
        HBITMAP bmp = CreateCompatibleBitmap(hdc, rc.right, rc.bottom);
        HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, bmp);

        // Draw background
        FillRect(memDC, &rc, g_hBackgroundBrush);

        // Draw the tab control into the memory DC. This is key for achieving
        // a transparent-like effect for the tab control's background.
        SendMessage(g_hTabControl, WM_PRINTCLIENT, (WPARAM)memDC, (LPARAM)(PRF_ERASEBKGND | PRF_CLIENT | PRF_CHILDREN));

        // Copy the completed image from the memory DC to the screen
        BitBlt(hdc, 0, 0, rc.right, rc.bottom, memDC, 0, 0, SRCCOPY);

        SelectObject(memDC, oldBmp);
        DeleteObject(bmp);
        DeleteDC(memDC);

        EndPaint(hwnd, &ps);
        break;
    }

    case WM_DESTROY:
    {
        SaveWindowPosition(hwnd);
        ReleaseGdiResources();
        PostQuitMessage(0);
        break;
    }

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// =============================================================
//                   GDI Resource Management
// =============================================================

/**
 * @brief Initializes GDI resources like brushes, pens, and fonts.
 */
void InitializeGdiResources()
{
    g_hBackgroundBrush = CreateSolidBrush(RGB(30, 30, 30));
    g_hTabBrush = CreateSolidBrush(RGB(37, 37, 38));
    g_hButtonBrush = CreateSolidBrush(RGB(60, 60, 60));
    g_hBorderPen = CreatePen(PS_SOLID, 1, RGB(50, 50, 50));

    LOGFONT lf = {};
    lf.lfHeight = 24;
    lf.lfWeight = FW_BOLD;
    g_hTabFont = CreateFontIndirect(&lf);
}

/**
 * @brief Cleans up all allocated GDI resources and icons.
 */
void ReleaseGdiResources()
{
    // Destroy all loaded button icons
    for (auto& tab : g_tabButtonData)
    {
        for (auto& buttonInfo : tab)
        {
            if (buttonInfo.hIcon && buttonInfo.hIcon != g_hDefaultIcon)
            {
                DestroyIcon(buttonInfo.hIcon);
            }
        }
    }

    // Delete GDI objects
    DeleteObject(g_hBackgroundBrush);
    DeleteObject(g_hTabBrush);
    DeleteObject(g_hButtonBrush);
    DeleteObject(g_hBorderPen);
    DeleteObject(g_hTabFont);
}

// =============================================================
//                   Control Creation and Resizing
// =============================================================

/**
 * @brief Creates the main tab control and populates it with tabs.
 * @param hwnd Handle to the parent window.
 */
void InitializeTabControl(HWND hwnd)
{
    g_hTabControl = CreateWindowEx(
        0, WC_TABCONTROL, L"",
        WS_CHILD | WS_VISIBLE | TCS_FLATBUTTONS | TCS_BUTTONS, 0,
        0, 0, 0, hwnd, NULL, GetModuleHandle(NULL), NULL
    );

    SendMessage(g_hTabControl, WM_SETFONT, (WPARAM)g_hTabFont, TRUE);

    // Insert tabs and create corresponding buttons for each tab
    TCITEM tie = { TCIF_TEXT };
    for (int i = 0; i < g_tabCount; ++i)
    {
        tie.pszText = (LPWSTR)g_tabNames[i].c_str();
        TabCtrl_InsertItem(g_hTabControl, i, &tie);
        CreateButtonsForTab(hwnd, i);
    }
}

/**
 * @brief Creates the grid of buttons for a specific tab.
 * @param hwnd Handle to the parent window.
 * @param tabIndex The index of the tab for which to create buttons.
 */
void CreateButtonsForTab(HWND hwnd, int tabIndex)
{
    RECT rcTab;
    GetClientRect(g_hTabControl, &rcTab);
    TabCtrl_AdjustRect(g_hTabControl, FALSE, &rcTab);

    int btnWidth = (rcTab.right - rcTab.left) / g_buttonCols;
    int btnHeight = (rcTab.bottom - rcTab.top) / g_buttonRows;

    for (int row = 0; row < g_buttonRows; ++row)
    {
        for (int col = 0; col < g_buttonCols; ++col)
        {
            int i = row * g_buttonCols + col;
            int id = BUTTON_ID_BASE + tabIndex * g_buttonCountPerTab + i;
            g_tabButtonData[tabIndex][i].hButton = CreateWindowEx(
                0, L"BUTTON", g_tabButtonData[tabIndex][i].name.c_str(),
                WS_CHILD | BS_PUSHBUTTON | BS_OWNERDRAW,
                rcTab.left + col * btnWidth, rcTab.top + row * btnHeight, btnWidth, btnHeight,
                hwnd, (HMENU)(INT_PTR)id, GetModuleHandle(NULL), NULL
            );
        }
    }
}

/**
 * @brief Resizes the tab control and all buttons when the main window is resized.
 * @param hwnd Handle to the main window.
 */
void UpdateLayoutOnResize(HWND hwnd)
{
    RECT rcClient;
    GetClientRect(hwnd, &rcClient);
    MoveWindow(g_hTabControl, 0, 0, rcClient.right, rcClient.bottom, TRUE);

    RECT rcTab;
    GetClientRect(g_hTabControl, &rcTab);
    TabCtrl_AdjustRect(g_hTabControl, FALSE, &rcTab);

    int btnWidth = (rcTab.right - rcTab.left) / g_buttonCols;
    int btnHeight = (rcTab.bottom - rcTab.top) / g_buttonRows;

    for (int tab = 0; tab < g_tabCount; ++tab)
    {
        for (int i = 0; i < g_buttonCountPerTab; ++i)
        {
            if (g_tabButtonData[tab][i].hButton)
            {
                int row = i / g_buttonCols;
                int col = i % g_buttonCols;
                MoveWindow(g_tabButtonData[tab][i].hButton, rcTab.left + col * btnWidth,
                    rcTab.top + row * btnHeight, btnWidth, btnHeight, TRUE);
            }
        }
    }
    InvalidateRect(hwnd, NULL, FALSE);
}

// =============================================================
//               Configuration (INI File) Handling
// =============================================================

/**
 * @brief Creates the INI file with default settings if it doesn't exist.
 * @return True if the file was created successfully, false otherwise.
 */
bool GenerateDefaultConfigFile()
{
    return WriteUtf16LeFile(g_configFilePath.c_str(), GetDefaultConfigString());
}

/**
 * @brief Safely reads a string from the INI file, handling buffer resizing.
 * @param section The section name in the INI file.
 * @param key The key name in the section.
 * @param defaultValue The default value to return if the key is not found.
 * @param filePath The path to the INI file.
 * @return The retrieved string value.
 */
std::wstring ReadStringFromIni(LPCWSTR section, LPCWSTR key, LPCWSTR defaultValue, LPCWSTR filePath)
{
    std::vector<wchar_t> buffer(256);
    while (true)
    {
        DWORD charsCopied = GetPrivateProfileStringW(
            section, key, defaultValue, buffer.data(), static_cast<DWORD>(buffer.size()), filePath
        );
        if (charsCopied < buffer.size() - 1)
        {
            return std::wstring(buffer.data());
        }
        buffer.resize(buffer.size() * 2);
    }
}

/**
 * @brief Loads all configuration settings from the INI file.
 */
void LoadConfigurationFromFile()
{
    // If INI file doesn't exist, create it with default values
    if (!PathFileExists(g_configFilePath.c_str()))
    {
        if (!GenerateDefaultConfigFile())
        {
            MessageBox(NULL, L"Failed to create config.ini file.", L"Error", MB_OK | MB_ICONERROR);
            std::exit(1);
        }
    }

    // Read general settings
    g_tabCount = GetPrivateProfileInt(L"Tabs", L"Count", 10, g_configFilePath.c_str());
    if (g_tabCount <= 0 || g_tabCount > MAX_TABS) g_tabCount = 10;

    g_buttonRows = GetPrivateProfileInt(L"Tabs", L"ButtonRows", 3, g_configFilePath.c_str());
    if (g_buttonRows <= 0) g_buttonRows = 3;

    g_buttonCols = GetPrivateProfileInt(L"Tabs", L"ButtonCols", 8, g_configFilePath.c_str());
    if (g_buttonCols <= 0) g_buttonCols = 8;

    g_buttonCountPerTab = g_buttonRows * g_buttonCols;

    // Resize data structures
    g_tabNames.resize(g_tabCount);
    g_tabButtonData.resize(g_tabCount, std::vector<ButtonInfo>(g_buttonCountPerTab));

    // Read tab names
    for (int i = 0; i < g_tabCount; i++)
    {
        std::wstring key = L"Tab" + std::to_wstring(i);
        std::wstring defaultName = L"Tab" + std::to_wstring(i + 1);
        g_tabNames[i] = ReadStringFromIni(L"Tabs", key.c_str(), defaultName.c_str(), g_configFilePath.c_str());
        if (g_tabNames[i].empty() || g_tabNames[i].length() > 30)
        {
            g_tabNames[i] = defaultName;
        }
    }

    // Read button information for each tab
    for (int tab = 0; tab < g_tabCount; tab++)
    {
        std::wstring section = L"Tab" + std::to_wstring(tab);
        for (int btn = 0; btn < g_buttonCountPerTab; btn++)
        {
            std::wstring keyBase = L"Button" + std::to_wstring(btn);
            ButtonInfo& info = g_tabButtonData[tab][btn];

            info.name = ReadStringFromIni(section.c_str(), (keyBase + L"_Name").c_str(), L"", g_configFilePath.c_str());
            trim(info.name);

            info.path = ReadStringFromIni(section.c_str(), (keyBase + L"_Path").c_str(), L"", g_configFilePath.c_str());
            trim(info.path);

            info.parameters = ReadStringFromIni(section.c_str(), (keyBase + L"_Params").c_str(), L"", g_configFilePath.c_str());
            trim(info.parameters);

            info.adminMode = GetPrivateProfileInt(section.c_str(), (keyBase + L"_Admin").c_str(), 0, g_configFilePath.c_str()) != 0;
            
            info.autoEnter = GetPrivateProfileInt(section.c_str(), (keyBase + L"_AutoEnter").c_str(), 0, g_configFilePath.c_str()) != 0;

            int action = GetPrivateProfileInt(section.c_str(), (keyBase + L"_Action").c_str(), 0, g_configFilePath.c_str());
            info.type = (action == 1) ? ButtonInfo::ActionType::SendKeys : ButtonInfo::ActionType::Launch;

            if (info.type == ButtonInfo::ActionType::Launch) {
                info.hIcon = info.path.empty() ? NULL : ExtractIconFromFile(info.path);
            }
            else {
                info.hIcon = NULL;
            }
        }
    }
}

/**
 * @brief Saves the information for a single button to the INI file.
 * @param tabIndex The tab index of the button.
 * @param buttonIndex The index of the button within the tab.
 * @param info The ButtonInfo structure containing the data to save.
 * @return True on success, false on failure.
 */
bool SaveButtonConfigurationToFile(int tabIndex, int buttonIndex, const ButtonInfo& info)
{
    std::wstring section = L"Tab" + std::to_wstring(tabIndex);
    std::wstring btnKey = L"Button" + std::to_wstring(buttonIndex);

    BOOL ok = TRUE;
    ok &= WritePrivateProfileStringW(section.c_str(), (btnKey + L"_Name").c_str(), info.name.c_str(), g_configFilePath.c_str());
    ok &= WritePrivateProfileStringW(section.c_str(), (btnKey + L"_Path").c_str(), info.path.c_str(), g_configFilePath.c_str());
    ok &= WritePrivateProfileStringW(section.c_str(), (btnKey + L"_Params").c_str(), info.parameters.c_str(), g_configFilePath.c_str());
    ok &= WritePrivateProfileStringW(section.c_str(), (btnKey + L"_Admin").c_str(), info.adminMode ? L"1" : L"0", g_configFilePath.c_str());
    
    std::wstring action = (info.type == ButtonInfo::ActionType::SendKeys) ? L"1" : L"0";
    ok &= WritePrivateProfileStringW(section.c_str(), (btnKey + L"_Action").c_str(), action.c_str(), g_configFilePath.c_str());
    
    ok &= WritePrivateProfileStringW(section.c_str(), (btnKey + L"_AutoEnter").c_str(), info.autoEnter ? L"1" : L"0", g_configFilePath.c_str());
    
    return ok != 0;
}

// =============================================================
//                   Window State Persistence
// =============================================================

/**
 * @brief Saves the main window's position and size to the registry.
 * @param hwnd Handle to the main window.
 */
void SaveWindowPosition(HWND hwnd)
{
    WINDOWPLACEMENT wp = { sizeof(WINDOWPLACEMENT) };
    if (GetWindowPlacement(hwnd, &wp))
    {
        HKEY hKey;
        if (RegCreateKeyEx(HKEY_CURRENT_USER, L"Software\\MultiTabLauncher", 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
        {
            RegSetValueEx(hKey, L"Placement", 0, REG_BINARY, (BYTE*)&wp, sizeof(wp));
            RegCloseKey(hKey);
        }
    }
}

/**
 * @brief Loads the main window's position and size from the registry.
 * @param hwnd Handle to the main window.
 */
void RestoreWindowPosition(HWND hwnd)
{
    WINDOWPLACEMENT wp = { sizeof(WINDOWPLACEMENT) };
    HKEY hKey;
    DWORD size = sizeof(wp);
    if (RegOpenKeyEx(HKEY_CURRENT_USER, L"Software\\MultiTabLauncher", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        if (RegQueryValueEx(hKey, L"Placement", NULL, NULL, (LPBYTE)&wp, &size) == ERROR_SUCCESS)
        {
            wp.showCmd = SW_SHOWNORMAL; // Ensure window is shown normally on restore
            SetWindowPlacement(hwnd, &wp);
        }
        RegCloseKey(hKey);
    }
}


// =============================================================
//                        Core Application Logic
// =============================================================

/**
 * @brief Executes a process using ShellExecuteW, with fallback logic.
 * @param filePath Path to the executable or document.
 * @param parameters Command-line parameters.
 * @param asAdmin True to run the process with administrator privileges.
 * @return True on success, false on failure.
 */
bool LaunchApplication(const std::wstring& filePath, const std::wstring& parameters, bool asAdmin)
{
    std::wstring operation = asAdmin ? L"runas" : L"open";

    // Expand environment variables (e.g., %USERPROFILE%) before execution
    std::wstring expandedPath = ExpandEnvironmentVariables(filePath);
    std::wstring expandedParams = ExpandEnvironmentVariables(parameters);

    HINSTANCE result = ShellExecuteW(
        NULL, operation.c_str(), expandedPath.c_str(),
        expandedParams.empty() ? NULL : expandedParams.c_str(),
        NULL, SW_SHOWNORMAL
    );

    if ((INT_PTR)result > 32)
    {
        return true; // Success
    }

    // If execution fails, try to find the executable's absolute path and retry
    std::wstring absPath = ResolveExecutablePath(expandedPath.c_str());
    if (!absPath.empty() && absPath != expandedPath)
    {
        result = ShellExecuteW(
            NULL, operation.c_str(), absPath.c_str(),
            expandedParams.empty() ? NULL : expandedParams.c_str(),
            NULL, SW_SHOWNORMAL
        );
        if ((INT_PTR)result > 32)
        {
            return true; // Success on retry
        }
    }

    // Display a detailed error message on failure
    std::wstring msg = L"ShellExecuteW failed.\n";
    msg += L"File: " + filePath + L"\n";
    msg += L"Error code: " + std::to_wstring((INT_PTR)result) + L"\n\n";

    switch ((INT_PTR)result)
    {
    case 0: msg += L"The operating system is out of memory or resources."; break;
    case ERROR_FILE_NOT_FOUND: msg += L"The specified file was not found."; break;
    case ERROR_PATH_NOT_FOUND: msg += L"The specified path was not found."; break;
    case ERROR_BAD_FORMAT: msg += L"The .exe file is invalid."; break;
    case SE_ERR_ACCESSDENIED: msg += L"Access is denied."; break;
    case SE_ERR_ASSOCINCOMPLETE: msg += L"The filename association is incomplete or invalid."; break;
    case SE_ERR_NOASSOC: msg += L"There is no application associated with the given file name extension."; break;
    case SE_ERR_OOM: msg += L"Not enough memory to complete the operation."; break;
    default: msg += L"An unknown error occurred."; break;
    }
    MessageBoxW(NULL, msg.c_str(), L"Execution Error", MB_OK | MB_ICONERROR);

    return false;
}

/**
 * @brief Handles the click event for a launch button.
 * @param tabIndex The index of the tab containing the clicked button.
 * @param buttonIndex The index of the clicked button.
 */
void OnLaunchButtonClick(int tabIndex, int buttonIndex)
{
    const ButtonInfo& buttonInfo = g_tabButtonData[tabIndex][buttonIndex];

    if (buttonInfo.type == ButtonInfo::ActionType::SendKeys)
    {
        if (!buttonInfo.parameters.empty())
        {
            SendTextToWindow(buttonInfo.parameters, buttonInfo.autoEnter);
        }
    }
    else // Default is Launch
    {
        if (!buttonInfo.path.empty())
        {
            LaunchApplication(buttonInfo.path, buttonInfo.parameters, buttonInfo.adminMode);
            SetCurrentDirectoryW(g_executableDirectory.c_str());
        }
    }
}

bool SendTextToWindow(const std::wstring& command, bool isAutoEnter)
{
    HWND hTarget = nullptr;
    HWND hWnd = GetTopWindow(nullptr);
    while (hWnd != nullptr) {
        if (hWnd != g_hMainWindow && IsWindowVisible(hWnd) && GetWindowTextLength(hWnd) > 0) {
            hTarget = hWnd;
            break;
        }
        hWnd = GetNextWindow(hWnd, GW_HWNDNEXT);
    }

    if (hTarget == nullptr) {
        return false;
    }

    SetForegroundWindow(hTarget);
    BringWindowToTop(hTarget);

    std::this_thread::sleep_for(std::chrono::milliseconds(200)); 

    std::vector<INPUT> inputs;
    inputs.reserve(command.length() * 2 + (isAutoEnter ? 2 : 0)); 

    for (wchar_t wch : command) {
        // Key Down
        INPUT input_down{};
        input_down.type = INPUT_KEYBOARD;
        input_down.ki.wScan = wch;
        input_down.ki.dwFlags = KEYEVENTF_UNICODE;
        inputs.push_back(input_down);

        // Key Up
        INPUT input_up{};
        input_up.type = INPUT_KEYBOARD;
        input_up.ki.wScan = wch;
        input_up.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
        inputs.push_back(input_up);
    }

    if (isAutoEnter) {
        // Key Down
        INPUT enter_down{};
        enter_down.type = INPUT_KEYBOARD;
        enter_down.ki.wVk = VK_RETURN;
        inputs.push_back(enter_down);

        // Key Up
        INPUT enter_up{};
        enter_up.type = INPUT_KEYBOARD;
        enter_up.ki.wVk = VK_RETURN;
        enter_up.ki.dwFlags = KEYEVENTF_KEYUP;
        inputs.push_back(enter_up);
    }

    if (!inputs.empty()) {
        const UINT sent_count = SendInput(static_cast<UINT>(inputs.size()), inputs.data(), sizeof(INPUT));
        return sent_count == inputs.size();
    }

    return true; 
}


// =============================================================
//                     Utility Functions
// =============================================================

/**
 * @brief Extracts the large icon associated with a file.
 * @param filePath Path to the file (can be relative or absolute).
 * @return HICON handle to the extracted icon, or a default icon on failure.
 */
HICON ExtractIconFromFile(const std::wstring& filePath)
{
    if (filePath.empty()) return NULL;

    std::wstring pathToIcon = filePath;
    // If the path is relative, try to find its full path to help SHGetFileInfo
    if (PathIsRelativeW(pathToIcon.c_str()))
    {
        pathToIcon = ResolveExecutablePath(filePath.c_str());
    }

    if (!pathToIcon.empty())
    {
        SHFILEINFOW sfi = {};
        if (SHGetFileInfoW(pathToIcon.c_str(), 0, &sfi, sizeof(sfi), SHGFI_ICON | SHGFI_LARGEICON))
        {
            return sfi.hIcon;
        }
    }
    // Fallback to a copy of the default application icon if extraction fails
    return (HICON)CopyIcon(g_hDefaultIcon);
}

/**
 * @brief Searches for an executable in the app's directory and system PATH.
 * @param targetFile The name of the file to find.
 * @return The full path to the file if found, otherwise the original file name.
 */
std::wstring ResolveExecutablePath(const wchar_t* targetFile)
{
    // 1. Check in the application's own directory first
    if (!g_executableDirectory.empty())
    {
        std::filesystem::path targetPath = std::filesystem::path(g_executableDirectory) / targetFile;
        if (std::filesystem::exists(targetPath))
        {
            return targetPath.wstring();
        }
    }

    // 2. Search using the system's PATH environment variables
    std::wstring buffer(MAX_PATH, L'\0');
    do {
        DWORD pathLen = SearchPathW(NULL, targetFile, NULL, static_cast<DWORD>(buffer.length()), buffer.data(), NULL);
        if (pathLen == 0)
        {
            break; // Not found
        }
        if (pathLen < buffer.length())
        {
            buffer.resize(pathLen);
            return buffer;
        }
        buffer.resize(buffer.length() * 2);
    } while (true);

    // 3. Fallback to the original name. ShellExecute might still find it.
    return std::wstring(targetFile);
}

/**
 * @brief Expands environment variables in a string (e.g., %USERPROFILE%).
 * @param str The input string.
 * @return The expanded string.
 */
std::wstring ExpandEnvironmentVariables(const std::wstring& str)
{
    if (str.find(L'%') == std::wstring::npos) return str;

    DWORD len = ExpandEnvironmentStringsW(str.c_str(), NULL, 0);
    if (len == 0) return str;

    std::vector<wchar_t> buf(len);
    ExpandEnvironmentStringsW(str.c_str(), buf.data(), len);
    return std::wstring(buf.data());
}

/**
 * @brief Saves a wstring to a file with UTF-16 LE encoding and a BOM.
 * @param filename The name of the file to save.
 * @param text The content to write.
 * @return True on success, false on failure.
 */
bool WriteUtf16LeFile(const wchar_t* filename, const std::wstring& text)
{
    std::ofstream file(filename, std::ios::binary);
    if (!file.is_open()) return false;

    // Write UTF-16 LE Byte Order Mark (BOM)
    unsigned char bom[] = { 0xFF, 0xFE };
    file.write(reinterpret_cast<const char*>(bom), sizeof(bom));

    if (!text.empty())
    {
        file.write(reinterpret_cast<const char*>(text.c_str()), text.size() * sizeof(wchar_t));
    }

    file.close();
    return true;
}

// --- String Trimming Utilities ---
inline void ltrim(std::wstring& s)
{
    s.erase(s.begin(), std::find_if(s.begin(), s.end(), [](wchar_t ch)
        {
            return !std::iswspace(ch);
        }));
}

inline void rtrim(std::wstring& s)
{
    s.erase(std::find_if(s.rbegin(), s.rend(), [](wchar_t ch)
        {
            return !std::iswspace(ch);
        }).base(), s.end());
}

inline void trim(std::wstring& s)
{
    ltrim(s);
    rtrim(s);
}

// =============================================================
//               Button Settings Dialog and Helpers
// =============================================================

void UpdateDialogControls(HWND hDlg)
{
    BOOL isLaunch = (IsDlgButtonChecked(hDlg, IDC_RADIO_LAUNCH) == BST_CHECKED);

    // 'Launch Program' 관련 컨트롤들
    ShowWindow(GetDlgItem(hDlg, IDC_LBL_PATH), isLaunch ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(hDlg, IDC_EDIT_PATH), isLaunch ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(hDlg, IDC_BUTTON_BROWSE), isLaunch ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(hDlg, IDC_LBL_PARAMS), isLaunch ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(hDlg, IDC_EDIT_PARAMS), isLaunch ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(hDlg, IDC_CHECK_ADMIN), isLaunch ? SW_SHOW : SW_HIDE);

    // 'Send Keystrokes' 관련 컨트롤들
    ShowWindow(GetDlgItem(hDlg, IDC_LBL_COMMAND), !isLaunch ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(hDlg, IDC_EDIT_COMMAND), !isLaunch ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(hDlg, IDC_CHECK_AUTOENTER), !isLaunch ? SW_SHOW : SW_HIDE);
}

/**
 * @brief Displays the modal dialog to edit button information.
 * @param tabIdx The tab index of the button.
 * @param btnIdx The index of the button.
 * @return IDOK if the user clicked OK, IDCANCEL otherwise.
 */
int DisplayButtonSettingsDialog(int tabIdx, int btnIdx)
{
    return (int)DialogBoxParam(
        GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_BUTTONINFO), g_hMainWindow,
        ButtonSettingsDialogProcedure, (LPARAM)&g_tabButtonData[tabIdx][btnIdx]
    );
}

/**
 * @brief Dialog procedure for the button information editor.
 */
INT_PTR CALLBACK ButtonSettingsDialogProcedure(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
    {
        ButtonInfo* pInfo = (ButtonInfo*)lParam;
        SetWindowLongPtr(hDlg, GWLP_USERDATA, (LONG_PTR)pInfo);

        if (pInfo)
        {
            SetDlgItemTextW(hDlg, IDC_EDIT_NAME, pInfo->name.c_str());
            if (pInfo->type == ButtonInfo::ActionType::SendKeys)
            {
                CheckDlgButton(hDlg, IDC_RADIO_SENDKEYS, BST_CHECKED);
                SetDlgItemTextW(hDlg, IDC_EDIT_COMMAND, pInfo->parameters.c_str());
                CheckDlgButton(hDlg, IDC_CHECK_AUTOENTER, pInfo->autoEnter ? BST_CHECKED : BST_UNCHECKED);
            }
            else // Default to Launch
            {
                CheckDlgButton(hDlg, IDC_RADIO_LAUNCH, BST_CHECKED);
                SetDlgItemTextW(hDlg, IDC_EDIT_PATH, pInfo->path.c_str());
                SetDlgItemTextW(hDlg, IDC_EDIT_PARAMS, pInfo->parameters.c_str());
                CheckDlgButton(hDlg, IDC_CHECK_ADMIN, pInfo->adminMode ? BST_CHECKED : BST_UNCHECKED);
            }
            UpdateDialogControls(hDlg);
        }

        // Subclass edit controls to handle Ctrl+A for selecting all text
        SetWindowSubclass(GetDlgItem(hDlg, IDC_EDIT_NAME), SelectAllEditSubclassProcedure, 1, 0);
        SetWindowSubclass(GetDlgItem(hDlg, IDC_EDIT_PATH), SelectAllEditSubclassProcedure, 1, 0);
        SetWindowSubclass(GetDlgItem(hDlg, IDC_EDIT_PARAMS), SelectAllEditSubclassProcedure, 1, 0);
        SetWindowSubclass(GetDlgItem(hDlg, IDC_EDIT_COMMAND), SelectAllEditSubclassProcedure, 1, 0);
        return TRUE;
    }

    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
        case IDC_RADIO_LAUNCH:
        case IDC_RADIO_SENDKEYS:
        {
            UpdateDialogControls(hDlg);
            break;
        }
        case IDC_BUTTON_BROWSE:
        {
            // Open a file dialog to browse for an executable
            std::vector<wchar_t> fileName(MAX_PATH, L'\0');
            OPENFILENAMEW ofn = { sizeof(ofn) };
            ofn.hwndOwner = hDlg;
            ofn.lpstrFilter = L"Programs (*.exe;*.com;*.bat)\0*.exe;*.com;*.bat\0All Files (*.*)\0*.*\0";
            ofn.lpstrFile = fileName.data();
            ofn.nMaxFile = MAX_PATH;
            ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
            if (GetOpenFileNameW(&ofn))
            {
                SetDlgItemTextW(hDlg, IDC_EDIT_PATH, fileName.data());
            }
            break;
        }
        case IDC_BUTTON_URL:
        {
            ShellExecuteW(NULL, L"open", L"https://github.com/edgarp9/MultiTabLauncher", NULL, NULL, SW_SHOWNORMAL);
            break;
        }
        case IDOK:
        {
            ButtonInfo* pInfo = (ButtonInfo*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
            if (pInfo)
            {
                pInfo->name = GetTextFromDialogControl(hDlg, IDC_EDIT_NAME);
                trim(pInfo->name);

                if (IsDlgButtonChecked(hDlg, IDC_RADIO_SENDKEYS) == BST_CHECKED)
                {
                    pInfo->type = ButtonInfo::ActionType::SendKeys;
                    pInfo->parameters = GetTextFromDialogControl(hDlg, IDC_EDIT_COMMAND);
                    trim(pInfo->parameters);
                    pInfo->autoEnter = (IsDlgButtonChecked(hDlg, IDC_CHECK_AUTOENTER) == BST_CHECKED);

                    // Clear launch-related fields
                    pInfo->path.clear();
                    pInfo->adminMode = false;
                }
                else
                {
                    pInfo->type = ButtonInfo::ActionType::Launch;
                    pInfo->path = GetTextFromDialogControl(hDlg, IDC_EDIT_PATH);
                    trim(pInfo->path);
                    pInfo->parameters = GetTextFromDialogControl(hDlg, IDC_EDIT_PARAMS);
                    trim(pInfo->parameters);
                    pInfo->adminMode = (IsDlgButtonChecked(hDlg, IDC_CHECK_ADMIN) == BST_CHECKED);
                    pInfo->autoEnter = false;
                }
            }
            EndDialog(hDlg, IDOK);
            break;
        }
        case IDCANCEL:
        {
            EndDialog(hDlg, IDCANCEL);
            break;
        }
        }
        break;
    }
    }
    return FALSE;
}

/**
 * @brief Safely retrieves text from a dialog control, handling buffer allocation.
 * @param hDlg Handle to the dialog box.
 * @param nCtlId The identifier of the control.
 * @return The text from the control.
 */
std::wstring GetTextFromDialogControl(HWND hDlg, int nCtlId)
{
    HWND hCtrl = GetDlgItem(hDlg, nCtlId);
    if (!hCtrl) return L"";

    int len = GetWindowTextLengthW(hCtrl);
    if (len <= 0) return L"";

    std::wstring text(len, L'\0');
    GetWindowTextW(hCtrl, &text[0], len + 1);
    return text;
}

/**
 * @brief Subclass procedure for edit controls to handle Ctrl+A (select all).
 */
LRESULT CALLBACK SelectAllEditSubclassProcedure(HWND hEdit, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR)
{
    if (msg == WM_KEYDOWN && wParam == 'A' && (GetKeyState(VK_CONTROL) & 0x8000))
    {
        SendMessageW(hEdit, EM_SETSEL, 0, -1);
        return 0; // Message handled, don't pass to default proc
    }
    return DefSubclassProc(hEdit, msg, wParam, lParam);
}


// =============================================================
//                 Default INI Configuration
// =============================================================

/**
 * @brief Provides the default string content for the INI file.
 * @return A wstring containing the default configuration.
 */
std::wstring GetDefaultConfigString()
{
    return LR"([Tabs]
Count=10
ButtonRows=3
ButtonCols=8
Tab0=Home
Tab1=System
Tab2=Dev
Tab3=Debian
Tab4=Tab5
Tab5=Tab6
Tab6=Tab7
Tab7=Tab8
Tab8=Tab9
Tab9=Tab10

[Tab0]
Button0_Name=Explorer
Button0_Path=explorer.exe
Button0_Params=
Button0_Admin=0
Button0_Action=0
Button0_AutoEnter=0
Button1_Name=Notepad
Button1_Path=notepad.exe
Button1_Params=
Button1_Admin=0
Button1_Action=0
Button1_AutoEnter=0
Button2_Name=Snipping Tool
Button2_Path=snippingtool.exe
Button2_Params=
Button2_Admin=0
Button2_Action=0
Button2_AutoEnter=0
Button3_Name=Paint
Button3_Path=mspaint.exe
Button3_Params=
Button3_Admin=0
Button3_Action=0
Button3_AutoEnter=0
Button4_Name=Calculator
Button4_Path=calc.exe
Button4_Params=
Button4_Admin=0
Button4_Action=0
Button4_AutoEnter=0
Button5_Name=Chrome
Button5_Path=C:\Program Files\Google\Chrome\Application\chrome.exe
Button5_Params=
Button5_Admin=0
Button5_Action=0
Button5_AutoEnter=0
Button6_Name=
Button6_Path=
Button6_Params=
Button6_Admin=0
Button6_Action=0
Button6_AutoEnter=0
Button7_Name=
Button7_Path=
Button7_Params=
Button7_Admin=0
Button7_Action=0
Button7_AutoEnter=0
Button8_Name=Task Manager
Button8_Path=Taskmgr.exe
Button8_Params=
Button8_Admin=0
Button8_Action=0
Button8_AutoEnter=0
Button9_Name=Command Prompt
Button9_Path=cmd.exe
Button9_Params=
Button9_Admin=0
Button9_Action=0
Button9_AutoEnter=0
Button10_Name=Remote Desktop
Button10_Path=mstsc.exe
Button10_Params=
Button10_Admin=0
Button10_Action=0
Button10_AutoEnter=0
Button11_Name=Sandbox
Button11_Path=WindowsSandbox.exe
Button11_Params=
Button11_Admin=0
Button11_Action=0
Button11_AutoEnter=0
Button12_Name=
Button12_Path=
Button12_Params=
Button12_Admin=0
Button12_Action=0
Button12_AutoEnter=0
Button13_Name=Desktop
Button13_Path=explorer.exe
Button13_Params=%USERPROFILE%\Desktop
Button13_Admin=0
Button13_Action=0
Button13_AutoEnter=0
Button14_Name=Downloads
Button14_Path=explorer.exe
Button14_Params=%USERPROFILE%\Downloads
Button14_Admin=0
Button14_Action=0
Button14_AutoEnter=0
Button15_Name=
Button15_Path=
Button15_Params=
Button15_Admin=0
Button15_Action=0
Button15_AutoEnter=0

[Tab1]
Button0_Name=Device Manager
Button0_Path=devmgmt.msc
Button0_Params=
Button0_Admin=0
Button0_Action=0
Button0_AutoEnter=0
Button1_Name=Control Panel
Button1_Path=control.exe
Button1_Params=
Button1_Admin=0
Button1_Action=0
Button1_AutoEnter=0
Button2_Name=Disk Management
Button2_Path=diskmgmt.msc
Button2_Params=
Button2_Admin=0
Button2_Action=0
Button2_AutoEnter=0
Button3_Name=Computer Mgmt
Button3_Path=compmgmt.msc
Button3_Params=
Button3_Admin=0
Button3_Action=0
Button3_AutoEnter=0
Button4_Name=Programs
Button4_Path=appwiz.cpl
Button4_Params=
Button4_Admin=0
Button4_Action=0
Button4_AutoEnter=0
Button5_Name=System Info
Button5_Path=msinfo32.exe
Button5_Params=
Button5_Admin=0
Button5_Action=0
Button5_AutoEnter=0

[Tab2]
Button0_Name=VS Code
Button0_Path=code.exe
Button0_Params=
Button0_Admin=0
Button0_Action=0
Button0_AutoEnter=0
Button1_Name=VS 2022
Button1_Path=devenv.exe
Button1_Params=
Button1_Admin=0
Button1_Action=0
Button1_AutoEnter=0
Button2_Name=VS Installer
Button2_Path=C:\Program Files (x86)\Microsoft Visual Studio\Installer\setup.exe
Button2_Params=
Button2_Admin=0
Button2_Action=0
Button2_AutoEnter=0

[Tab3]
Button0_Name=update_upgrade
Button0_Path=
Button0_Params=sudo apt update && sudo apt upgrade
Button0_Admin=0
Button0_Action=1
Button0_AutoEnter=0
Button1_Name=poweroff
Button1_Path=
Button1_Params=sudo poweroff
Button1_Admin=0
Button1_Action=1
Button1_AutoEnter=0
Button2_Name=reboot
Button2_Path=
Button2_Params=sudo reboot
Button2_Admin=0
Button2_Action=1
Button2_AutoEnter=0
Button8_Name=Show IP
Button8_Path=
Button8_Params=ip addr show
Button8_Admin=0
Button8_Action=1
Button8_AutoEnter=0

)";
}