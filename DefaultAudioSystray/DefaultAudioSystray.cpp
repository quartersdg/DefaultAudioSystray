// DefaultAudioSystray.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "Shellapi.h"
#include "windowsx.h"
#include "Mmdeviceapi.h"
#include <mmeapi.h>
#include "Functiondiscoverykeys_devpkey.h"
#include "DefaultAudioSystray.h"


#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
HWND ghWnd;
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

#define SYSTRAY_MESSAGE (WM_USER + 100)

const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);

// ----------------------------------------------------------------------------
// PolicyConfig.h
// Undocumented COM-interface IPolicyConfig.
// Use for set default audio render endpoint
// @author EreTIk
// ----------------------------------------------------------------------------

interface DECLSPEC_UUID("f8679f50-850a-41cf-9c72-430f290290c8")
    IPolicyConfig;
class DECLSPEC_UUID("870af99c-171d-4f9e-af0d-e63df40c2bc9")
    CPolicyConfigClient;
// ----------------------------------------------------------------------------
// class CPolicyConfigClient
// {870af99c-171d-4f9e-af0d-e63df40c2bc9}
//
// interface IPolicyConfig
// {f8679f50-850a-41cf-9c72-430f290290c8}
//
// Query interface:
// CComPtr[IPolicyConfig] PolicyConfig;
// PolicyConfig.CoCreateInstance(__uuidof(CPolicyConfigClient));
//
// @compatible: Windows 7 and Later
// ----------------------------------------------------------------------------
interface IPolicyConfig : public IUnknown
{
public:

    virtual HRESULT GetMixFormat(
        PCWSTR,
        WAVEFORMATEX **
        );

    virtual HRESULT STDMETHODCALLTYPE GetDeviceFormat(
        PCWSTR,
        INT,
        WAVEFORMATEX **
        );

    virtual HRESULT STDMETHODCALLTYPE ResetDeviceFormat(
        PCWSTR
        );

    virtual HRESULT STDMETHODCALLTYPE SetDeviceFormat(
        PCWSTR,
        WAVEFORMATEX *,
        WAVEFORMATEX *
        );

    virtual HRESULT STDMETHODCALLTYPE GetProcessingPeriod(
        PCWSTR,
        INT,
        PINT64,
        PINT64
        );

    virtual HRESULT STDMETHODCALLTYPE SetProcessingPeriod(
        PCWSTR,
        PINT64
        );

    virtual HRESULT STDMETHODCALLTYPE GetShareMode(
        PCWSTR,
    struct DeviceShareMode *
        );

    virtual HRESULT STDMETHODCALLTYPE SetShareMode(
        PCWSTR,
    struct DeviceShareMode *
        );

    virtual HRESULT STDMETHODCALLTYPE GetPropertyValue(
        PCWSTR,
        const PROPERTYKEY &,
        PROPVARIANT *
        );

    virtual HRESULT STDMETHODCALLTYPE SetPropertyValue(
        PCWSTR,
        const PROPERTYKEY &,
        PROPVARIANT *
        );

    virtual HRESULT STDMETHODCALLTYPE SetDefaultEndpoint(
        __in PCWSTR wszDeviceId,
        __in ERole eRole
        );

    virtual HRESULT STDMETHODCALLTYPE SetEndpointVisibility(
        PCWSTR,
        INT
        );
};

interface DECLSPEC_UUID("568b9108-44bf-40b4-9006-86afe5b5a620")
    IPolicyConfigVista;
class DECLSPEC_UUID("294935CE-F637-4E7C-A41B-AB255460B862")
    CPolicyConfigVistaClient;
// ----------------------------------------------------------------------------
// class CPolicyConfigVistaClient
// {294935CE-F637-4E7C-A41B-AB255460B862}
//
// interface IPolicyConfigVista
// {568b9108-44bf-40b4-9006-86afe5b5a620}
//
// Query interface:
// CComPtr[IPolicyConfigVista] PolicyConfig;
// PolicyConfig.CoCreateInstance(__uuidof(CPolicyConfigVistaClient));
//
// @compatible: Windows Vista and Later
// ----------------------------------------------------------------------------
interface IPolicyConfigVista : public IUnknown
{
public:

    virtual HRESULT GetMixFormat(
        PCWSTR,
        WAVEFORMATEX **
        );  // not available on Windows 7, use method from IPolicyConfig

    virtual HRESULT STDMETHODCALLTYPE GetDeviceFormat(
        PCWSTR,
        INT,
        WAVEFORMATEX **
        );

    virtual HRESULT STDMETHODCALLTYPE SetDeviceFormat(
        PCWSTR,
        WAVEFORMATEX *,
        WAVEFORMATEX *
        );

    virtual HRESULT STDMETHODCALLTYPE GetProcessingPeriod(
        PCWSTR,
        INT,
        PINT64,
        PINT64
        );  // not available on Windows 7, use method from IPolicyConfig

    virtual HRESULT STDMETHODCALLTYPE SetProcessingPeriod(
        PCWSTR,
        PINT64
        );  // not available on Windows 7, use method from IPolicyConfig

    virtual HRESULT STDMETHODCALLTYPE GetShareMode(
        PCWSTR,
    struct DeviceShareMode *
        );  // not available on Windows 7, use method from IPolicyConfig

    virtual HRESULT STDMETHODCALLTYPE SetShareMode(
        PCWSTR,
    struct DeviceShareMode *
        );  // not available on Windows 7, use method from IPolicyConfig

    virtual HRESULT STDMETHODCALLTYPE GetPropertyValue(
        PCWSTR,
        const PROPERTYKEY &,
        PROPVARIANT *
        );

    virtual HRESULT STDMETHODCALLTYPE SetPropertyValue(
        PCWSTR,
        const PROPERTYKEY &,
        PROPVARIANT *
        );

    virtual HRESULT STDMETHODCALLTYPE SetDefaultEndpoint(
        __in PCWSTR wszDeviceId,
        __in ERole eRole
        );

    virtual HRESULT STDMETHODCALLTYPE SetEndpointVisibility(
        PCWSTR,
        INT
        );  // not available on Windows 7, use method from IPolicyConfig
};



HRESULT SetDefaultAudioPlaybackDevice(LPCWSTR devID)
{
    IPolicyConfigVista *pPolicyConfig;
    ERole reserved = eConsole;

    HRESULT hr = CoCreateInstance(__uuidof(CPolicyConfigVistaClient),
        NULL, CLSCTX_ALL, __uuidof(IPolicyConfigVista), (LPVOID *)&pPolicyConfig);
    if (SUCCEEDED(hr))
    {
        hr = pPolicyConfig->SetDefaultEndpoint(devID, reserved);
        pPolicyConfig->Release();
    }
    return hr;
}


LPWSTR GetMMDeviceFriendlyName(IMMDevice *pDevice) {
    HRESULT hr;
    LPWSTR Result = 0;
    IPropertyStore *pPropertyStore;
    hr = pDevice->OpenPropertyStore(STGM_READ, &pPropertyStore);
    if (SUCCEEDED(hr)) {
        PROPVARIANT varName;
        PropVariantInit(&varName);

        hr = pPropertyStore->GetValue(PKEY_Device_FriendlyName, &varName);

        int Length = wcslen(varName.pwszVal) + 1;
        Result = (LPWSTR)calloc(sizeof(WCHAR), Length);
        wcscpy_s(Result, Length, varName.pwszVal);

        PropVariantClear(&varName);

        pPropertyStore->Release();
    }
    return Result;
}

void ShowContextMenu(HWND hWnd, int x, int y) {
    HMENU hMenu;

    hMenu = CreatePopupMenu();
    RECT rect;
    UINT_PTR id;

    HRESULT hr;

    IMMDeviceEnumerator *pEnumerator;

    hr = CoCreateInstance(
        CLSID_MMDeviceEnumerator, NULL,
        CLSCTX_ALL, IID_IMMDeviceEnumerator,
        (void**)&pEnumerator);

    if (SUCCEEDED(hr)) {
        IMMDeviceCollection *pDevices;
        pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pDevices);

        IMMDevice *pDefaultEndpoint;
        LPWSTR pDefaultId = 0;

        hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eMultimedia, &pDefaultEndpoint);
        if (SUCCEEDED(hr)) {
            hr = pDefaultEndpoint->GetId(&pDefaultId);

            pDefaultEndpoint->Release();
        }

        UINT DeviceCount;
        hr = pDevices->GetCount(&DeviceCount);
        if (SUCCEEDED(hr)) {
            for (int DeviceIndex = 0;DeviceIndex < DeviceCount; DeviceIndex++) {
                IMMDevice *pDevice;
                hr = pDevices->Item(DeviceIndex, &pDevice);
                if (SUCCEEDED(hr)) {
                    LPWSTR pId;
                    hr = pDevice->GetId(&pId);
                    if (SUCCEEDED(hr)) {
                        LPWSTR FriendlyName = GetMMDeviceFriendlyName(pDevice);
                        if (FriendlyName) {
                            UINT Flags = MF_STRING;
                            if (0 == wcscmp(pDefaultId, pId)) {
                                Flags |= MF_CHECKED;
                            }
                            AppendMenu(hMenu, Flags, IDM_1 + DeviceIndex + 1, FriendlyName);

                            free(FriendlyName);
                        }

                        CoTaskMemFree(pId);
                    }
                    pDevice->Release();
                }
            }
        }
        pDevices->Release();

        pEnumerator->Release();
    }

    AppendMenu(hMenu, MF_STRING, IDM_1, _TEXT("Exit"));

    SetForegroundWindow(hWnd);
    TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_BOTTOMALIGN, x, y, 0, hWnd, &rect);
}

void TryToSwitchDefaultAudioDevice(UINT wmId) {
    HRESULT hr;

    IMMDeviceEnumerator *pEnumerator;

    hr = CoCreateInstance(
        CLSID_MMDeviceEnumerator, NULL,
        CLSCTX_ALL, IID_IMMDeviceEnumerator,
        (void**)&pEnumerator);

    if (SUCCEEDED(hr)) {
        IMMDeviceCollection *pDevices;
        pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pDevices);

        UINT DeviceCount;
        hr = pDevices->GetCount(&DeviceCount);
        if (SUCCEEDED(hr)) {
            int DeviceIndex = wmId - IDM_1 - 1;
            if (DeviceIndex < DeviceCount) {
                IMMDevice *pDevice;
                hr = pDevices->Item(DeviceIndex, &pDevice);
                if (SUCCEEDED(hr)) {
                    LPWSTR pId;
                    hr = pDevice->GetId(&pId);
                    if (SUCCEEDED(hr)) {
                        SetDefaultAudioPlaybackDevice(pId);

                        CoTaskMemFree(pId);
                    }
                    pDevice->Release();
                }
            }
        }
        pDevices->Release();

        pEnumerator->Release();
    }

}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_DEFAULTAUDIOSYSTRAY, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_DEFAULTAUDIOSYSTRAY));

    MSG msg;

    NOTIFYICONDATA nid = {};
    nid.cbSize = sizeof(nid);
    nid.hWnd = ghWnd;
    nid.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;
    wcscpy_s(nid.szTip,128,_TEXT("Switch Default Audio"));
    nid.hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_TRAYICON));
    nid.uCallbackMessage = SYSTRAY_MESSAGE;
    nid.uVersion = NOTIFYICON_VERSION_4;
    nid.uID = 1;

    Shell_NotifyIcon(NIM_ADD, &nid);
    Shell_NotifyIcon(NIM_SETVERSION, &nid);

    CoInitialize(NULL);

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    Shell_NotifyIcon(NIM_DELETE, &nid);

    return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_DEFAULTAUDIOSYSTRAY));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_DEFAULTAUDIOSYSTRAY);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   //ShowWindow(hWnd, nCmdShow);
   //UpdateWindow(hWnd);

   ghWnd = hWnd;

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case SYSTRAY_MESSAGE:
    {
        switch (LOWORD(lParam)) {
        case WM_RBUTTONDOWN:
        case WM_CONTEXTMENU:
            ShowContextMenu(hWnd, GET_X_LPARAM(wParam), GET_Y_LPARAM(wParam));
        } break;
    } break;
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_1:
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                if (wmId > IDM_1) {
                    TryToSwitchDefaultAudioDevice(wmId);
                }
                else {
                    return DefWindowProc(hWnd, message, wParam, lParam);
                }
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
