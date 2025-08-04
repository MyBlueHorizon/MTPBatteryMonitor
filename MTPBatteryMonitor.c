#include <windows.h>
#include <PortableDeviceApi.h>
#include <PortableDevice.h>
#include <stdio.h>
#include <locale.h>  // 添加这个头文件以支持setlocale和LC_ALL

#pragma comment(lib, "PortableDeviceGUIDs.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

// 手动定义COM接口方法
#define IPortableDeviceManager_GetDevices(This,pPnPDeviceIDs,pcPnPDeviceIDs) \
    (This)->lpVtbl->GetDevices(This,pPnPDeviceIDs,pcPnPDeviceIDs)

#define IPortableDevice_Open(This,pszPnPDeviceID,pClientInfo) \
    (This)->lpVtbl->Open(This,pszPnPDeviceID,pClientInfo)

#define IPortableDevice_Content(This,ppContent) \
    (This)->lpVtbl->Content(This,ppContent)

#define IPortableDeviceContent_Properties(This,ppProperties) \
    (This)->lpVtbl->Properties(This,ppProperties)

#define IPortableDeviceKeyCollection_Add(This,key) \
    (This)->lpVtbl->Add(This,key)

#define IPortableDeviceProperties_GetValues(This,pszObjectID,pKeys,ppValues) \
    (This)->lpVtbl->GetValues(This,pszObjectID,pKeys,ppValues)

#define IPortableDeviceValues_GetUnsignedIntegerValue(This,key,pValue) \
    (This)->lpVtbl->GetUnsignedIntegerValue(This,key,pValue)

#define IPortableDeviceValues_Release(This) \
    (This)->lpVtbl->Release(This)

#define IPortableDeviceKeyCollection_Release(This) \
    (This)->lpVtbl->Release(This)

#define IPortableDeviceProperties_Release(This) \
    (This)->lpVtbl->Release(This)

#define IPortableDeviceContent_Release(This) \
    (This)->lpVtbl->Release(This)

#define IPortableDevice_Release(This) \
    (This)->lpVtbl->Release(This)

#define IPortableDeviceManager_Release(This) \
    (This)->lpVtbl->Release(This)

// 全局变量
HWND hWnd;
HWND hStatusLabel;

// 函数声明
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
HRESULT GetBatteryLevel(DWORD* pBatteryLevel);

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PSTR szCmdLine, int iCmdShow)
{
    // 设置本地化为中文
    setlocale(LC_ALL, "chs");

    WCHAR szAppName[] = L"MTP Battery Monitor";
    WNDCLASSW wndclass = {0};
    
    wndclass.style = CS_HREDRAW | CS_VREDRAW;
    wndclass.lpfnWndProc = WndProc;
    wndclass.hInstance = hInstance;
    wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
    wndclass.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
    wndclass.lpszClassName = szAppName;
    
    if (!RegisterClassW(&wndclass))
    {
        MessageBoxW(NULL, L"程序需要Windows NT支持!", szAppName, MB_ICONERROR);
        return 0;
    }
    
    hWnd = CreateWindowW(
        szAppName,
        L"MTP设备电量显示器",
        WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        300,
        150,
        NULL,
        NULL,
        hInstance,
        NULL);
    
    hStatusLabel = CreateWindowW(
        L"STATIC", 
        L"正在获取电量...", 
        WS_CHILD | WS_VISIBLE | SS_CENTER, 
        10, 20, 260, 50, 
        hWnd, 
        NULL, 
        hInstance, 
        NULL);
    
    HFONT hFont = CreateFontW(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 
                           OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                           DEFAULT_PITCH | FF_SWISS, L"微软雅黑");
    SendMessageW(hStatusLabel, WM_SETFONT, (WPARAM)hFont, TRUE);
    
    ShowWindow(hWnd, iCmdShow);
    UpdateWindow(hWnd);
    
    SetTimer(hWnd, 1, 10000, NULL);
    
    DWORD batteryLevel = 0;
    if (SUCCEEDED(GetBatteryLevel(&batteryLevel)))
    {
        WCHAR szText[100];
        swprintf(szText, 100, L"当前MTP设备电量: %d%%", batteryLevel);
        SetWindowTextW(hStatusLabel, szText);
    }
    else
    {
        SetWindowTextW(hStatusLabel, L"无法获取MTP设备电量\n请确保设备已连接");
    }
    
    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    
    return (int)msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_TIMER:
        {
            DWORD batteryLevel = 0;
            if (SUCCEEDED(GetBatteryLevel(&batteryLevel)))
            {
                WCHAR szText[100];
                swprintf(szText, 100, L"当前MTP设备电量: %d%%", batteryLevel);
                SetWindowTextW(hStatusLabel, szText);
            }
            else
            {
                SetWindowTextW(hStatusLabel, L"无法获取MTP设备电量\n请确保设备已连接");
            }
        }
        break;
        
    case WM_DESTROY:
        KillTimer(hWnd, 1);
        PostQuitMessage(0);
        break;
        
    default:
        return DefWindowProcW(hWnd, message, wParam, lParam);
    }
    return 0;
}

HRESULT GetBatteryLevel(DWORD* pBatteryLevel)
{
    HRESULT hr = S_OK;
    IPortableDeviceManager* pDeviceManager = NULL;
    IPortableDevice* pDevice = NULL;
    IPortableDeviceContent* pContent = NULL;
    IPortableDeviceProperties* pProperties = NULL;
    IPortableDeviceKeyCollection* pPropertiesToRead = NULL;
    IPortableDeviceValues* pPropertyValues = NULL;
    PWSTR* pPnpDeviceIDs = NULL;
    DWORD cPnPDeviceIDs = 0;
    
    hr = CoInitialize(NULL);
    if (FAILED(hr)) return hr;
    
    hr = CoCreateInstance(
        &CLSID_PortableDeviceManager,
        NULL,
        CLSCTX_INPROC_SERVER,
        &IID_IPortableDeviceManager,
        (VOID**)&pDeviceManager);
    
    if (FAILED(hr)) goto Cleanup;
    
    hr = IPortableDeviceManager_GetDevices(pDeviceManager, NULL, &cPnPDeviceIDs);
    if (FAILED(hr) || cPnPDeviceIDs == 0)
    {
        hr = E_FAIL;
        goto Cleanup;
    }
    
    pPnpDeviceIDs = (PWSTR*)CoTaskMemAlloc(sizeof(PWSTR) * cPnPDeviceIDs);
    if (!pPnpDeviceIDs)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    
    hr = IPortableDeviceManager_GetDevices(pDeviceManager, pPnpDeviceIDs, &cPnPDeviceIDs);
    if (FAILED(hr)) goto Cleanup;
    
    hr = CoCreateInstance(
        &CLSID_PortableDevice,
        NULL,
        CLSCTX_INPROC_SERVER,
        &IID_IPortableDevice,
        (VOID**)&pDevice);
    
    if (FAILED(hr)) goto Cleanup;
    
    hr = IPortableDevice_Open(pDevice, pPnpDeviceIDs[0], NULL);
    if (FAILED(hr)) goto Cleanup;
    
    hr = IPortableDevice_Content(pDevice, &pContent);
    if (FAILED(hr)) goto Cleanup;
    
    hr = IPortableDeviceContent_Properties(pContent, &pProperties);
    if (FAILED(hr)) goto Cleanup;
    
    hr = CoCreateInstance(
        &CLSID_PortableDeviceKeyCollection,
        NULL,
        CLSCTX_INPROC_SERVER,
        &IID_IPortableDeviceKeyCollection,
        (VOID**)&pPropertiesToRead);
    
    if (FAILED(hr)) goto Cleanup;
    
    hr = IPortableDeviceKeyCollection_Add(pPropertiesToRead, &WPD_DEVICE_POWER_LEVEL);
    if (FAILED(hr)) goto Cleanup;
    
    hr = IPortableDeviceProperties_GetValues(pProperties, NULL, pPropertiesToRead, &pPropertyValues);
    if (FAILED(hr)) goto Cleanup;
    
    hr = IPortableDeviceValues_GetUnsignedIntegerValue(pPropertyValues, &WPD_DEVICE_POWER_LEVEL, pBatteryLevel);
    
Cleanup:
    if (pPropertyValues) IPortableDeviceValues_Release(pPropertyValues);
    if (pPropertiesToRead) IPortableDeviceKeyCollection_Release(pPropertiesToRead);
    if (pProperties) IPortableDeviceProperties_Release(pProperties);
    if (pContent) IPortableDeviceContent_Release(pContent);
    if (pDevice) IPortableDevice_Release(pDevice);
    
    if (pPnpDeviceIDs)
    {
        for (DWORD i = 0; i < cPnPDeviceIDs; i++)
        {
            CoTaskMemFree(pPnpDeviceIDs[i]);
        }
        CoTaskMemFree(pPnpDeviceIDs);
    }
    
    if (pDeviceManager) IPortableDeviceManager_Release(pDeviceManager);
    
    CoUninitialize();
    return hr;
}
