#include <windows.h>
#include <PortableDeviceApi.h>
#include <PortableDevice.h>
#include <stdio.h>
#include <locale.h>

#pragma comment(lib, "PortableDeviceGUIDs.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

// COM接口方法定义
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

HWND hWnd;
HWND hStatusLabel;

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
HRESULT GetBatteryLevel(DWORD* pBatteryLevel);

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow)
{
    // 设置程序为Unicode模式
    _wsetlocale(LC_ALL, L"chs");

    // 注册窗口类
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
    wc.lpszClassName = L"MTPBatteryMonitor";
    
    if (!RegisterClassW(&wc)) {
        MessageBoxW(NULL, L"窗口注册失败!", L"错误", MB_ICONERROR);
        return 0;
    }

    // 创建主窗口
    hWnd = CreateWindowW(
        L"MTPBatteryMonitor",
        L"MTP设备电量监控",
        WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT,
        320, 180,
        NULL, NULL, hInstance, NULL);

    // 创建状态标签
    hStatusLabel = CreateWindowW(
        L"STATIC", L"正在检测MTP设备...",
        WS_CHILD | WS_VISIBLE | SS_CENTER,
        10, 20, 280, 80,
        hWnd, NULL, hInstance, NULL);

    // 设置字体为微软雅黑
    HFONT hFont = CreateFontW(
        18, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
        L"微软雅黑");
    SendMessageW(hStatusLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    // 设置定时器，每5秒刷新一次
    SetTimer(hWnd, 1, 5000, NULL);

    // 立即刷新一次
    DWORD batteryLevel = 0;
    if (SUCCEEDED(GetBatteryLevel(&batteryLevel))) {
        WCHAR szText[100];
        swprintf_s(szText, 100, L"当前电量: %d%%", batteryLevel);
        SetWindowTextW(hStatusLabel, szText);
    } else {
        SetWindowTextW(hStatusLabel, L"未检测到MTP设备\n请连接设备后重试");
    }

    // 消息循环
    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return (int)msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
        case WM_TIMER: {
            DWORD batteryLevel = 0;
            if (SUCCEEDED(GetBatteryLevel(&batteryLevel))) {
                WCHAR szText[100];
                swprintf_s(szText, 100, L"当前电量: %d%%", batteryLevel);
                SetWindowTextW(hStatusLabel, szText);
            } else {
                SetWindowTextW(hStatusLabel, L"未检测到MTP设备\n请连接设备后重试");
            }
            break;
        }
        case WM_DESTROY:
            KillTimer(hWnd, 1);
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProcW(hWnd, msg, wParam, lParam);
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
