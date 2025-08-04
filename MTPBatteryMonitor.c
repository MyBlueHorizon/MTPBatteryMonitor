#include <windows.h>
#include <PortableDeviceApi.h>
#include <PortableDevice.h>
#include <stdio.h>

#pragma comment(lib, "PortableDeviceGUIDs.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

// 全局变量
HWND hWnd;
HWND hStatusLabel;

// 函数声明
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
HRESULT GetBatteryLevel(DWORD* pBatteryLevel);

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PSTR szCmdLine, int iCmdShow)
{
    static TCHAR szAppName[] = TEXT("MTP Battery Monitor");
    WNDCLASS wndclass;
    
    wndclass.style = CS_HREDRAW | CS_VREDRAW;
    wndclass.lpfnWndProc = WndProc;
    wndclass.cbClsExtra = 0;
    wndclass.cbWndExtra = 0;
    wndclass.hInstance = hInstance;
    wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
    wndclass.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
    wndclass.lpszMenuName = NULL;
    wndclass.lpszClassName = szAppName;
    
    if (!RegisterClass(&wndclass))
    {
        MessageBox(NULL, TEXT("This program requires Windows NT!"), szAppName, MB_ICONERROR);
        return 0;
    }
    
    hWnd = CreateWindow(
        szAppName,                  // window class name
        TEXT("MTP设备电量显示器"),  // window caption
        WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX, // window style
        CW_USEDEFAULT,              // initial x position
        CW_USEDEFAULT,              // initial y position
        300,                       // initial x size
        150,                       // initial y size
        NULL,                      // parent window handle
        NULL,                      // window menu handle
        hInstance,                 // program instance handle
        NULL);                     // creation parameters
    
    // 创建标签控件
    hStatusLabel = CreateWindow(
        TEXT("STATIC"), 
        TEXT("正在获取电量..."), 
        WS_CHILD | WS_VISIBLE | SS_CENTER, 
        10, 20, 260, 50, 
        hWnd, 
        NULL, 
        hInstance, 
        NULL);
    
    // 设置字体
    HFONT hFont = CreateFont(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 
                             OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                             DEFAULT_PITCH | FF_SWISS, TEXT("Segoe UI"));
    SendMessage(hStatusLabel, WM_SETFONT, (WPARAM)hFont, TRUE);
    
    ShowWindow(hWnd, iCmdShow);
    UpdateWindow(hWnd);
    
    // 设置定时器，每10秒更新一次电量
    SetTimer(hWnd, 1, 10000, NULL);
    
    // 立即更新一次电量
    DWORD batteryLevel = 0;
    if (SUCCEEDED(GetBatteryLevel(&batteryLevel)))
    {
        TCHAR szText[100];
        wsprintf(szText, TEXT("当前MTP设备电量: %d%%"), batteryLevel);
        SetWindowText(hStatusLabel, szText);
    }
    else
    {
        SetWindowText(hStatusLabel, TEXT("无法获取MTP设备电量\n请确保设备已连接"));
    }
    
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
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
                TCHAR szText[100];
                wsprintf(szText, TEXT("当前MTP设备电量: %d%%"), batteryLevel);
                SetWindowText(hStatusLabel, szText);
            }
            else
            {
                SetWindowText(hStatusLabel, TEXT("无法获取MTP设备电量\n请确保设备已连接"));
            }
        }
        break;
        
    case WM_DESTROY:
        KillTimer(hWnd, 1);
        PostQuitMessage(0);
        break;
        
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
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
    
    CoInitialize(NULL);
    
    // 创建设备管理器
    hr = CoCreateInstance(
        &CLSID_PortableDeviceManager,
        NULL,
        CLSCTX_INPROC_SERVER,
        &IID_IPortableDeviceManager,
        (VOID**)&pDeviceManager);
    
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    DWORD cPnPDeviceIDs = 0;
    PWSTR* pPnpDeviceIDs = NULL;
    
    // 获取设备列表
    hr = IPortableDeviceManager_GetDevices(pDeviceManager, NULL, &cPnPDeviceIDs);
    if (FAILED(hr) || cPnPDeviceIDs == 0)
    {
        goto Cleanup;
    }
    
    pPnpDeviceIDs = (PWSTR*)CoTaskMemAlloc(sizeof(PWSTR) * cPnPDeviceIDs);
    hr = IPortableDeviceManager_GetDevices(pDeviceManager, pPnpDeviceIDs, &cPnPDeviceIDs);
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    // 打开第一个设备
    hr = CoCreateInstance(
        &CLSID_PortableDevice,
        NULL,
        CLSCTX_INPROC_SERVER,
        &IID_IPortableDevice,
        (VOID**)&pDevice);
    
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    hr = IPortableDevice_Open(pDevice, pPnpDeviceIDs[0], NULL);
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    // 获取设备内容
    hr = IPortableDevice_Content(pDevice, &pContent);
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    // 获取设备属性
    hr = IPortableDeviceContent_Properties(pContent, &pProperties);
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    // 创建要读取的属性集合
    hr = CoCreateInstance(
        &CLSID_PortableDeviceKeyCollection,
        NULL,
        CLSCTX_INPROC_SERVER,
        &IID_IPortableDeviceKeyCollection,
        (VOID**)&pPropertiesToRead);
    
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    // 添加电池电量属性
    hr = IPortableDeviceKeyCollection_Add(pPropertiesToRead, &WPD_DEVICE_POWER_LEVEL);
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    // 读取属性值
    hr = IPortableDeviceProperties_GetValues(pProperties, NULL, pPropertiesToRead, &pPropertyValues);
    if (FAILED(hr))
    {
        goto Cleanup;
    }
    
    // 获取电池电量值
    hr = IPortableDeviceValues_GetUnsignedIntegerValue(pPropertyValues, &WPD_DEVICE_POWER_LEVEL, pBatteryLevel);
    
Cleanup:
    // 释放资源
    if (pPropertyValues != NULL)
    {
        IPortableDeviceValues_Release(pPropertyValues);
    }
    
    if (pPropertiesToRead != NULL)
    {
        IPortableDeviceKeyCollection_Release(pPropertiesToRead);
    }
    
    if (pProperties != NULL)
    {
        IPortableDeviceProperties_Release(pProperties);
    }
    
    if (pContent != NULL)
    {
        IPortableDeviceContent_Release(pContent);
    }
    
    if (pDevice != NULL)
    {
        IPortableDevice_Release(pDevice);
    }
    
    if (pPnpDeviceIDs != NULL)
    {
        for (DWORD i = 0; i < cPnPDeviceIDs; i++)
        {
            CoTaskMemFree(pPnpDeviceIDs[i]);
        }
        CoTaskMemFree(pPnpDeviceIDs);
    }
    
    if (pDeviceManager != NULL)
    {
        IPortableDeviceManager_Release(pDeviceManager);
    }
    
    CoUninitialize();
    return hr;
}
