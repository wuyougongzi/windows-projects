#include <Windows.h>
//#include "../explorehookdll/explorehook.h"
//SetExploreHook
//UnSetExploreHook
//#pragma comment(lib, "explorehookdll.lib")

//comment(lib, "KeyboardHookDLL.lib")
typedef void (*PFUN)(HWND);

 //__declspec(dllexport) void SetExploreHook(HWND hBackWnd);
 //__declspec(dllexport) void UnSetExploreHook();

//__declspec(dllimport) void __stdcall SetExploreHook(HWND);


WCHAR g_szWindowClass[] = L"Richer";
WCHAR g_szWindowTitle[] = L"explorelisten";

HRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);

int WINAPI WinMain( __in HINSTANCE hInstance, __in_opt HINSTANCE hPrevInstance, __in LPSTR lpCmdLine, __in int nShowCmd )
{
	HWND hShellIconwnd = FindWindowW(L"Shell_TrayWnd", NULL);

	HWND statWnd = FindWindowEx(hShellIconwnd, NULL, L"Start", NULL);
	//设置隐藏的窗口
	WNDCLASSEX wcex;

	wcex.cbSize			= sizeof(WNDCLASSEX);
	wcex.style          = CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc    = WndProc;
	wcex.cbClsExtra     = 0;
	wcex.cbWndExtra     = 0;
	wcex.hInstance      = hInstance;
	wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPLICATION));
	wcex.hCursor        = LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW);
	wcex.lpszMenuName   = NULL;
	wcex.lpszClassName  = g_szWindowClass;
	wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_APPLICATION));

	if (!RegisterClassEx(&wcex))
		return FALSE;

	HWND hWnd = CreateWindow(
		g_szWindowClass,
		g_szWindowTitle,
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		1,
		1,
		NULL,
		NULL,
		hInstance,
		NULL
		);

	if (!hWnd)
		return FALSE;

	ShowWindow(hWnd, SW_HIDE);
	UpdateWindow(hWnd);
	
	//todo:设置钩子
	//SetDIPSHook(hwnd);
	//SetWindowsHook()
	//SetWindowsHookEx(WH_GETMESSAGE, )

	HMODULE hModule = ::LoadLibrary(L"explorehookdll.dll");
	PFUN SetExploreHook= (PFUN)::GetProcAddress(hModule,"SetExploreHook");

	SetExploreHook(statWnd);

	MSG msg;
	while(GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return (int)msg.wParam;
}

HRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	return DefWindowProc(hWnd, message, wParam, lParam);
}