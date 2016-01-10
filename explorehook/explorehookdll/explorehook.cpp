#include "explorehook.h"
#include <stdio.h>

#define WM_HOOKSTART	WM_APP + 1

#pragma data_seg (".shared")

HINSTANCE g_HInstance = NULL;
static HHOOK			s_hCallHook	= NULL;
static WNDPROC g_pFunOldProc = NULL;

static HWND hTargetHwnd = NULL;

#pragma data_seg () 
#pragma comment(linker,"/SECTION:.shared,RWS") 

LRESULT CALLBACK CallHookProc(int nCode, WPARAM wParam, LPARAM lParam);


void LogInfo(char* info)
{
	if (info == NULL)
		return;

	FILE* pFile = NULL;
	errno_t err = fopen_s(&pFile, "hook.log", "a");
	if (err != 0 || pFile == NULL)
		return;

	SYSTEMTIME sysTime;
	::GetLocalTime(&sysTime);
	char pBuffer[256] = {0};
	sprintf_s(pBuffer, 256, "%4d-%02d-%02d %02d:%02d:%02d  %s\r\n", sysTime.wYear, sysTime.wMonth, sysTime.wDay,
		sysTime.wHour, sysTime.wMinute, sysTime.wSecond, info);

	fwrite(pBuffer, strlen(pBuffer), 1, pFile);
	fclose(pFile);
}


LRESULT CALLBACK HookProcessProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LogInfo("HookProcessProc success");
	switch(uMsg)
	{
	case WM_NCHITTEST:
			break;
	case WM_SETCURSOR:
	case WM_MOUSEMOVE:
		{
			LogInfo("WM_MOUSEMOVE success");

			// 			POINT pt;
			// 			pt.x = 
			//MessageBox(NULL, L"gouzi success", L"hook", 0);
		}
		break;
	default:
		break;
	}
	return S_OK;
}

LRESULT CALLBACK CallHookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	LogInfo("ExploreHookProcess in");
	CWPSTRUCT *pStruct = (CWPSTRUCT*)lParam;

	char lparamStr[20] = {0};
	sprintf_s(lparamStr, "message: %x" , pStruct->message);
	LogInfo(lparamStr);

// 	char hwndStr[20] = {0};
// 	sprintf_s(lparamStr, "tartgetwnd: %x" , pStruct->hwnd);
// 	LogInfo(lparamStr);
	
	if (pStruct != NULL && pStruct->message >= WM_HOOKSTART)
	{
		g_pFunOldProc = (WNDPROC)::SetWindowLongPtr(hTargetHwnd, GWLP_WNDPROC, (LONG_PTR)HookProcessProc);
		LogInfo("SetWindowLongPtr in");
	}
	
		// MSDN:第一个参数ignore；windows NT之后
	//	return CallNextHookEx(NULL, code, wParam, lParam); 
	return S_OK;
}

void /*__stdcall */SetExploreHook(HWND hwnd)
{
	hTargetHwnd = hwnd;
	//最后一个为0 表示所有系统钩子
//	HWND startWnd = FindWindowEx(hwnd, NULL, L"Start", NULL);
	DWORD threadID = -1;
	DWORD word = ::GetWindowThreadProcessId(hTargetHwnd, &threadID);

	char theadIDStr[80] = {0};
	sprintf_s(theadIDStr, "targetWndThreadID %d", threadID);
	LogInfo(theadIDStr);

	s_hCallHook = SetWindowsHookEx(WH_MOUSE, (HOOKPROC)CallHookProc, g_HInstance, word);
	if (s_hCallHook == NULL)
	{
		DWORD dwErrorCode = ::GetLastError();
		char buffer[64] = {0};
		sprintf_s(buffer, 64, "SetWindowsHookEx error, error code is %d", dwErrorCode);
		LogInfo(buffer);
		return ;
	}

// 	char hwndStr[80] = {0};
// 	sprintf_s(hwndStr, "TargetWndHwnd %x", hwnd);
// 	LogInfo(hwndStr);
	::SendMessage(hTargetHwnd, WM_HOOKSTART, 10, 10);
	//SendMessage(hwnd, WM_HOOKSHELLTRAYICON, 0, 0);
}


void /*__stdcall */UnSetExploreHook()
{
	UnhookWindowsHookEx(s_hCallHook);
}
