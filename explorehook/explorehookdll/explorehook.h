
#ifndef EXPLOREHOOK_H__
#define EXPLOREHOOK_H__
#include <windows.h>

extern HINSTANCE g_HInstance;

/* __declspec(dllexport) */void /*__stdcall */SetExploreHook(HWND);
/*__declspec(dllexport) */void /*__stdcall */UnSetExploreHook();

#endif	//EXPLOREHOOK_H__