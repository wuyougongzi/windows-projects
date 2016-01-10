#include "explorehook.h"

HRESULT WINAPI DllMain(_In_ HANDLE _HDllHandle, _In_ DWORD _Reason, _In_opt_ LPVOID _Reserved)
{
	switch(_Reason)
	{
	case DLL_PROCESS_ATTACH:
		g_HInstance = (HINSTANCE)_HDllHandle;
		break;
	case DLL_PROCESS_DETACH:
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	}
	return TRUE;
}