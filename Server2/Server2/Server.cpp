#include <InitGuid.h>
#include <ComCat.h>
#include "ATLDevGuid.h"

#include "Math.h"

//全局实例对象
//需要用这个调用CallModuleFileName API
HINSTANCE g_hinstDLL = 0;


BOOL WINAPI DllMain(HINSTANCE hinstDll, DWORD fdwReason, LPVOID lpvReserved)
{
	if(fdwReason == DLL_PROCESS_ATTACH)
		g_hinstDLL = hinstDll;

	return TRUE;
}

STDAPI DllGetClassObject(__in REFCLSID rclsid, __in REFIID riid, __deref_out LPVOID FAR* ppv)
{
	HRESULT hr;
	MathClassFartory* pcf = NULL;

	if(rclsid != CLSID_Math)
		return E_FAIL;

	pcf = new MathClassFartory;
	if(pcf == 0)
		return E_OUTOFMEMORY;

	hr = pcf->QueryInterface(riid, ppv);
	if(FAILED(hr))
	{
		delete pcf;
		pcf = NULL;
	}

	return hr;
}

// Helper function to create a component category and associated
// description
HRESULT CreateComponentCategory(CATID catid, WCHAR* catDescription)
{
	ICatRegister* pcr = NULL ;
	HRESULT hr = S_OK ;

	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_ICatRegister,
		(void**)&pcr);
	if (FAILED(hr))
		return hr;

	// Make sure the HKCR\Component Categories\{..catid...}
	// key is registered
	CATEGORYINFO catinfo;
	catinfo.catid = catid;
	catinfo.lcid = 0x0409 ; // english

	// Make sure the provided description is not too long.
	// Only copy the first 127 characters if it is
	int len = wcslen(catDescription);
	if (len>127)
		len = 127;
	wcsncpy(catinfo.szDescription, catDescription, len);
	// Make sure the description is null terminated
	catinfo.szDescription[len] = '\0';

	hr = pcr->RegisterCategories(1, &catinfo);
	pcr->Release();

	return hr;
}


STDAPI DllCanUnloadNow()
{
	if(g_lObjs || g_lLocks )
		return E_FAIL;

	return S_OK;

}

STDAPI DllUnregisterServer(void)
{
	// For now, don't do anything...
	return S_OK;
}

static BOOL SetRegKeyValue(
	LPTSTR pszKey,
	LPTSTR pszSubKey,
	LPTSTR pszValue)
{
	BOOL bOk = FALSE;
	LONG ec;
	HKEY hKey;
	TCHAR szKey[128];

	lstrcpy(szKey, pszKey);

	if(pszSubKey != NULL)
	{
		lstrcat(szKey, "\\");
		lstrcat(szKey, pszSubKey);
	}
	ec = RegCreateKeyEx(
		HKEY_CLASSES_ROOT,
		szKey,
		0,
		NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_ALL_ACCESS,
		NULL,
		&hKey,
		NULL
		);

	if(ec == ERROR_SUCCESS)
	{
		if(pszValue != NULL)
		{
			ec =RegSetValueEx(
				hKey,
				NULL,
				0,
				REG_SZ,
				(BYTE *)pszValue,
				(lstrlen(pszValue) + 1) *sizeof(TCHAR));
		}

		if(ec == ERROR_SUCCESS)
			bOk = TRUE;

		RegCloseKey(hKey);
	}

	return bOk;
}


// Helper function to register a CLSID as belonging to a component
// category
HRESULT RegisterCLSIDInCategory(REFCLSID clsid, CATID catid)
{
	// Register your component categories information.
	ICatRegister* pcr = NULL ;
	HRESULT hr = S_OK ;
	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_ICatRegister,
		(void**)&pcr);

	if (SUCCEEDED(hr))
	{
		// Register this category as being "implemented" by
		// the class.
		CATID rgcatid[1] ;
		rgcatid[0] = catid;
		hr = pcr->RegisterClassImplCategories(clsid, 1, rgcatid);
	}

	if (pcr != NULL)
		pcr->Release();

	return hr;
}

STDAPI DllRegisterServer()
{
	HRESULT hr = NOERROR;
	CHAR szModulePath[MAX_PATH];
	CHAR szID[128];
	CHAR szCLSID[128];
	
	WCHAR wszID[128];
	WCHAR wszCLSID[128];

	GetModuleFileName(
		g_hinstDLL,
		szModulePath,
		sizeof(szModulePath) / sizeof(CHAR));

	StringFromGUID2(CLSID_Math, wszID, sizeof(wszID));
	wcscpy(wszCLSID, L"CLSID\\");
	wcscat(wszCLSID, wszID);
	wcstombs(szID, wszID, sizeof(szID));
	wcstombs(szCLSID, wszCLSID, sizeof(szID));


	SetRegKeyValue(
		"Chapter2.Math.1",
		NULL,
		"Chapter2 Math Component" );

	SetRegKeyValue(
		"Chapter2.Math.1",
		"CLSID",
		szID);

	// Create version independent ProgID keys.
	SetRegKeyValue(
		"Chapter2.Math",
		NULL,
		"Chapter2 Math Component");
	SetRegKeyValue(
		"Chapter2.Math",
		"CurVer",
		"Chapter2.Math.1");
	SetRegKeyValue(
		"Chapter2.Math",
		"CLSID",
		szID);

	// Create entries under CLSID.
	SetRegKeyValue(
		szCLSID,
		NULL,
		"Chapter 2 Math Component");
	SetRegKeyValue(
		szCLSID,
		"ProgID",
		"Chapter2.Math.1");
	SetRegKeyValue(
		szCLSID,
		"VersionIndependentProgID",
		"Chapter2.Math");
	SetRegKeyValue(
		szCLSID,
		"InprocServer32",
		szModulePath);

	CreateComponentCategory(
		CATID_ATLDevGuide,
		L"ATL Developer's Guide Examples");
	
	RegisterCLSIDInCategory(CLSID_Math, CATID_ATLDevGuide);
	
	return S_OK;
	
}

