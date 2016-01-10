#include <Windows.h>
#include <InitGuid.h>
#include <iostream>

#include "..\Server2\IMath.h"

using namespace std;

int main()
{
	cout << "init com \n";

	if(FAILED(CoInitialize(NULL)))
	{
		cout << "init failed\n";
	}

	CLSID clsid;
	HRESULT hr = ::CLSIDFromProgID(L"Chapter2.Math.1", &clsid);
	if(FAILED(hr))
	{
		cout << "get CLISI from ProgId failed \n";
	}

	IClassFactory* pcf = NULL;

	hr = CoGetClassObject(
		clsid, 
		CLSCTX_INPROC,
		NULL,
		IID_IClassFactory,
		(void**)&pcf);

	if(FAILED(hr))
	{
		cout << "get IClass Factory failed \n";
	}

	IUnknown* pUnk = NULL;
	hr = pcf->CreateInstance(NULL, IID_IUnknown, (void**)&pUnk);
	
	pcf->Release();

	if(FAILED(hr))
	{
		cout << "get Iunkdown failed \n";
	}


	IMath* pMath = NULL;
	hr = pUnk->QueryInterface(IID_IMath, (void**)&pMath);
	pUnk->Release();
	
	if(FAILED(hr))
	{
		cout <<"get Imath failed \n";
	}

	long op = 1;
	long op2 = 3;
	long res = 0;
	pMath->Add(op, op2, res);

	cout << "Imath add : res = " << res << " \n";

	pMath->Subtract(op, op2, res);

	cout << "Imath subtract : res = " << res << " \n";


	system("pause");
}

