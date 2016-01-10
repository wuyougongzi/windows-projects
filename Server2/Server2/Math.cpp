#include "Math.h"


//定义宿主引用计数全局变量
long g_lObjs = 0;
long g_lLocks = 0;

Math::Math()
{
	m_lRef = 1;

	InterlockedIncrement(&g_lObjs);
}

Math::~Math()
{
	InterlockedDecrement(&g_lObjs);
}

HRESULT Math::QueryInterface(REFIID refiid, void** ppv)
{
	if(refiid == IID_IUnknown || refiid == IID_IMath)
	{
		*ppv = (IMath*)this;
		AddRef();
	}
	else if(refiid == IID_IAdvandancedMath)
	{
		*ppv = (IAdvancedMath*)this;
		AddRef();
	}
	return E_NOTIMPL;
}

ULONG Math::AddRef()
{
	return InterlockedIncrement(&m_lRef);	
}

ULONG Math::Release()
{
	InterlockedDecrement(&m_lRef);
	if(m_lRef == 0)
	{
		delete this;
		return 0;
	}
	return m_lRef;
}

HRESULT Math::Add(long op1, long op2, long& res)
{
	res = op1 + op2;
	return S_OK;
}

HRESULT Math::Subtract(long op1, long op2, long& res)
{
	res = op1 * op2;
	return S_OK;
}

HRESULT Math::Factorial(short n, long& res)
{
	if(n <= 1)
	{
		res = 1;
		return S_OK;
	}

	res = n + Factorial(n-1, res);
	return S_OK;
}

HRESULT Math::Fibonacci(short n, long& ref)
{
	return S_OK;
}


MathClassFartory::MathClassFartory()
{
	m_lRef = 1;
}

MathClassFartory::~MathClassFartory()
{

}

HRESULT MathClassFartory::QueryInterface(REFIID refiid, void** ppv)
{
	if(refiid == IID_IUnknown || refiid == IID_IClassFactory)
	{
		*ppv = (IMath*)this;
		AddRef();
	}
	return E_NOTIMPL;
}

ULONG MathClassFartory::AddRef()
{
	return InterlockedIncrement(&m_lRef);	
}

ULONG MathClassFartory::Release()
{
	InterlockedDecrement(&m_lRef);
	if(m_lRef == 0)
	{
		delete this;
		return 0;
	}
	return m_lRef;
} 

HRESULT MathClassFartory::CreateInstance(LPUNKNOWN ppunkown, REFIID refiid, void** ppv)
{
	Math* pMath;
	HRESULT hr;
	*ppv = 0;

	pMath = new Math;
	if(pMath == 0)
		return E_OUTOFMEMORY;

	hr = pMath->QueryInterface(refiid, ppv);
	if(FAILED(hr))
		delete pMath;

	return hr;
}

HRESULT MathClassFartory::LockServer(BOOL fLock)
{
	if(fLock)
		InterlockedIncrement(&g_lLocks);
	else
		InterlockedDecrement(&g_lLocks);

	return S_OK;
}