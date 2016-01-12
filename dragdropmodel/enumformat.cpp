
#include "stdafx.h"
#include "enumformat.h"

HRESULT CreateEnumFormatEtc(UINT nNumFormats, FORMATETC* pFormatEtc, IEnumFORMATETC** ppEnumFormatEtc)
{
	if(nNumFormats == 0 || pFormatEtc == NULL || ppEnumFormatEtc == NULL)
		return E_INVALIDARG;

	*ppEnumFormatEtc = new CEnumFormatEtc(pFormatEtc, nNumFormats);
	return (*ppEnumFormatEtc) ? S_OK : E_OUTOFMEMORY;
}

void DeepCopyFormatEtc(FORMATETC* dest, FORMATETC* source)
{
	*dest = *source;
	if(source->ptd)
	{
		dest->ptd = (DVTARGETDEVICE*)CoTaskMemAlloc(sizeof(DVTARGETDEVICE));
		*(dest->ptd) = *(source->ptd);
	}
}

CEnumFormatEtc::CEnumFormatEtc(FORMATETC* pFormatEtc, int nNumFormats)
{
	m_lRefCount = 1;
	m_nIndex = 0;
	m_nNumFormats = nNumFormats;
	m_pFormatEtc = new FORMATETC[nNumFormats];

	for(int i = 0; i < nNumFormats; ++i)
	{
		DeepCopyFormatEtc(&m_pFormatEtc[i], &pFormatEtc[i]);
	}
}

CEnumFormatEtc::~CEnumFormatEtc()
{
	if(m_pFormatEtc)
	{
		for(ULONG i = 0; i < m_nNumFormats; ++i)
		{
			if(m_pFormatEtc[i].ptd)
				CoTaskMemFree(m_pFormatEtc[i].ptd);
		}
		delete[] m_pFormatEtc;
	}
}

ULONG CEnumFormatEtc::AddRef()
{
	return InterlockedIncrement(&m_lRefCount);
}

ULONG CEnumFormatEtc::Release()
{
	LONG count = InterlockedDecrement(&m_lRefCount);

	if(count = 0)
	{
		delete this;
		return 0;
	}
	else
	{
		return count;
	}
}

HRESULT CEnumFormatEtc::QueryInterface(REFIID iid, void** ppvObject)
{
	if(iid == IID_IEnumFORMATETC || iid == IID_IUnknown)
	{
		AddRef();
		*ppvObject = this;
		return S_OK;
	}
	else
	{
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}
}

HRESULT CEnumFormatEtc::Next(ULONG celt, FORMATETC* pFormatEtc, ULONG* pceltFetched)
{
	ULONG copied = 0;
	
	if(celt == 0 || pFormatEtc == NULL)
		return E_INVALIDARG;

	while(m_nIndex < m_nNumFormats && copied < celt)
	{
		DeepCopyFormatEtc(&pFormatEtc[copied], &m_pFormatEtc[m_nIndex]);
		copied++;
		m_nIndex++;
	}

	if(pceltFetched != NULL)
		*pceltFetched  = copied;

	return (copied == celt) ? S_OK : S_FALSE;
}

HRESULT CEnumFormatEtc::Skip(ULONG celt)
{
	m_nIndex += celt;
	return (m_nIndex <= m_nNumFormats) ? S_OK : S_FALSE;
}

HRESULT CEnumFormatEtc::Reset()
{
	m_nIndex = 0;
	return S_OK;
}

HRESULT CEnumFormatEtc::Clone(IEnumFORMATETC** ppEnumFormatEtc)
{
	HRESULT hResult;
	hResult = CreateEnumFormatEtc(m_nNumFormats, m_pFormatEtc, ppEnumFormatEtc);
	if(hResult == S_OK)
		((CEnumFormatEtc *)*ppEnumFormatEtc)->m_nIndex = m_nIndex;

	return hResult;
}