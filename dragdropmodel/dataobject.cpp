#include "stdafx.h"
#include "dataobject.h"

HRESULT CreateEnumFormatEtc(UINT nNumFormats, FORMATETC* pFormatEtc, IEnumFORMATETC** ppEnumFormatEtc);


CDataObject::CDataObject(FORMATETC* fmtetc, STGMEDIUM* stgmed, int count) 
{
	m_lRefCount  = 1;

	m_pFormatEtc.resize(count);
	m_pStgMedium.resize(count);
	for(int i = 0; i < count; i++)
	{
		m_pFormatEtc[i] = fmtetc[i];
		m_pStgMedium[i] = stgmed[i];
	}
}

CDataObject::~CDataObject()
{

}

ULONG CDataObject::AddRef()
{
	return InterlockedIncrement(&m_lRefCount);
}

ULONG CDataObject::Release()
{
	LONG count = InterlockedDecrement(&m_lRefCount);
	if(count == 0)
	{
		delete this;
		return 0;
	}
	else
	{
		return count;
	}
}

HRESULT CDataObject::QueryInterface(REFIID iid, void** ppvObject)
{
	if(iid == IID_IDataObject || iid == IID_IUnknown)
	{
		AddRef();
		*ppvObject = this;
		return S_OK;
	}
	else
	{
		*ppvObject = 0;
		return E_NOINTERFACE;
	}
}

HGLOBAL DupMem(HGLOBAL hMem)
{
	DWORD len = GlobalSize(hMem);
	PVOID source = GlobalLock(hMem);
	
	PVOID dest = GlobalAlloc(GMEM_FIXED, len);
	memcpy(dest, source, len);
	GlobalUnlock(hMem);
	return dest;
}

int CDataObject::LookupFormatEtc(FORMATETC* pFormatEtc)
{
	for(size_t i = 0; i < m_pFormatEtc.size(); i++)
	{
		if((pFormatEtc->tymed & m_pFormatEtc[i].tymed)
		&& pFormatEtc->cfFormat == m_pFormatEtc[i].cfFormat
		&& pFormatEtc->dwAspect == m_pFormatEtc[i].dwAspect)
		{
			return i;
		}
	}
	return -1;
}

HRESULT CDataObject::GetData(FORMATETC* pFormatEtc, STGMEDIUM* pMedium)
{
	/*
	Called by a data consumer to obtain data from a source data object
	*/

	if(pFormatEtc == NULL || pMedium == NULL)
		return E_INVALIDARG;

	int idx;
	if((idx = LookupFormatEtc(pFormatEtc)) == -1)
		return DV_E_FORMATETC;
	
	pMedium->tymed = m_pFormatEtc[idx].tymed;
	pMedium->pUnkForRelease = NULL;
	
	switch(m_pFormatEtc[idx].tymed)
	{
	case TYMED_HGLOBAL:
		{
			pMedium->hGlobal = DupMem(m_pStgMedium[idx].hGlobal);
			break;
		}
		
	case TYMED_ISTREAM:
		pMedium->pstm = m_pStgMedium[idx].pstm;
		pMedium->pstm->AddRef();
		break;
	default:
		return DV_E_FORMATETC;
	}
	return S_OK;
}

HRESULT CDataObject::GetDataHere(FORMATETC* pFormatEtc, STGMEDIUM* pMedium)
{
	/*
	Called by a data consumer to obtain data from a source data object. 
	This method differs from the GetData method in that the caller must 
	allocate and free the specified storage medium
	*/
	return DATA_E_FORMATETC;
}

HRESULT CDataObject::QueryGetData(FORMATETC* pFormatEtc)
{
	/*
	Determines whether the data object is capable of rendering the data as specified
	*/
	return (LookupFormatEtc(pFormatEtc) == -1) ? DV_E_FORMATETC : S_OK;
}

HRESULT CDataObject::GetCanonicalFormatEtc(FORMATETC* pFormatEct, FORMATETC* pFormatEtcOut)
{
	/*
	Provides a potentially different but logically equivalent FORMATETC structure
	*/
	if (pFormatEtcOut == NULL)
		return E_INVALIDARG;
	return DATA_S_SAMEFORMATETC;
}

HRESULT CDataObject::SetData(FORMATETC *pformatetc, STGMEDIUM *pmedium, BOOL fRelease)
{
	/*
	Called by an object containing a data source to transfer data to the object
	that implements this method
	*/
	if(pformatetc == NULL || pmedium == NULL)
		return E_INVALIDARG;

	m_pFormatEtc.push_back(*pformatetc);
	m_pStgMedium.push_back(*pmedium);
	return S_OK;
}

HRESULT CDataObject::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC** ppEnumFormatEtc)
{
	/*
	Creates an object to enumerate the formats supported by a data object.
	*/

	if(dwDirection == DATADIR_GET)
	{
		//to support all Windows platforms we need to implement IEnumFormatEtc ourselves.
		//return CreateEnumFormatEtc(m_nNumFormats, m_pFormatEtc, ppEnumFormatEtc);
		size_t formatSize = m_pFormatEtc.size();
		FORMATETC* pFormatEtc;
		pFormatEtc = new FORMATETC[formatSize];
		
		for(size_t i = 0; i < formatSize; ++i)
		{
			pFormatEtc[i] = m_pFormatEtc[i];
		}

		return CreateEnumFormatEtc(formatSize, pFormatEtc, ppEnumFormatEtc);
	}
	else
	{
		return E_NOTIMPL;
	}
}

HRESULT CDataObject::DAdvise(FORMATETC* pFormatEtc, DWORD advf, IAdviseSink* pAdvSink, DWORD* pdwConnection)
{
	/*
	Creates a connection between a data object and an advise sink so the advise sink can receive notifications of changes in the data object
	*/

	return OLE_E_ADVISENOTSUPPORTED;
}

HRESULT CDataObject::DUnadvise(DWORD dwConnection)
{
	//Destroys a notification previously set up with the DAdvise method.

	return OLE_E_ADVISENOTSUPPORTED;
}

HRESULT CDataObject::EnumDAdvise (IEnumSTATDATA** ppEnumAdvise)
{
	/*
	Creates and returns a pointer to an object to enumerate the current advisory connections
	*/
	return OLE_E_ADVISENOTSUPPORTED;
}


