
#include "stdafx.h"
#include "droptarget.h"
#include <ShObjIdl.h>

CDropTarget::CDropTarget()
{
	m_lRefCount = 1;
	m_pDropTargetHelper = NULL;
	if ( FAILED( CoCreateInstance( CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_IDropTargetHelper, (LPVOID*)&m_pDropTargetHelper ) ) )
		m_pDropTargetHelper = NULL;
}

CDropTarget::~CDropTarget()
{
	if(m_pDropTargetHelper != NULL)
	{
		m_pDropTargetHelper->Release();
		m_pDropTargetHelper = NULL;
	}
}

HRESULT CDropTarget::QueryInterface(REFIID riid, void **ppvObject)
{
	if(riid == IID_IDropTarget || riid == IID_IUnknown)
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

ULONG CDropTarget::AddRef()
{
	return InterlockedIncrement(&m_lRefCount);
}

ULONG CDropTarget::Release()
{
	LONG count = InterlockedDecrement(&m_lRefCount);
	if(count == 0)
	{
		delete this;
		return 0;
	}
	return count;
}

HRESULT CDropTarget::DragEnter(__RPC__in_opt IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, __RPC__inout DWORD *pdwEffect)
{
	if(pDataObj == NULL)
		return E_INVALIDARG;

	if(m_pDropTargetHelper != NULL)
		m_pDropTargetHelper->DragEnter(NULL, pDataObj, (LPPOINT)&pt, *pdwEffect);

	m_bAllowDrop = QueryDataObject(pDataObj);
	if(m_bAllowDrop)
	{
		*pdwEffect = DROPEFFECT_COPY;
		DoDragEnter(pt);
	}
	else
	{
		*pdwEffect = DROPEFFECT_NONE;
	}

	return S_OK;
}

HRESULT CDropTarget::DragOver(DWORD grfKeyState, POINTL pt, __RPC__inout DWORD *pdwEffect)
{
	if ( m_pDropTargetHelper )
	{
		m_pDropTargetHelper->DragOver( (LPPOINT)&pt, *pdwEffect );
		m_pDropTargetHelper->Show(true);
	}

	if(m_bAllowDrop)
	{
		*pdwEffect = DROPEFFECT_COPY;
		DoDragOver(pt);
	}
	else
	{
		*pdwEffect = DROPEFFECT_NONE;
	}

	return S_OK;
}

HRESULT CDropTarget::DragLeave()
{
	if ( m_pDropTargetHelper )
		m_pDropTargetHelper->DragLeave();

	DoDragLeave();
	return S_OK;
}

void CDropTarget::DropData(IDataObject* pDataObject, POINTL pt)
{
	FORMATETC fmtect = {CF_BITMAP, 0, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
	STGMEDIUM stgmed;

	if(pDataObject->QueryGetData(&fmtect) == S_OK)
	{
		if(pDataObject->GetData(&fmtect, &stgmed) == S_OK)
		{
			DoDrop(stgmed.hGlobal, pt);
			ReleaseStgMedium(&stgmed);
		}
	}
}

HRESULT CDropTarget::Drop(__RPC__in_opt IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, __RPC__inout DWORD *pdwEffect)
{
	if(pDataObj == NULL)
		return E_INVALIDARG;

	if ( m_pDropTargetHelper )
		m_pDropTargetHelper->Drop( pDataObj, (LPPOINT)&pt, *pdwEffect );

	if(m_bAllowDrop)
	{
		*pdwEffect = DROPEFFECT_COPY;
		DropData(pDataObj, pt);
	}
	else
	{
		*pdwEffect = DROPEFFECT_NONE;
	}
	return S_OK;
}

bool  CDropTarget::QueryDataObject(IDataObject *pDataObject)
{
	FORMATETC fmtetc = {CF_BITMAP, 0 , DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
	return pDataObject->QueryGetData(&fmtetc) == S_OK ? true : false;
}


bool CDropTarget::DoDragEnter(POINTL pt)
{
	//to implement
	return true;
}

bool CDropTarget::DoDragOver(POINTL pt)
{
	//to implement
	return true;
}

bool CDropTarget::DoDragLeave()
{
	//to implement
	return true;
}

bool CDropTarget::DoDrop(HGLOBAL hGolbal, POINTL pt)
{
	//to implement
	return true;
}
