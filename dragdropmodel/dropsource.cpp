
#include "stdafx.h"
#include "dropsource.h"

CDropSource::CDropSource()
{
	m_lRefCount = 1;
}

CDropSource::~CDropSource()
{
}

ULONG CDropSource::AddRef(void)
{
	return InterlockedIncrement(&m_lRefCount);
}

ULONG CDropSource::Release(void)
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

HRESULT CDropSource::QueryInterface(REFIID iid, void** ppvObject)
{
	if(iid == IID_IDropSource || iid == IID_IUnknown)
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

HRESULT CDropSource::QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState)
{
	if(fEscapePressed == TRUE)
		return DRAGDROP_S_CANCEL;

	if((grfKeyState & MK_LBUTTON) == 0)
		return DRAGDROP_S_DROP;

	return S_OK;
}

HRESULT CDropSource::GiveFeedback(DWORD dwEffect)
{
	return DRAGDROP_S_USEDEFAULTCURSORS;
}

HRESULT CreateDropSource(IDropSource** ppDropSource)
{
	if(ppDropSource == NULL)
		return E_INVALIDARG;

	*ppDropSource = new CDropSource();
	return (*ppDropSource) ? S_OK : E_OUTOFMEMORY;
}
