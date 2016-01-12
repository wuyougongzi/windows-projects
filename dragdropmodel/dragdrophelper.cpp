#include "stdafx.h"
#include "dragdrophelper.h"
#include <ShlObj.h>

CDragDropHelper* CDragDropHelper::m_pInstance = NULL;

CDragDropHelper::CDragDropHelper()
{

}

CDragDropHelper::~CDragDropHelper()
{

}

CDragDropHelper* CDragDropHelper::GetInstance()
{
	if (m_pInstance == NULL)
	{
		m_pInstance = new CDragDropHelper;
	}
	return m_pInstance;
}

void CDragDropHelper::DoDragDrop(HGLOBAL hGlobal)
{
	IDataObject* pDataObject = NULL;
	IDropSource* pDropSource = NULL;
	DWORD		 dwEffect;

	FORMATETC fmtetc = { CF_HDROP, 0, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	STGMEDIUM stgmed = { TYMED_HGLOBAL, { 0 }, 0 };
	stgmed.hGlobal = hGlobal;

	pDropSource = new CDropSource();
	pDataObject = new CDataObject(&fmtetc, &stgmed, 1);

	DoDragDrop(pDataObject, pDropSource, DROPEFFECT_COPY, &dwEffect);
	pDataObject->Release();
	pDropSource->Release();
}

void CDragDropHelper::DoDragDrop(HGLOBAL hGlobal, HBITMAP hbitmap)
{
	IDataObject* pDataObject = NULL;
	IDropSource* pDropSource = NULL;

	FORMATETC fmtetc = { CF_BITMAP, 0, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	STGMEDIUM stgmed = { TYMED_HGLOBAL, { 0 }, 0 };
	stgmed.hGlobal = hGlobal;

	pDropSource = new CDropSource();
	pDataObject = new CDataObject(&fmtetc, &stgmed, 1);

	BITMAP bm;
	::GetObject(hbitmap, sizeof(bm), &bm);
	HBITMAP hDragBitmap = (HBITMAP)::CopyImage(hbitmap, IMAGE_BITMAP, bm.bmWidth, bm.bmHeight, 0);

	IDragSourceHelper* pDragSourceHelper = NULL;
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_IDragSourceHelper, (LPVOID*)&pDragSourceHelper);
	SHDRAGIMAGE info;
	info.sizeDragImage.cx = bm.bmWidth;
	info.sizeDragImage.cy = bm.bmHeight;
	info.ptOffset.x = bm.bmWidth / 2;
	info.ptOffset.y = bm.bmHeight / 2;

	info.hbmpDragImage = hDragBitmap;

	HRESULT hr = pDragSourceHelper->InitializeFromBitmap(&info, pDataObject);
	if(FAILED(hr))
	{
		DeleteObject(info.hbmpDragImage);
	}

	DWORD dwEffect = 0;
	DoDragDrop(pDataObject, pDropSource, DROPEFFECT_COPY, &dwEffect);

	pDropSource->Release();
	pDataObject->Release();
}