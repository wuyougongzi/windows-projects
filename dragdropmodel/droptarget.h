
#ifndef DROPTARGET_H
#define DROPTARGET_H

class CDropTarget : public IDropTarget
{
public:
	CDropTarget();
	virtual ~CDropTarget();

public:
	HRESULT __stdcall QueryInterface(REFIID riid, void __RPC_FAR *__RPC_FAR *ppvObject);
	ULONG __stdcall AddRef(void);
	ULONG __stdcall Release(void);

public:
	HRESULT __stdcall DragEnter(__RPC__in_opt IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, __RPC__inout DWORD *pdwEffect);
	HRESULT __stdcall DragOver(DWORD grfKeyState, POINTL pt, __RPC__inout DWORD *pdwEffect);
	HRESULT __stdcall DragLeave();
	HRESULT __stdcall Drop(__RPC__in_opt IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, __RPC__inout DWORD *pdwEffect);

protected:
	virtual bool DoDragEnter(POINTL pt);
	virtual bool DoDragOver(POINTL pt);
	virtual bool DoDragLeave();
	virtual bool DoDrop(HGLOBAL hGolbal, POINTL pt);

private:
	bool QueryDataObject(IDataObject *pDataObject);
	void DropData(IDataObject* pDataObject, POINTL pt);

private:
	LONG						m_lRefCount;
	bool						m_bAllowDrop;
	struct IDropTargetHelper*	m_pDropTargetHelper;
};

#endif