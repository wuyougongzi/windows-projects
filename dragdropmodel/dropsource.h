
#ifndef DROPSOURCE_H
#define DROPSOURCE_H

class CDropSource : public IDropSource
{
public:
	CDropSource();
	~CDropSource();
public:
	HRESULT	__stdcall QueryInterface(REFIID iid, void** ppvObject);
	ULONG	__stdcall AddRef();
	ULONG	__stdcall Release();

	HRESULT __stdcall QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState);
	HRESULT __stdcall GiveFeedback(DWORD dwEffect);

private:
	LONG	m_lRefCount;
};

#endif