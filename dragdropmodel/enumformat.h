#ifndef ENUMFORMAT_H
#define ENUMFORMAT_H

typedef std::vector<FORMATETC> FormatEtcArray;
typedef std::vector<FORMATETC*> PFormatEtcArray;
typedef std::vector<STGMEDIUM*> PStgMediumArray;

class CEnumFormatEtc : public IEnumFORMATETC
{
public:
	CEnumFormatEtc(FORMATETC* pFormatEtc, int nNumFormats);
	~CEnumFormatEtc();

public:
	HRESULT	__stdcall QueryInterface(REFIID iid, void** ppvObject);
	ULONG	__stdcall AddRef(void);
	ULONG	__stdcall Release(void);

	HRESULT __stdcall Next(ULONG celt, FORMATETC* pFormatEtc, ULONG* pceltFetched);
	HRESULT __stdcall Skip(ULONG celt);
	HRESULT __stdcall Reset(void);
	HRESULT __stdcall Clone(IEnumFORMATETC** ppEnumFormatEtc);

private:
	LONG		m_lRefCount;
	ULONG		m_nIndex;			// current enumerator index
	ULONG		m_nNumFormats;		// number of FORMATETC members
	FORMATETC*	m_pFormatEtc;		// array of FORMATETC objects
};

#endif
