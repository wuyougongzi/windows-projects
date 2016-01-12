#ifndef DATAOBJECT_H
#define DATAOBJECT_H

#include "dropsource.h"

typedef std::vector<FORMATETC>	FormatEtcVec;
typedef std::vector<STGMEDIUM>	StgMediumVec;

class CDataObject : public IDataObject
{
public:
	CDataObject(FORMATETC* fmt, STGMEDIUM* stgmed, int count);
	~CDataObject();

public:
	HRESULT	__stdcall QueryInterface(REFIID iid, void** ppvObject);
	ULONG	__stdcall AddRef(void);
	ULONG	__stdcall Release(void);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetData( 
		/* [unique][in] */ FORMATETC *pformatetcIn,
		/* [out] */ STGMEDIUM *pmedium);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDataHere( 
		/* [unique][in] */ FORMATETC *pformatetc,
		/* [out][in] */ STGMEDIUM *pmedium);

	virtual HRESULT STDMETHODCALLTYPE QueryGetData( 
		/* [unique][in] */ __RPC__in_opt FORMATETC *pformatetc);

	virtual HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc( 
		/* [unique][in] */ __RPC__in_opt FORMATETC *pformatectIn,
		/* [out] */ __RPC__out FORMATETC *pformatetcOut);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE SetData( 
		/* [unique][in] */ FORMATETC *pformatetc,
		/* [unique][in] */ STGMEDIUM *pmedium,
		/* [in] */ BOOL fRelease);

	virtual HRESULT STDMETHODCALLTYPE EnumFormatEtc( 
		/* [in] */ DWORD dwDirection,
		/* [out] */ __RPC__deref_out_opt IEnumFORMATETC **ppenumFormatEtc);

	virtual HRESULT STDMETHODCALLTYPE DAdvise( 
		/* [in] */ __RPC__in FORMATETC *pformatetc,
		/* [in] */ DWORD advf,
		/* [unique][in] */ __RPC__in_opt IAdviseSink *pAdvSink,
		/* [out] */ __RPC__out DWORD *pdwConnection);

	virtual HRESULT STDMETHODCALLTYPE DUnadvise( 
		/* [in] */ DWORD dwConnection);

	virtual HRESULT STDMETHODCALLTYPE EnumDAdvise( 
		/* [out] */ __RPC__deref_out_opt IEnumSTATDATA **ppenumAdvise);

private:
	int LookupFormatEtc(FORMATETC* pFormatEtc);

private:
	LONG					m_lRefCount;
	std::vector<FORMATETC>	m_pFormatEtc;
	std::vector<STGMEDIUM>	m_pStgMedium;
};

#endif 
