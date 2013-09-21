#pragma once

#include<Windows.h>
#include<ActivScp.h>

#ifdef _WIN64
#define COOKIETYPE DWORDLONG
#else
#define COOKIETYPE DWORD
#endif

class CTypeScriptParse : public IActiveScriptParse
{
private:
	ULONG m_ref;

	IActiveScriptParse* m_jsparse;

public:
	CTypeScriptParse(IActiveScript* js);
	virtual ~CTypeScriptParse();

	/** IUnknown **/
	HRESULT STDMETHODCALLTYPE QueryInterface(
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);

	ULONG STDMETHODCALLTYPE AddRef(void);

	ULONG STDMETHODCALLTYPE Release(void);

	/** IActiveScriptParse **/
	HRESULT STDMETHODCALLTYPE InitNew(void);

	HRESULT STDMETHODCALLTYPE AddScriptlet(
		/* [in] */ __RPC__in LPCOLESTR pstrDefaultName,
		/* [in] */ __RPC__in LPCOLESTR pstrCode,
		/* [in] */ __RPC__in LPCOLESTR pstrItemName,
		/* [in] */ __RPC__in LPCOLESTR pstrSubItemName,
		/* [in] */ __RPC__in LPCOLESTR pstrEventName,
		/* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
		/* [in] */ COOKIETYPE dwSourceContextCookie,
		/* [in] */ ULONG ulStartingLineNumber,
		/* [in] */ DWORD dwFlags,
		/* [out] */ __RPC__deref_out_opt BSTR *pbstrName,
		/* [out] */ __RPC__out EXCEPINFO *pexcepinfo);

	HRESULT STDMETHODCALLTYPE ParseScriptText(
		/* [in] */ __RPC__in LPCOLESTR pstrCode,
		/* [in] */ __RPC__in LPCOLESTR pstrItemName,
		/* [in] */ __RPC__in_opt IUnknown *punkContext,
		/* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
		/* [in] */ COOKIETYPE dwSourceContextCookie,
		/* [in] */ ULONG ulStartingLineNumber,
		/* [in] */ DWORD dwFlags,
		/* [out] */ __RPC__out VARIANT *pvarResult,
		/* [out] */ __RPC__out EXCEPINFO *pexcepinfo);
};