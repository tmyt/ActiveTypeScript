#pragma once

#include<Windows.h>
#include<ActivScp.h>

#ifdef _WIN64
#define COOKIETYPE DWORDLONG
#else
#define COOKIETYPE DWORD
#endif

class CTypeScript : public IActiveScript, IActiveScriptParse
{
private:
	ULONG m_ref;

	IActiveScript* m_js;
	IActiveScriptParse* m_jsparse;

public:
	CTypeScript();
	virtual ~CTypeScript();

	/** IUnknown **/
	HRESULT STDMETHODCALLTYPE QueryInterface(
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);

	ULONG STDMETHODCALLTYPE AddRef(void);

	ULONG STDMETHODCALLTYPE Release(void);

	/** IActiveScript **/
	HRESULT STDMETHODCALLTYPE SetScriptSite(
		/* [in] */ __RPC__in_opt IActiveScriptSite *pass);

	HRESULT STDMETHODCALLTYPE GetScriptSite(
		/* [in] */ __RPC__in REFIID riid,
		/* [iid_is][out] */ __RPC__deref_out_opt void **ppvObject);

	HRESULT STDMETHODCALLTYPE SetScriptState(
		/* [in] */ SCRIPTSTATE ss);

	HRESULT STDMETHODCALLTYPE GetScriptState(
		/* [out] */ __RPC__out SCRIPTSTATE *pssState);

	HRESULT STDMETHODCALLTYPE Close(void);

	HRESULT STDMETHODCALLTYPE AddNamedItem(
		/* [in] */ __RPC__in LPCOLESTR pstrName,
		/* [in] */ DWORD dwFlags);

	HRESULT STDMETHODCALLTYPE AddTypeLib(
		/* [in] */ __RPC__in REFGUID rguidTypeLib,
		/* [in] */ DWORD dwMajor,
		/* [in] */ DWORD dwMinor,
		/* [in] */ DWORD dwFlags);

	HRESULT STDMETHODCALLTYPE GetScriptDispatch(
		/* [in] */ __RPC__in LPCOLESTR pstrItemName,
		/* [out] */ __RPC__deref_out_opt IDispatch **ppdisp);

	HRESULT STDMETHODCALLTYPE GetCurrentScriptThreadID(
		/* [out] */ __RPC__out SCRIPTTHREADID *pstidThread);

	HRESULT STDMETHODCALLTYPE GetScriptThreadID(
		/* [in] */ DWORD dwWin32ThreadId,
		/* [out] */ __RPC__out SCRIPTTHREADID *pstidThread);

	HRESULT STDMETHODCALLTYPE GetScriptThreadState(
		/* [in] */ SCRIPTTHREADID stidThread,
		/* [out] */ __RPC__out SCRIPTTHREADSTATE *pstsState);

	HRESULT STDMETHODCALLTYPE InterruptScriptThread(
		/* [in] */ SCRIPTTHREADID stidThread,
		/* [in] */ __RPC__in const EXCEPINFO *pexcepinfo,
		/* [in] */ DWORD dwFlags);

	HRESULT STDMETHODCALLTYPE Clone(
		/* [out] */ __RPC__deref_out_opt IActiveScript **ppscript);

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