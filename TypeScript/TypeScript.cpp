#include<Windows.h>
#include<Shlwapi.h>

#include"TypeScript.h"
#include"TypeScriptParse.h"
#include"Common.h"

CTypeScript::CTypeScript() : m_ref(1)
{
	InterlockedIncrement(&ulLockCount);

	CLSID CLSID_Script;
	CLSIDFromProgID(L"JScript", &CLSID_Script);

	CoCreateInstance(CLSID_Script, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_js));
}

CTypeScript::~CTypeScript()
{
	m_js->Release();

	InterlockedDecrement(&ulLockCount);
}

/** IUnknown **/
HRESULT STDMETHODCALLTYPE CTypeScript::QueryInterface(
	/* [in] */ REFIID riid,
	/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
{
	if (riid == IID_IUnknown || riid == IID_IActiveScript)
	{
		AddRef();
		*ppvObject = this;
		return S_OK;
	}
	if (riid == IID_IActiveScriptParse){
		*ppvObject = new CTypeScriptParse(m_js);
		return S_OK;
	}
	return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE CTypeScript::AddRef(void)
{
	return ++m_ref;
}

ULONG STDMETHODCALLTYPE CTypeScript::Release(void)
{
	if (--m_ref == 0){
		delete this;
		return 0;
	}
	return m_ref;
}

/** IActiveScript **/
HRESULT STDMETHODCALLTYPE CTypeScript::SetScriptSite(
	/* [in] */ __RPC__in_opt IActiveScriptSite *pass)
{
	return m_js->SetScriptSite(pass);
}

HRESULT STDMETHODCALLTYPE CTypeScript::GetScriptSite(
	/* [in] */ __RPC__in REFIID riid,
	/* [iid_is][out] */ __RPC__deref_out_opt void **ppvObject)
{
	return m_js->GetScriptSite(riid, ppvObject);
}

HRESULT STDMETHODCALLTYPE CTypeScript::SetScriptState(
	/* [in] */ SCRIPTSTATE ss)
{
	return m_js->SetScriptState(ss);
}

HRESULT STDMETHODCALLTYPE CTypeScript::GetScriptState(
	/* [out] */ __RPC__out SCRIPTSTATE *pssState)
{
	return m_js->GetScriptState(pssState);
}

HRESULT STDMETHODCALLTYPE CTypeScript::Close(void)
{
	return m_js->Close();
}

HRESULT STDMETHODCALLTYPE CTypeScript::AddNamedItem(
	/* [in] */ __RPC__in LPCOLESTR pstrName,
	/* [in] */ DWORD dwFlags)
{
	return m_js->AddNamedItem(pstrName, dwFlags);
}

HRESULT STDMETHODCALLTYPE CTypeScript::AddTypeLib(
	/* [in] */ __RPC__in REFGUID rguidTypeLib,
	/* [in] */ DWORD dwMajor,
	/* [in] */ DWORD dwMinor,
	/* [in] */ DWORD dwFlags)
{
	return m_js->AddTypeLib(rguidTypeLib, dwMajor, dwMinor, dwFlags);
}

HRESULT STDMETHODCALLTYPE CTypeScript::GetScriptDispatch(
	/* [in] */ __RPC__in LPCOLESTR pstrItemName,
	/* [out] */ __RPC__deref_out_opt IDispatch **ppdisp)
{
	return m_js->GetScriptDispatch(pstrItemName, ppdisp);
}

HRESULT STDMETHODCALLTYPE CTypeScript::GetCurrentScriptThreadID(
	/* [out] */ __RPC__out SCRIPTTHREADID *pstidThread)
{
	return m_js->GetCurrentScriptThreadID(pstidThread);
}

HRESULT STDMETHODCALLTYPE CTypeScript::GetScriptThreadID(
	/* [in] */ DWORD dwWin32ThreadId,
	/* [out] */ __RPC__out SCRIPTTHREADID *pstidThread)
{
	return m_js->GetScriptThreadID(dwWin32ThreadId, pstidThread);
}

HRESULT STDMETHODCALLTYPE CTypeScript::GetScriptThreadState(
	/* [in] */ SCRIPTTHREADID stidThread,
	/* [out] */ __RPC__out SCRIPTTHREADSTATE *pstsState)
{
	return m_js->GetScriptThreadState(stidThread, pstsState);
}

HRESULT STDMETHODCALLTYPE CTypeScript::InterruptScriptThread(
	/* [in] */ SCRIPTTHREADID stidThread,
	/* [in] */ __RPC__in const EXCEPINFO *pexcepinfo,
	/* [in] */ DWORD dwFlags)
{
	return m_js->InterruptScriptThread(stidThread, pexcepinfo, dwFlags);
}

HRESULT STDMETHODCALLTYPE CTypeScript::Clone(
	/* [out] */ __RPC__deref_out_opt IActiveScript **ppscript)
{
	return E_NOTIMPL;
}
