#include"TypeScriptFactory.h"
#include"TypeScrpit.h"

DWORD g_lLocks;

/** IUnknown **/
HRESULT STDMETHODCALLTYPE CTypeScriptFactory::QueryInterface(
	REFIID riid,
	_COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
{
	if (riid == IID_IUnknown || riid == IID_IClassFactory){
		AddRef();
		*ppvObject = this;
		return S_OK;
	}
	return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE CTypeScriptFactory::AddRef(void)
{
	return ++m_ref;
}

ULONG STDMETHODCALLTYPE CTypeScriptFactory::Release(void)
{
	return --m_ref;
}

/** IClassFactory **/
HRESULT STDMETHODCALLTYPE CTypeScriptFactory::CreateInstance(
	_In_opt_  IUnknown *pUnkOuter,
	_In_  REFIID riid,
	_COM_Outptr_  void **ppvObject)
{
	if (pUnkOuter != nullptr)
		return CLASS_E_NOAGGREGATION;

	auto p = new CTypeScript();
	if (p == nullptr)
		return E_OUTOFMEMORY;

	auto hr = p->QueryInterface(riid, ppvObject);
	p->Release();

	return hr;
}

HRESULT STDMETHODCALLTYPE CTypeScriptFactory::LockServer(BOOL fLock)
{
	if (fLock)
		InterlockedIncrement(&g_lLocks);
	else
		InterlockedDecrement(&g_lLocks);
	return S_OK;
}