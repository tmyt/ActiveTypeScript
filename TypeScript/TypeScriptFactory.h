#pragma once

#include<Windows.h>

// {11369DE6-251F-445e-A44E-1A694D638F43}
DEFINE_GUID(CLSID_TypeScript,
	0x11369de6, 0x251f, 0x445e, 0xa4, 0x4e, 0x1a, 0x69, 0x4d, 0x63, 0x8f, 0x43);

class CTypeScriptFactory : public IClassFactory
{
private:
	ULONG m_ref;

public:
	CTypeScriptFactory() : m_ref(1) {}

	/** IUnknown **/
	HRESULT STDMETHODCALLTYPE QueryInterface(
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);

	ULONG STDMETHODCALLTYPE AddRef(void);

	ULONG STDMETHODCALLTYPE Release(void);

	/** IClassFactory **/
	HRESULT STDMETHODCALLTYPE CreateInstance(
		/* [annotation][unique][in] */
		_In_opt_  IUnknown *pUnkOuter,
		/* [annotation][in] */
		_In_  REFIID riid,
		/* [annotation][iid_is][out] */
		_COM_Outptr_  void **ppvObject);

	HRESULT STDMETHODCALLTYPE LockServer(
		/* [in] */ BOOL fLock);
};
