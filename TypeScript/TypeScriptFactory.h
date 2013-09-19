#pragma once

#include<Windows.h>

// {A780C1EE-7853-4CFE-A9C6-92C9800EE7C7}
DEFINE_GUID(CLSID_TypeScript, 
		0xa780c1ee, 0x7853, 0x4cfe, 0xa9, 0xc6, 0x92, 0xc9, 0x80, 0xe, 0xe7, 0xc7);
#define szClsid _T("{A780C1EE-7853-4CFE-A9C6-92C9800EE7C7}")
#define szProgid _T("TypeScript")

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
