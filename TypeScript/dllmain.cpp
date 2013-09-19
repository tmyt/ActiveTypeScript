#include<Windows.h>
#include"Common.h"
#include"TypeScriptFactory.h"

BOOL DllMain(HINSTANCE hInstance, DWORD fdwReason, LPVOID lpReserved)
{
	return TRUE;
}

STDAPI DllCanUnloadNow()
{
	return !ulLockCount ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	static CTypeScriptFactory factory;
	if(!ppv) return E_INVALIDARG;

	if(rclsid == CLSID_TypeScript)
		return factory.QueryInterface(riid, ppv);
	*ppv = nullptr;
	return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllRegisterServer()
{
	return E_NOTIMPL;
}

STDAPI DllUnregisterServer()
{
	return E_NOTIMPL;
}
