#include<Windows.h>

BOOL DllMain(HINSTANCE hInstance, DWORD fdwReason, LPVOID lpReserved)
{
	return TRUE;
}

STDAPI DllCanUnloadNow()
{
	return E_NOTIMPL;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return E_NOTIMPL;
}

STDAPI DllRegisterServer()
{
	return E_NOTIMPL;
}

STDAPI DllUnregisterServer()
{
	return E_NOTIMPL;
}