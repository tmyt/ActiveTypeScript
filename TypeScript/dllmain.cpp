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
	TCHAR szKey[256];
	wsprintf_s(szKey, _T("CLSID\\%s"), szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("TypeScript Language")))
		return E_FAIL;

	GetModuleFileName(g_hinstDll, szModulePath, sizeof(szModulePath) / sizeof(TCHAR));
	wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, szModulePath))
		return E_FAIL;
	
	wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, TEXT("ThreadingModel"), TEXT("Both")))
		return E_FAIL;
	
	wsprintf(szKey, TEXT("CLSID\\%s\\ProgID"), szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, szProgid))
		return E_FAIL;

	wsprintf(szKey, TEXT("CLSID\\%s\\OLEScript"), szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, NULL))
		return E_FAIL;

	wsprintf(szKey, TEXT("%s"), szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("TypeScript Language")))
		return E_FAIL;
  
	wsprintf(szKey, TEXT("%s\\CLSID"), szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, (LPTSTR)szClsid))
		return E_FAIL;

	wsprintf(szKey, TEXT("%s\\OLEScript"), szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, NULL))
		return E_FAIL;

	return S_OK;
}

STDAPI DllUnregisterServer()
{
TCHAR szKey[256];

	wsprintf(szKey, TEXT("CLSID\\%s"), szClsid);
	SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
	
	wsprintf(szKey, TEXT("%s"), szProgid);
	SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
	return S_OK;
}
