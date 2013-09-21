#include<Windows.h>
#include<Shlwapi.h>
#include<tchar.h>

#define INITGUID
#include<guiddef.h>
#include"Common.h"
#include"TypeScriptFactory.h"

#pragma comment(lib, "shlwapi.lib")

HINSTANCE hInstance;

BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData);

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD fdwReason, LPVOID lpReserved)
{
	::hInstance = hInstance;
	return TRUE;
}

STDAPI DllCanUnloadNow()
{
	return !ulLockCount ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	static CTypeScriptFactory factory;
	if (!ppv) return E_INVALIDARG;

	if (rclsid == CLSID_TypeScript)
		return factory.QueryInterface(riid, ppv);
	*ppv = nullptr;
	return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllRegisterServer()
{
	TCHAR szKey[256];
	TCHAR szModulePath[MAX_PATH];

	wsprintf(szKey, _T("CLSID\\%s"), szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("TypeScript Language")))
		return E_FAIL;

	GetModuleFileName(hInstance, szModulePath, sizeof(szModulePath) / sizeof(TCHAR));
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

	wsprintf(szKey, TEXT("CLSID\\%s\\Implemented Category"), szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, NULL))
		return E_FAIL;

	wsprintf(szKey, TEXT("CLSID\\%s\\Implemented Category\\{F0B7A1A2-9847-11CF-8F20-00805F2CD064}"), szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, NULL))
		return E_FAIL;

	wsprintf(szKey, TEXT("%s"), szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("TypeScript Language")))
		return E_FAIL;

	wsprintf(szKey, TEXT("%s\\CLSID"), szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, (LPTSTR) szClsid))
		return E_FAIL;

	wsprintf(szKey, TEXT("%s\\OLEScript"), szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, NULL))
		return E_FAIL;

	wsprintf(szKey, TEXT("%s\\Implemented Category"), szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, NULL))
		return E_FAIL;

	wsprintf(szKey, TEXT("%s\\Implemented Category\\{F0B7A1A2-9847-11CF-8F20-00805F2CD064}"), szProgid);
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

BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData)
{
	HKEY  hKey;
	LONG  lResult;
	DWORD dwSize;

	lResult = RegCreateKeyEx(hKeyRoot, lpszKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE|KEY_CREATE_SUB_KEY, NULL, &hKey, NULL);
	if (lResult != ERROR_SUCCESS)
		return FALSE;

	if (lpszData != NULL)
	{
		dwSize = (lstrlen(lpszData) + 1) * sizeof(TCHAR);
		RegSetValueEx(hKey, lpszValue, 0, REG_SZ, (LPBYTE) lpszData, dwSize);
	}
	RegCloseKey(hKey);

	return TRUE;
}