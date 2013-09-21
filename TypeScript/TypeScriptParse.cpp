#include<Windows.h>
#include<Shlwapi.h>

#include"TypeScriptParse.h"
#include"Common.h"

CTypeScriptParse::CTypeScriptParse(IActiveScript* js) : m_ref(1)
{
	InterlockedIncrement(&ulLockCount);
	js->QueryInterface(IID_PPV_ARGS(&m_jsparse));
}

CTypeScriptParse::~CTypeScriptParse()
{
	m_jsparse->Release();

	InterlockedDecrement(&ulLockCount);
}

/** IUnknown **/
HRESULT STDMETHODCALLTYPE CTypeScriptParse::QueryInterface(
	/* [in] */ REFIID riid,
	/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
{
	if (riid == IID_IUnknown || riid == IID_IActiveScriptParse)
	{
		AddRef();
		*ppvObject = this;
		return S_OK;
	}
	return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE CTypeScriptParse::AddRef(void)
{
	return ++m_ref;
}

ULONG STDMETHODCALLTYPE CTypeScriptParse::Release(void)
{
	if (--m_ref == 0){
		delete this;
		return 0;
	}
	return m_ref;
}

/** IActiveScriptParse **/
HRESULT STDMETHODCALLTYPE CTypeScriptParse::InitNew(void)
{
	return m_jsparse->InitNew();
}

HRESULT STDMETHODCALLTYPE CTypeScriptParse::AddScriptlet(
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
	/* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CTypeScriptParse::ParseScriptText(
	/* [in] */ __RPC__in LPCOLESTR pstrCode,
	/* [in] */ __RPC__in LPCOLESTR pstrItemName,
	/* [in] */ __RPC__in_opt IUnknown *punkContext,
	/* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
	/* [in] */ COOKIETYPE dwSourceContextCookie,
	/* [in] */ ULONG ulStartingLineNumber,
	/* [in] */ DWORD dwFlags,
	/* [out] */ __RPC__out VARIANT *pvarResult,
	/* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
{
	int n = WideCharToMultiByte(CP_UTF8, 0, pstrCode, -1, 0, 0, 0, 0);
	char* buffer = new char[n + 1];
	WideCharToMultiByte(CP_UTF8, 0, pstrCode, -1, buffer, n, 0, 0);

	TCHAR path[MAX_PATH], dll[MAX_PATH], out[MAX_PATH];
	GetTempPath(MAX_PATH, path);
	GetTempFileName(path, TEXT("xx"), 0, path);

	DWORD w;
	HANDLE hFile = CreateFile(path, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, 0, 0);
	WriteFile(hFile, buffer, n - 1, &w, 0);
	CloseHandle(hFile);
	delete [] buffer;
	PathQuoteSpaces(path);
	StrCpy(out, path);
	StrCat(out, TEXT(".js"));

	GetModuleFileName(hInstance, dll, MAX_PATH);
	PathRemoveFileSpec(dll);
	PathCombine(dll, dll, TEXT("lib\\tsc.exe"));
	PathQuoteSpaces(dll);

	TCHAR args[MAX_PATH * 4];
	wsprintf(args, TEXT("%s --noResolve --out %s %s"), dll, out, path);

	PROCESS_INFORMATION pi;
	STARTUPINFO si = {0};
	CreateProcess(0, args, 0, 0, 0, 0, 0, 0, &si, &pi);
	WaitForSingleObject(pi.hProcess, INFINITE);
	CloseHandle(pi.hProcess);

	DeleteFile(path);
	hFile = CreateFile(out, GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0);
	DWORD size = GetFileSize(hFile, nullptr);
	buffer = new char[size + 1];
	memset(buffer, 0, size + 1);
	ReadFile(hFile, buffer, size, &w, nullptr);
	CloseHandle(hFile);

	n = MultiByteToWideChar(CP_UTF8, 0, buffer, size, 0, 0);
	wchar_t * wbuf = new wchar_t[n + 1];
	MultiByteToWideChar(CP_UTF8, 0, buffer, -1, wbuf, n);
	delete [] buffer;

	HRESULT hr = m_jsparse->ParseScriptText(wbuf, pstrItemName, punkContext, pstrDelimiter, dwSourceContextCookie,
		ulStartingLineNumber, dwFlags, pvarResult, pexcepinfo);

	delete [] wbuf;
	DeleteFile(out);

	return hr;
}
