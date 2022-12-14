#include "task05.h"
#include "pch.h"
#include "globals.h"

HRESULT Task05_01(VOID)
{
	HRESULT hr = S_OK;
	IEnumWbemClassObject* pEnum = NULL;
	IWbemClassObject* pObj = NULL;
	IWbemClassObject* pClsDef = NULL;
	IWbemClassObject* pClsInParam = NULL;
	IWbemClassObject* pClsInParamInst = NULL;
	static LPCTSTR lpszMethod = _T("Terminate");
	static LPCTSTR lpszClass = _T("Win32_Process");
	BSTR bszClsMoniker = NULL;
	VARIANT v;
	VariantInit(&v);

	std::tcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");

	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		(BSTR)_T("SELECT * ")
		_T("FROM Win32_Process ")
		_T("WHERE Name='notepad.exe' AND Priority='4'"),
		0,
		NULL, &pEnum
	);
	if (FAILED(hr))
		goto fail;

	hr = pSvc->GetObject(
		(BSTR)lpszClass, 0,
		NULL, &pClsDef, NULL
	);

	hr = pClsDef->GetMethod(
		lpszMethod, 0,
		&pClsInParam, NULL
	);

	hr = pClsInParam->SpawnInstance(0, &pClsInParamInst);

	V_VT(&v) = VT_UI4;
	V_UI4(&v) = 0;
	pClsInParamInst->Put(_T("Reason"), 0, &v, CIM_UINT32);

	while (1) {
		ULONG uRet = 0;
		pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
		if (uRet == 0)
			break;
		pObj->Get(
			_T("Handle"), 0,
			&v, 0, 0
		);
		bszClsMoniker = SysAllocString(_T("Win32_Process.Handle='"));
		VarBstrCat(bszClsMoniker, V_BSTR(&v), &bszClsMoniker);
		VarBstrCat(bszClsMoniker, (BSTR)_T("'"), &bszClsMoniker);

		hr = pSvc->ExecMethod(
			bszClsMoniker, (BSTR)lpszMethod, 0,
			NULL, pClsInParamInst, NULL, NULL
		);
		SysFreeString(bszClsMoniker);
		if (FAILED(hr)) {
			_tprintf_s(_T("ExecMethod failed, hr: %lX\n"), hr);
			goto fail;
		}
	}

	goto fail;
fail:
	VariantClear(&v);
	pClsInParamInst->Release();
	pClsInParam->Release();
	pClsDef->Release();
	pObj->Release();
	pEnum->Release();
	return hr;
}

static HRESULT Task05_02(VOID)
{
	HRESULT hr = S_OK;
	IEnumWbemClassObject* pEnum = NULL;
	IWbemClassObject* pObj = NULL;
	IWbemClassObject* pClsDef = NULL;
	IWbemClassObject* pClsInParam = NULL;
	IWbemClassObject* pClsInParamInst = NULL;
	static LPCTSTR lpszMethod = _T("Terminate");
	static LPCTSTR lpszClass = _T("Win32_Process");
	BSTR bszWQLQueryChild = NULL;
	VARIANT v;
	VariantInit(&v);

	std::tcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");

	hr = pSvc->GetObject(
		(BSTR)lpszClass, 0,
		NULL, &pClsDef, NULL
	);

	hr = pClsDef->GetMethod(
		lpszMethod, 0,
		&pClsInParam, NULL
	);

	hr = pClsInParam->SpawnInstance(0, &pClsInParamInst);

	V_VT(&v) = VT_UI4;
	V_UI4(&v) = 0;
	pClsInParamInst->Put(_T("Reason"), 0, &v, CIM_UINT32);

	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		(BSTR)_T("SELECT * ")
		_T("FROM Win32_Process ")
		_T("WHERE Name='totalcmd.exe' OR Name='totalcmd64.exe'"),
		0,
		NULL, &pEnum
	);
	if (FAILED(hr))
		goto fail;

	while (1) {
		IEnumWbemClassObject* pEnumChild = NULL;
		IWbemClassObject* pObjChild = NULL;
		ULONG uRet = 0;
		bszWQLQueryChild = SysAllocString(
			_T("SELECT * ")
			_T("FROM Win32_Process ")
			_T("WHERE ParentProcessId=")
		);
		pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
		if (uRet == 0)
			break;
		pObj->Get(
			_T("Handle"), 0,
			&v, 0, 0
		);

		VarBstrCat(bszWQLQueryChild, V_BSTR(&v), &bszWQLQueryChild);

		hr = pSvc->ExecQuery(
			(BSTR)_T("WQL"),
			bszWQLQueryChild,
			0,
			NULL, &pEnumChild
		);
		if (FAILED(hr))
			goto fail;

		while (1) {
			ULONG uRet = 0;
			pEnumChild->Next(WBEM_INFINITE, 1, &pObjChild, &uRet);
			if (uRet == 0)
				break;
			pObjChild->Get(
				_T("__PATH"), 0,
				&v, 0, 0
			);

			hr = pSvc->ExecMethod(
				V_BSTR(&v), (BSTR)lpszMethod, 0,
				NULL, pClsInParamInst, NULL, NULL
			);
			if (FAILED(hr)) {
				_tprintf_s(_T("ExecMethod failed, hr: %lX\n"), hr);
				goto fail;
			}
		}
		SysFreeString(bszWQLQueryChild);
	}

fail:
	SysFreeString(bszWQLQueryChild);
	VariantClear(&v);
	pClsInParamInst->Release();
	pClsInParam->Release();
	pClsDef->Release();
	pObj->Release();
	pEnum->Release();
	return hr;
}

HRESULT Task05(VOID)
{
	HRESULT hr = S_OK;
	//hr = Task05_01();
	hr = Task05_02();
	return hr;
}