#include <iostream>
#include <Windows.h>
#include <WbemCli.h>

#pragma comment(lib, "wbemuuid.lib")
using namespace std;

int ShowFullInfoAboutKeyboard(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
    cout << "First: " << endl << endl;

    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_Keyboard"), WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* pclsObj;
    ULONG uReturn = 0;
    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
            &pclsObj, &uReturn);

        if (0 == uReturn)
        {
            break;
        }
        SAFEARRAY* sfArray;
        LONG lstart, lend;
        VARIANT vtProp;
        pclsObj->GetNames(0, WBEM_FLAG_ALWAYS, 0, &sfArray);
        hr = SafeArrayGetLBound(sfArray, 1, &lstart);
        if (FAILED(hr)) return hr;
        hr = SafeArrayGetUBound(sfArray, 1, &lend);
        if (FAILED(hr)) return hr;
        BSTR* pbstr;
        hr = SafeArrayAccessData(sfArray, (void HUGEP**) & pbstr);
        int nIdx = 0;
        if (SUCCEEDED(hr))
        {
            CIMTYPE pType;
            for (nIdx = lstart; nIdx < lend; nIdx++)
            {

                hr = pclsObj->Get(pbstr[nIdx], 0, &vtProp, &pType, 0);
                if (vtProp.vt == VT_NULL)
                {
                    continue;
                }
                if (pType == CIM_STRING && pType != CIM_EMPTY && pType != CIM_ILLEGAL)
                {
                    wcout << " OS Name : " << nIdx << vtProp.bstrVal << endl;
                }

                VariantClear(&vtProp);

            }
            hr = SafeArrayUnaccessData(sfArray);
            if (FAILED(hr)) return hr;
        }


        // Get the value of the Name property
        /*hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << " OS Name : " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);*/

        pclsObj->Release();

        cout << endl;
    }
}

int ShowDescriptionAndNumberOfFunctionKeysOfKeyboard(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
    cout << "Second: " << endl << endl;

    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_Keyboard"), WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* clsObj = NULL;
    int numElems;
    while ((hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems)) != WBEM_S_FALSE)
    {
        if (FAILED(hRes)) {
            break;
        }
        VARIANT vRet;
        VariantInit(&vRet);
        if (SUCCEEDED(clsObj->Get(L"Description", 0, &vRet, NULL, NULL)))
        {
            std::wcout << L"Description: " << vRet.bstrVal << endl;
            VariantClear(&vRet);
        }
        if (SUCCEEDED(clsObj->Get(L"NumberOfFunctionKeys", 0, &vRet, NULL, NULL)))
        {
            std::wcout << L"Number of function keys: " << vRet.uintVal << endl;
            VariantClear(&vRet);
        }

        clsObj->Release();
    }
    pEnumerator->Release();

    cout << endl;
}

int main()
{
    //First
    HRESULT hRes = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hRes)) {
        cout << "Unable to launch COM: 0x" << std::hex << hRes << endl;
        return 1;
    }
    if ((FAILED(hRes = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_CONNECT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, 0))))
    {
        cout << "Unable to initialize security: 0x" << std::hex << hRes << endl;
        return 1;
    }
    //Second
    IWbemLocator* pLocator = NULL;
    if (FAILED(hRes = CoCreateInstance(CLSID_WbemLocator, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pLocator)))) {
        cout << "Unable to create a WbemLocator: " << std::hex << hRes << endl;
        return 1;
    }
    //Third
    IWbemServices* pService = NULL;
    if (FAILED(hRes = pLocator->ConnectServer(BSTR(L"root\\CIMV2"), NULL, NULL, NULL, WBEM_FLAG_CONNECT_USE_MAX_WAIT, NULL, NULL, &pService))) {
        pLocator->Release();
        cout << "Unable to connect to \"CIMV2\": " << std::hex << hRes << endl;
        return 1;
    }

    //Fourth
    ShowFullInfoAboutKeyboard(hRes, pLocator, pService);
    ShowDescriptionAndNumberOfFunctionKeysOfKeyboard(hRes, pLocator, pService);

    // Fifth
    pService->Release();
    pLocator->Release();

    system("pause");
    return 0;
}

