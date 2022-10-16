#include <iostream>
#include <Windows.h>
#include <WbemCli.h>
#include <vector>
#include <Psapi.h>
#include <stdio.h>
#include <sstream>

#pragma comment(lib, "wbemuuid.lib")
using namespace std;

std::string ProcessIdToName(DWORD processId)
{
    std::string ret;
    HANDLE handle = OpenProcess(
        PROCESS_QUERY_LIMITED_INFORMATION,
        FALSE,
        processId /* This is the PID, you can find one from windows task manager */
    );
    if (handle)
    {
        DWORD buffSize = 1024;
        CHAR buffer[1024];
        if (QueryFullProcessImageNameA(handle, 0, buffer, &buffSize))
        {
            ret = buffer;
        }
        else
        {
            printf("Error GetModuleBaseNameA : %lu", GetLastError());
        }
        CloseHandle(handle);
    }
    else
    {
        printf("Error OpenProcess : %lu", GetLastError());
    }
    return ret;
}

int GetInfoAboutProcByReadingSize() {
    // Get the list of process identifiers.

    DWORD aProcesses[1024], cbNeeded, cProcesses;
    unsigned int i;

    if (!EnumProcesses(aProcesses, sizeof(aProcesses), &cbNeeded))
    {
        return 1;
    }

    // Calculate how many process identifiers were returned.

    cProcesses = cbNeeded / sizeof(DWORD);

    // Print the memory usage for each process
    std::vector<PROCESS_MEMORY_COUNTERS> psArr;
    std::vector<DWORD> psID;
    for (i = 0; i < cProcesses; i++)
    {
        DWORD processID = aProcesses[i];
        HANDLE hProcess;
        PROCESS_MEMORY_COUNTERS pmc;


        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION |
            PROCESS_VM_READ,
            FALSE, processID);
        if (FAILED(hProcess))
            return 1;

        if (GetProcessMemoryInfo(hProcess, &pmc, sizeof(pmc)))
        {
            psArr.push_back(pmc);
            psID.push_back(processID);
        }

        CloseHandle(hProcess);
    }
    auto maxReadedSize = psArr[0];
    auto maxID = psID[0];
    for (size_t i = 1; i < psArr.size() - 1; i++)
    {
        if (maxReadedSize.PeakWorkingSetSize < psArr[i].PeakWorkingSetSize)
        {
            maxReadedSize = psArr[i];
            maxID = psID[i];
        }
    }
    printf("\nProcess ID: %u\n", maxID);
    printf("\PeakWorkingSetSize: 0x%08X\n", maxReadedSize.PeakPagefileUsage);
    cout << "Name of process: " << ProcessIdToName(maxID) << endl;
}

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

void CreateMsAccessProcess()
{
    STARTUPINFO si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));
    CreateProcess(L"C:\\Program Files (x86)\\Microsoft Office\\Office15\\MSACCESS.EXE", NULL, NULL, NULL, TRUE
        , REALTIME_PRIORITY_CLASS, NULL, NULL, &si, &pi);
}

int ShowInfoAboutThreads(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService, int activeProcessId
    , int numberOfThreads)
{
    IEnumWbemClassObject* pEnumerator = NULL;

    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM WIN32_THREAD WHERE ProcessHandle=")
        , WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* clsObj = NULL;
    int numElems;
        if (!FAILED(hRes))
        {
            while (numberOfThreads != 0)
            {
                if ((hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems)) != WBEM_S_FALSE)
                {
                    VARIANT vRet;
                    VariantInit(&vRet);
                    if (SUCCEEDED(clsObj->Get(L"ProcessHandle", 0, &vRet, NULL, NULL)))
                    {
                        std::wcout << L"Id that created process: " << vRet.uintVal << endl;
                        VariantClear(&vRet);
                    }
                    if (SUCCEEDED(clsObj->Get(L"Priority", 0, &vRet, NULL, NULL)))
                    {
                        std::wcout << L"Dynamics priority: " << vRet.uintVal << endl;
                        VariantClear(&vRet);
                    }
                    if (SUCCEEDED(clsObj->Get(L"PriorityBase", 0, &vRet, NULL, NULL)))
                    {
                        std::wcout << L"Base priority: " << vRet.uintVal << endl;
                        VariantClear(&vRet);
                    }
                    if (SUCCEEDED(clsObj->Get(L"ElapsedTime", 0, &vRet, NULL, NULL)))
                    {
                        std::wcout << L"Time spent: " << vRet.uintVal << endl;
                        VariantClear(&vRet);
                    }
                    if (SUCCEEDED(clsObj->Get(L"ThreadState", 0, &vRet, NULL, NULL)))
                    {
                        std::wcout << L"State: " << vRet.uintVal << endl;
                        VariantClear(&vRet);
                    }

                    numberOfThreads--;
                    cout << endl;
                    pEnumerator->Release();
                }
            }

        clsObj->Release();
    }
}

int ShowInfoAboutRunningProcess(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
    CreateMsAccessProcess();

    cout << "Third: " << endl << endl;

    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_Process WHERE Name = 'MSACCESS.EXE'")
        , WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* clsObj = NULL;
    int numElems;
    int activeProcessId;
    int numberOfThreads;
    if ((hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems)) != WBEM_S_FALSE)
    {
        if (!FAILED(hRes)) 
        {
            VARIANT vRet;
            VariantInit(&vRet);
            if (SUCCEEDED(clsObj->Get(L"ExecutablePath", 0, &vRet, NULL, NULL)))
            {
                std::wcout << L"Path: " << vRet.bstrVal << endl;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"Name", 0, &vRet, NULL, NULL)))
            {
                std::wcout << L"Name: " << vRet.bstrVal << endl;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"Priority", 0, &vRet, NULL, NULL)))
            {
                std::wcout << L"Priority: " << vRet.uintVal << endl;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"ProcessId", 0, &vRet, NULL, NULL)))
            {
                activeProcessId = vRet.uintVal;
                std::wcout << L"Id: " << activeProcessId << endl;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"ThreadCount", 0, &vRet, NULL, NULL)))
            {
                numberOfThreads = vRet.uintVal;
                std::wcout << L"Thread count: " << numberOfThreads << endl;
                VariantClear(&vRet);
            }
        }

        clsObj->Release();
    }
    pEnumerator->Release();

    ShowInfoAboutThreads(hRes, pLocator, pService, activeProcessId, numberOfThreads);

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
    ShowInfoAboutRunningProcess(hRes, pLocator, pService);
    GetInfoAboutProcByReadingSize();

    // Fifth
    pService->Release();
    pLocator->Release();

    system("pause");
    return 0;
}

