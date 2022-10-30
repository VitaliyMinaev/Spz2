#include <iostream>
#include <Windows.h>
#include <WbemCli.h>
#include <vector>
#include <Psapi.h>
#include <stdio.h>
#include <sstream>
#include <comutil.h>

#include <tchar.h>
#include <wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "kernel32.lib")
using namespace std;

HANDLE hConsole;

// To change moment
const wchar_t* oleksiyPath = L"C:\\Program Files (x86)\\Microsoft Office\\Office15\\MSACCESS.EXE";
const wchar_t* myPath = L"C:\\Program Files\\Microsoft Office\\Root\\Office16\\EXCEL.EXE";
const wchar_t* pathToMsAccess = myPath;

// Prototypes
void PrintFail(const char* text, HRESULT res);
void PrintSuccess(const char* text);

// Task 1
int PrintKeyboardInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
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
                    wcout << "Property value: " << nIdx << vtProp.bstrVal << endl;
                }

                VariantClear(&vtProp);

            }
            hr = SafeArrayUnaccessData(sfArray);
            if (FAILED(hr)) return hr;
        }

        pclsObj->Release();

        cout << endl;
    }

    return 0;
}
// Task 2
int PrintKeyboardSpecificInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_Keyboard"), WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* clsObj = NULL;
    int numElems;
    hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems);

    if (FAILED(hRes)) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive info about keyboard: " << std::hex << hRes << endl;
        return 1;
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

    pEnumerator->Release();

    cout << endl;
    return 0;
}

// Task 3
void CreateMsAccessProcess()
{
    STARTUPINFO si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));
    CreateProcess(pathToMsAccess, NULL, NULL, NULL, TRUE
        , REALTIME_PRIORITY_CLASS, NULL, NULL, &si, &pi);
}
int ShowInfoAboutThreads(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService, int processId
    , int threadsCount)
{
    IEnumWbemClassObject* pEnumerator = NULL;

    stringstream oss;
    string queryStr = "SELECT * FROM WIN32_THREAD WHERE ProcessHandle=";
    oss << processId;
    queryStr += oss.str();

    BSTR query = _com_util::ConvertStringToBSTR(queryStr.c_str());
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), query, WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* clsObj = NULL;
    int numElems;
    if (!FAILED(hRes))
    {
        while (threadsCount != 0)
        {
            if (FAILED((hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems))) == false)
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

                cout << endl;
            }

            threadsCount--;

        }
        pEnumerator->Release();
        clsObj->Release();
    }
    return 0;
}
int PrintRunningProcessInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
    CreateMsAccessProcess();

    IEnumWbemClassObject* pEnumerator = NULL;
    // CHANGE !!!
    // CHANGE !!!
    // CHANGE !!!
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
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

    return 0;
}

// Task 4
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
int PrintInfoAboutProcessWithMaxReadedSize() {
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

    cout << "Process ID: " << maxID << endl;
    cout << "Peak page file usage" << maxReadedSize.PeakPagefileUsage << endl;
    cout << "Name of process: " << ProcessIdToName(maxID) << endl << endl;

    return 0;
}

// Task 5a
HRESULT KillLowPriorityNotepadProcess(IWbemServices* pService)
{
    HRESULT hr = S_OK;

    static LPCTSTR lpszMethod = _T("Terminate");
    static LPCTSTR lpszClass = _T("Win32_Process");

    IWbemClassObject* pClsInParam = NULL;
    IWbemClassObject* pClsInParamInst = NULL;

    BSTR bszClsMoniker = NULL;
    VARIANT v;

    VariantInit(&v);

    IEnumWbemClassObject* pEnum = NULL;

    hr = pService->ExecQuery(
        (BSTR)_T("WQL"),
        (BSTR)_T("SELECT * ")
        _T("FROM Win32_Process ")
        _T("WHERE Name='notepad.exe' AND Priority='4'"),
        0,
        NULL, &pEnum
    );

    IWbemClassObject* pObj = NULL;
    IWbemClassObject* pClsDef = NULL;
    if (FAILED(hr))
        goto fail;

    hr = pService->GetObject(
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

        hr = pService->ExecMethod(
            bszClsMoniker, (BSTR)lpszMethod, 0,
            NULL, pClsInParamInst, NULL, NULL
        );

        SysFreeString(bszClsMoniker);

        if (FAILED(hr)) {
            _tprintf_s(_T("ExecMethod failed, hr: %lX\n"), hr);
            goto fail;
        }
    }

    PrintSuccess("Task 5a was completed successfully.");
    return hr;

    fail:
        PrintFail("Task 5a - something goes wrong", hr);
        VariantClear(&v);
        pClsInParamInst->Release();
        pClsInParam->Release();
        pClsDef->Release();
        pObj->Release();
        pEnum->Release();
}
// Task 5b
HRESULT KillTotalCommanderChildProcess(IWbemServices* pService)
{
    HRESULT hr = S_OK;

    IEnumWbemClassObject* pEnum = NULL;
    IWbemClassObject* pClsInParam = NULL;

    IWbemClassObject* pClsInParamInst = NULL;

    static LPCTSTR lpszMethod = _T("Terminate");
    static LPCTSTR lpszClass = _T("Win32_Process");

    BSTR bszWQLQueryChild = NULL;

    VARIANT v;

    VariantInit(&v);

    IWbemClassObject* pClsDef = NULL;

    hr = pService->GetObject(
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

    hr = pService->ExecQuery(
        (BSTR)_T("WQL"),
        (BSTR)_T("SELECT * ")
        _T("FROM Win32_Process ")
        _T("WHERE Name='totalcmd.exe' OR Name='totalcmd64.exe'"),
        0,
        NULL, &pEnum
    );

    IWbemClassObject* pObj = NULL;

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

        hr = pService->ExecQuery(
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

            hr = pService->ExecMethod(
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

    PrintSuccess("Task 5b was completed successfully.");
    return hr;

fail:
    PrintFail("Task 5a - something goes wrong", hr);
    SysFreeString(bszWQLQueryChild);
    VariantClear(&v);
    pClsInParamInst->Release();
    pClsInParam->Release();
    pClsDef->Release();
    pObj->Release();
    pEnum->Release();
}

int main()
{
    // Colorized console
    hConsole = GetStdHandle(STD_OUTPUT_HANDLE);

    HRESULT hRes = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hRes)) {
        cout << "Unable to launch COM: 0x" << std::hex << hRes << endl;
        return 1;
    }
    else {
        PrintSuccess("Con library has been successfully initialized");
    }

    if ((FAILED(hRes = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_CONNECT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, 0))))
    {
        cout << "Unable to initialize security: 0x" << std::hex << hRes << endl;
        return 1;
    }
    else {
        PrintSuccess("Security layers has been successfully initialized");
    }

    IWbemLocator* pLocator = NULL;
    if (FAILED(hRes = CoCreateInstance(CLSID_WbemLocator, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pLocator)))) {
        cout << "Unable to create a WbemLocator: " << std::hex << hRes << endl;
        return 1;
    }
    else {
        PrintSuccess("WbemLocator has been successfully created");
    }

    IWbemServices* pService = NULL;
    if (FAILED(hRes = pLocator->ConnectServer(BSTR(L"root\\CIMV2"), NULL, NULL, NULL, WBEM_FLAG_CONNECT_USE_MAX_WAIT, NULL, NULL, &pService))) {
        pLocator->Release();
        cout << "Unable to connect to \"CIMV2\": " << std::hex << hRes << endl;
        return 1;
    }
    else {
        PrintSuccess("Connection to server has been successfully created");
    }

    // The First task
    cout << endl << "Full info about keyboard (Win32_KeyBoard): " << endl << endl;

    PrintKeyboardInfo(hRes, pLocator, pService);

    // The Second task
    cout << endl << "Description and number of funcrion keys: " << endl << endl;

    PrintKeyboardSpecificInfo(hRes, pLocator, pService);

    // The Third task
    cout << endl << "Info about running process: " << endl << endl;

    PrintRunningProcessInfo(hRes, pLocator, pService);

    // The Fourth task
    cout << endl << "Info about process, which has the biggest readed size: " << endl << endl;

    PrintInfoAboutProcessWithMaxReadedSize();

    // Task 5a
    cout << endl << "Stop lop priotiry notepad process: " << endl << endl;

    KillLowPriorityNotepadProcess(pService);
    PrintSuccess("Process were stopped!");

    // Task 5b
    cout << endl << "Stop total commander child process" << endl << endl;

    KillTotalCommanderChildProcess(pService);
    PrintSuccess("Total commander child process were stopped!");

    // Release
    pService->Release();
    pLocator->Release();

    system("pause");
    return 0;
}

void PrintFail(const char* text, HRESULT res) {
    SetConsoleTextAttribute(hConsole, 12);
    cout << text << std::hex << res << endl;
    SetConsoleTextAttribute(hConsole, 7);
}
void PrintSuccess(const char* text) {
    SetConsoleTextAttribute(hConsole, 10);
    cout << text << endl;
    SetConsoleTextAttribute(hConsole, 7);
}