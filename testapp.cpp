#define _WIN32_DCOM

#include <iostream>
#include <WbemIdl.h>
#include <Windows.h>
#pragma comment(lib, "wbemuuid.lib")

using namespace std;

int main()
{
	/* Colorize the console */
	HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);

	/*PART 1*/

	HRESULT hres;
	hres = CoInitializeEx(0, COINITBASE_MULTITHREADED);
	if (FAILED(hres) == true) {
		SetConsoleTextAttribute(hConsole, 12);
		cout << "Failed to initialize con library. Error code 0x" << hex << hres << endl;
		SetConsoleTextAttribute(hConsole, 7);
		return 1;
	}

	hres = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		EOAC_NONE,
		NULL);

	if (FAILED(hres) == true) {
		SetConsoleTextAttribute(hConsole, 12);
		cout << "Failed to initialize con library. Error code 0x" << hex << hres << endl;
		SetConsoleTextAttribute(hConsole, 7);
		return 1;
	}
	else {
		SetConsoleTextAttribute(hConsole, 10);
		cout << "Con library has been successfully initialized" << endl;
		SetConsoleTextAttribute(hConsole, 7);
	}


	/*PART 2*/

	IWbemLocator* pLoc = 0;

	hres = CoCreateInstance(CLSID_WbemLocator, 0,
		CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);

	if (FAILED(hres))
	{
		SetConsoleTextAttribute(hConsole, 12);
		cout << "Failed to create IWbemLocator object. Err code = 0x"
			<< hex << hres << endl;
		SetConsoleTextAttribute(hConsole, 7);
		CoUninitialize();
		return hres;     // Program has failed.
	}
	else {
		SetConsoleTextAttribute(hConsole, 10);
		cout << "IWbemLocator object has been successfully created" << endl;
		SetConsoleTextAttribute(hConsole, 7);
	}

	IWbemServices* pSvc = 0;

	// Connect to the root\default namespace with the current user.
	hres = pLoc->ConnectServer(
		BSTR(L"ROOT\\DEFAULT"),  //namespace
		NULL,       // User name 
		NULL,       // User password
		0,         // Locale 
		NULL,     // Security flags
		0,         // Authority 
		0,        // Context object 
		&pSvc);   // IWbemServices proxy


	if (FAILED(hres))
	{
		SetConsoleTextAttribute(hConsole, 12);
		cout << "Could not connect. Error code = 0x"
			<< hex << hres << endl;
		SetConsoleTextAttribute(hConsole, 7);
		pLoc->Release();
		CoUninitialize();
		return hres;      // Program has failed.
	}
	SetConsoleTextAttribute(hConsole, 10);
	cout << "Connected to WMI" << endl;
	SetConsoleTextAttribute(hConsole, 7);


	/*PART 3*/

	/*
	* IWbemServices* pSvc = 0;
	* IWbemLocator* pLoc = 0;
	*/

	// Set the proxy so that impersonation of the client occurs.
	hres = CoSetProxyBlanket(pSvc,
		RPC_C_AUTHN_WINNT,
		RPC_C_AUTHZ_NONE,
		NULL,
		RPC_C_AUTHN_LEVEL_CALL,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		EOAC_NONE
	);

	if (FAILED(hres))
	{
		SetConsoleTextAttribute(hConsole, 12);
		cout << "Could not set proxy blanket. Error code = 0x"
			<< hex << hres << endl;
		SetConsoleTextAttribute(hConsole, 7);
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return hres;      // Program has failed.
	}
	else {
		SetConsoleTextAttribute(hConsole, 10);
		cout << "Proxy blanket has been successfully created" << endl;
		SetConsoleTextAttribute(hConsole, 7);
	}

	/*PART 4*/

	// our realization


	/*PART 5*/

	pSvc->Release();
	pLoc->Release();
	CoUninitialize();

	system("pause");
	return 0;
}