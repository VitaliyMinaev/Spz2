#define _WIN32_DCOM

#include <iostream>
#include <windows.h>
#include <wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")

using namespace std;

int main() 
{
	/*PART 1*/

	HRESULT hres;
	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{
		cout << "Failed to initialize COM library. Error code = 0x"
			<< hex << hres << endl;
		return hres;
	}
	hres = CoInitializeSecurity(
		NULL,                        // Security descriptor    
		-1,                          // COM negotiates authentication service
		NULL,                        // Authentication services
		NULL,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication level for proxies
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation level for proxies
		NULL,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities of the client or server
		NULL);                       // Reserved

	if (FAILED(hres))
	{
		cout << "Failed to initialize security. Error code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		return hres;                  // Program has failed.
	}


	/*PART 2*/

	IWbemLocator* pLoc = 0;

	hres = CoCreateInstance(CLSID_WbemLocator, 0,
		CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);

	if (FAILED(hres))
	{
		cout << "Failed to create IWbemLocator object. Err code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		return hres;     // Program has failed.
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
		cout << "Could not connect. Error code = 0x"
			<< hex << hres << endl;
		pLoc->Release();
		CoUninitialize();
		return hres;      // Program has failed.
	}

	cout << "Connected to WMI" << endl;


	/*PART 3*/

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
		cout << "Could not set proxy blanket. Error code = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return hres;      // Program has failed.
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