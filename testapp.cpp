#define _WIN32_DCOM

#include <iostream>
#include <WbemIdl.h>
#pragma comment(lib, "wbemuuid.lib")

using namespace std;

int main() 
{
	HRESULT hr;
	hr = CoInitializeEx(0, COINITBASE_MULTITHREADED);
	if (FAILED(hr) == true) {
		cout << "Failed to initialize con library. Error code 0x" << hex << hr << endl;
		return 1;
	}
	hr = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		EOAC_NONE,
		NULL);

	if (FAILED(hr) == true) {
		cout << "Failed to initialize con library. Error code 0x" << hex << hr << endl;
		return 1;
	}

	system("pause");
	return 0;
}