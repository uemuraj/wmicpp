#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <system_error>

#include "wmicpp.h"

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ PWSTR pCmdLine, _In_ int nCmdShow)
{
	try
	{
		wmicpp::Initialized wc;
		wmicpp::WbemNamespace services(L"ROOT\\CIMV2");

		for (auto obj : services.GetObjects(L"Win32_BaseBoard"))
		{
			variant_t value;
			obj->Get(L"Manufacturer", 0, &value, nullptr, nullptr);
			::OutputDebugStringW(value.bstrVal);
			::OutputDebugStringW(L" ");
			obj->Get(L"Product", 0, &value, nullptr, nullptr);
			::OutputDebugStringW(value.bstrVal);
			::OutputDebugStringW(L"\r\n");
		}

		return 0;
	}
	catch (const std::exception & e)
	{
		::OutputDebugStringA(e.what());
		::OutputDebugStringA("\r\n");
		return -1;
	}
}
