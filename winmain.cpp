#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <locale>
#include <sstream>
#include <system_error>

#include "wmicpp.h"

void Dump(const wchar_t * className, int index, IWbemClassObject * obj)
{
	std::wostringstream out;

	wmicpp::WebmPropertyNames names(obj);

	for (auto & name : names)
	{
		variant_t value;

		if (auto hr = obj->Get(name, 0, &value, nullptr, nullptr); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}

		out.str(L"");

		out << className << L'[' << index << L"].";

		if (value.vt == VT_NULL)
		{
			out << name << L" = null" << std::endl;
		}
		else if (value.vt == VT_BSTR)
		{
			out << name << L" = " << value.bstrVal << std::endl;
		}
		else
		{
			variant_t changed;

			if (auto hr = ::VariantChangeTypeEx(&changed, &value, LOCALE_SYSTEM_DEFAULT, VARIANT_ALPHABOOL, VT_BSTR); SUCCEEDED(hr))
			{
				out << name << L" = " << changed.bstrVal << std::endl;
			}
			else
			{
				out << name << L'(' << value.vt << L')' << std::endl;
			}
		}

		::OutputDebugStringW(out.str().c_str());
	}
}

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ PWSTR pCmdLine, _In_ int nCmdShow)
{
	try
	{
		std::locale::global(std::locale(""));

		wmicpp::Initialized wc;
		wmicpp::WbemNamespace cimv2(L"ROOT\\CIMV2");

		for (int index = 0; auto obj : cimv2.GetObjects(L"Win32_OperatingSystem"))
		{
			Dump(L"Win32_OperatingSystem", index++, obj);
		}

		for (int index = 0; auto obj : cimv2.GetObjects(L"Win32_BaseBoard"))
		{
			Dump(L"Win32_BaseBoard", index++, obj);
		}

		for (int index = 0; auto obj : cimv2.GetObjects(L"Win32_BIOS"))
		{
			Dump(L"Win32_BIOS", index++, obj);
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
