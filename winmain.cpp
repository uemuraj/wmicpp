#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <locale>
#include <sstream>
#include <system_error>

#include "wmicpp.h"

// https://github.com/jim-dale/MountISO.Wmi/blob/main/src/DiskImageManager.h
// TODO: ISO ファイルのマウント／アンマウントができるか？

void Dump(const wchar_t * className, int index, IWbemClassObject * obj)
{
	std::wostringstream out;

	wmi::PropertyNames names(obj);

	for (auto & name : names)
	{
		variant_t value;

		if (auto hr = obj->Get(name, 0, &value, nullptr, nullptr); FAILED(hr))
		{
			throw std::system_error(hr, wmi::wbem_category(), __FUNCTION__);
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
		else if (value.vt == VT_UNKNOWN)
		{
			wmi::com_ptr_t<IWbemClassObject> object(value.punkVal);

			out << name << std::endl;

			::OutputDebugStringW(out.str().c_str());

			out.str(L"");

			auto valueName = std::wstring(L"  ") + name;
			Dump(valueName.c_str(), 0, object);
		}
		else if (value.vt == (VT_ARRAY | VT_UNKNOWN))
		{
			IUnknown ** values{};
			auto count = wmi::SafeArrayAccessDataFunc(value.parray, &values);

			out << name << L'[' << count << L']' << std::endl;

			::OutputDebugStringW(out.str().c_str());

			out.str(L"");

			auto valueName = std::wstring(L"  ") + name;

			for (decltype(count) i = 0; i < count; ++i)
			{
				wmi::com_ptr_t<IWbemClassObject> object(values[i]);

				Dump(valueName.c_str(), i, object);
			}

			::SafeArrayUnaccessData(value.parray);
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

		//wmi::Services cimv2(CLSID_WbemLocator, L"ROOT\\CIMV2");

		//int index = 0;

		//for (auto obj : cimv2.GetClassObjects(L"Win32_OperatingSystem"))
		//{
		//	Dump(L"Win32_OperatingSystem", index++, obj);
		//}

		//index = 0;

		//for (auto obj : cimv2.GetClassObjects(L"Win32_BaseBoard"))
		//{
		//	Dump(L"Win32_BaseBoard", index++, obj);
		//}

		//index = 0;

		//for (auto obj : cimv2.GetClassObjects(L"Win32_BIOS"))
		//{
		//	Dump(L"Win32_BIOS", index++, obj);
		//}

		//index = 0;

		//for (auto obj : cimv2.GetClassObjects(L"Win32_Product"))
		//{
		//	Dump(L"Win32_Product", index++, obj);
		//}

		wmi::Services serverManager(CLSID_WbemLocator, LR"(ROOT/Microsoft/Windows/ServerManager)");
#if 0
		wmi::Method method(serverManager, L"MSFT_ServerManagerTasks", L"GetServerFeature");

		auto result = method.Exec();

		Dump(L"GetServerFeature", 0, result);
#else
		wmi::Method method(serverManager, L"MSFT_ServerManagerDeploymentTasks", L"GetServerComponentsAsync");

		auto result = method.Exec();

		Dump(L"GetServerComponentsAsync", 0, result);
#endif
		return 0;
	}
	catch (const std::exception & e)
	{
		::OutputDebugStringA(e.what());
		::OutputDebugStringA("\r\n");
		return -1;
	}
}
