#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <system_error>

#include "wmicpp.h"

namespace wmicpp
{
	Initialized::Initialized()
	{
		if (auto hr = ::CoInitializeEx(0, COINIT_MULTITHREADED); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}

		if (auto hr = ::CoInitializeSecurity(nullptr, -1, nullptr, nullptr, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE, nullptr); FAILED(hr))
		{
			::CoUninitialize();
			RAISE_SYSTEM_ERROR(hr);
		}
	}

	Initialized::~Initialized()
	{
		::CoUninitialize();
	};

	WbemNamespace::WbemNamespace(const wchar_t * resource)
	{
		com_ptr_t<IWbemLocator> locator;

		if (auto hr = locator.CreateInstance(CLSID_WbemLocator); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}

		if (auto hr = locator->ConnectServer(bstr_t(resource), nullptr, nullptr, nullptr, 0, nullptr, nullptr, &m_services); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}

		if (auto hr = ::CoSetProxyBlanket(m_services, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, nullptr, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}
	}

	WbemObjects WbemNamespace::GetObjects(const wchar_t * className)
	{
		com_ptr_t<IEnumWbemClassObject> enumerator;

		if (auto hr = m_services->CreateInstanceEnum(bstr_t("Win32_BaseBoard"), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &enumerator); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}

		return WbemObjects(enumerator);
	}

	WbemObject_Iteratir::WbemObject_Iteratir(IEnumWbemClassObject * p) : m_objects(p)
	{
		operator++();
	}

	IWbemClassObject * WbemObject_Iteratir::operator*()
	{
		return m_object;
	}

	WbemObject_Iteratir & WbemObject_Iteratir::operator++()
	{
		ULONG count{};

		if (auto hr = m_objects->Next(WBEM_INFINITE, 1, &m_object, &count); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}

		if (count == 0)
		{
			m_objects.Release();
		}

		return *this;
	}
}
