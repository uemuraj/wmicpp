#define WIN32_LEAN_AND_MEAN
#include <windows.h>

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

		if (auto hr = m_services->CreateInstanceEnum(bstr_t(className), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, nullptr, &enumerator); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}
		else
		{
			return WbemObjects(enumerator);
		}
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

	WebmPropertyNames::WebmPropertyNames(IWbemClassObject * p) : m_psa(nullptr), m_names(nullptr), m_count(0)
	{
		if (auto hr = p->GetNames(nullptr, WBEM_FLAG_ALWAYS, nullptr, &m_psa); FAILED(hr))
		{
			RAISE_SYSTEM_ERROR(hr);
		}

		try
		{
			VARTYPE vt{};

			if (auto hr = ::SafeArrayGetVartype(m_psa, &vt); FAILED(hr))
			{
				RAISE_SYSTEM_ERROR(hr);
			}

			if (vt != VT_BSTR)
			{
				RAISE_SYSTEM_ERROR(E_UNEXPECTED);
			}

			LONG lbound{};

			if (auto hr = ::SafeArrayGetLBound(m_psa, 1, &lbound); FAILED(hr))
			{
				RAISE_SYSTEM_ERROR(hr);
			}

			LONG ubound{};

			if (auto hr = ::SafeArrayGetUBound(m_psa, 1, &ubound); FAILED(hr))
			{
				RAISE_SYSTEM_ERROR(hr);
			}

			if (lbound > ubound)
			{
				RAISE_SYSTEM_ERROR(E_UNEXPECTED);
			}

			m_count = ubound - lbound;

			if (auto hr = ::SafeArrayAccessData(m_psa, (void **) &m_names); FAILED(hr))
			{
				RAISE_SYSTEM_ERROR(hr);
			}
		}
		catch (...)
		{
			::SafeArrayDestroy(m_psa);

			throw;
		}
	}

	WebmPropertyNames::~WebmPropertyNames() noexcept
	{
		::SafeArrayUnaccessData(m_psa);
		::SafeArrayDestroy(m_psa);
	}
}
