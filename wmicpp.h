#pragma once

#include <comdef.h>
#include <wbemidl.h>

#include <functional>
#include <string_view>
#include <system_error>

namespace wmi
{
	template<typename T>
	using com_ptr_t = _com_ptr_t<_com_IIID<T, &__uuidof(T)>>;


	struct wbem_category : public std::error_category
	{
		com_ptr_t<IWbemStatusCodeText> m_codeText;

		wbem_category() : m_codeText(CLSID_WbemStatusCodeText)
		{}

		const char * name() const noexcept override
		{
			return "wbem error";
		}

		std::string message(int ev) const override
		{
			std::string msg;

			bstr_t text;

			if (auto hr = m_codeText->GetErrorCodeText(ev, 0, 0, text.GetAddress()); SUCCEEDED(hr))
			{
				if (int cch = ::WideCharToMultiByte(CP_THREAD_ACP, 0, text, text.length(), nullptr, 0, nullptr, nullptr); cch > 0)
				{
					msg.resize(cch);

					::WideCharToMultiByte(CP_THREAD_ACP, 0, text, text.length(), msg.data(), cch, nullptr, nullptr);
				}
			}

			return msg;
		}
	};


	class ClassObjects_Iteratir
	{
		com_ptr_t<IEnumWbemClassObject> m_objects;
		com_ptr_t<IWbemClassObject> m_object;

	public:
		ClassObjects_Iteratir(IEnumWbemClassObject * objects) : m_objects(objects)
		{
			operator++();
		}

		IWbemClassObject * operator*()
		{
			return m_object;
		}

		ClassObjects_Iteratir & operator++()
		{
			ULONG count{};

			if (auto hr = m_objects->Next(WBEM_INFINITE, 1, &m_object, &count); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			if (m_object && count == 0)
			{
				m_object.Release();
			}

			return *this;
		}

		bool operator!=(IWbemClassObject * object)
		{
			return m_object != object;
		}
	};


	class ClassObjects
	{
		com_ptr_t<IWbemServices> m_services;
		bstr_t m_className;

	public:
		ClassObjects(IWbemServices * services, const wchar_t * className) : m_services(services), m_className(className)
		{}

		ClassObjects_Iteratir begin()
		{
			com_ptr_t<IEnumWbemClassObject> enumerator;

			if (auto hr = m_services->CreateInstanceEnum(m_className, WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_USE_AMENDED_QUALIFIERS/*WBEM_FLAG_RETURN_IMMEDIATELY*/, nullptr, &enumerator); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			return ClassObjects_Iteratir(enumerator);
		}

		IWbemClassObject * end()
		{
			return nullptr;
		}
	};


	class WbemObjectSink : public IWbemObjectSink
	{
		unsigned long m_count;
		HANDLE m_event;
		std::function<void(const wchar_t *, int, IWbemClassObject *)> m_callback;

	public:
		WbemObjectSink(std::function<void(const wchar_t *, int, IWbemClassObject *)> callback) : m_count(0), m_event(::CreateEvent(nullptr, true, false, nullptr)), m_callback(callback)
		{
			if (!m_event)
			{
				throw std::system_error(::GetLastError(), std::system_category(), __FUNCTION__);
			}
		}

		~WbemObjectSink()
		{
			::CloseHandle(m_event);
		}

		operator HANDLE()
		{
			return m_event;
		}

		STDMETHODIMP_(ULONG) AddRef(void) override
		{
			return ::InterlockedIncrement(&m_count);
		}

		STDMETHODIMP_(ULONG) Release(void) override
		{
			auto count = ::InterlockedDecrement(&m_count);

			if (count == 0)
			{
				delete this;
			}

			return count;
		}

		STDMETHODIMP QueryInterface(REFIID riid, _COM_Outptr_ void ** ppvObject) override
		{
			if (riid == __uuidof(IUnknown) || riid == __uuidof(IWbemObjectSink))
			{
				AddRef();

				*ppvObject = (IWbemObjectSink *) this;

				return S_OK;
			}
			else
			{
				*ppvObject = nullptr;

				return E_NOINTERFACE;
			}
		}

		STDMETHODIMP Indicate(long lObjectCount, __RPC__in_ecount_full(lObjectCount) IWbemClassObject ** apObjArray) override
		{
			for (long index = 0; index < lObjectCount; ++index)
			{
				m_callback(L"Callbacl", index, apObjArray[index]);
			}

			return WBEM_S_NO_ERROR;
		}

		STDMETHODIMP SetStatus(long lFlags, HRESULT hResult, __RPC__in_opt BSTR strParam, __RPC__in_opt IWbemClassObject * pObjParam) override
		{
			::SetEvent(m_event);

			return WBEM_S_NO_ERROR;
		}
	};


	struct ComInitialize
	{
		ComInitialize()
		{
			if (auto hr = ::CoInitializeEx(nullptr, COINIT_MULTITHREADED); FAILED(hr))
			{
				throw std::system_error(hr, std::system_category(), __FUNCTION__);
			}
		}

		~ComInitialize()
		{
			::CoUninitialize();
		}
	};


	struct ComInitializeSecurity : ComInitialize
	{
		ComInitializeSecurity()
		{
			if (auto hr = ::CoInitializeSecurity(nullptr, -1, nullptr, nullptr, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE, nullptr); FAILED(hr))
			{
				throw std::system_error(hr, std::system_category(), __FUNCTION__);
			}
		}
	};


	class Services
	{
		com_ptr_t<IWbemServices> m_services;

	public:
		Services(const CLSID & rclsid, const wchar_t * resource)
		{
			thread_local static ComInitializeSecurity com;

			com_ptr_t<IWbemLocator> locator;

			if (auto hr = locator.CreateInstance(rclsid); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			if (auto hr = locator->ConnectServer(bstr_t(resource), nullptr, nullptr, nullptr, 0, nullptr, nullptr, &m_services); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			if (auto hr = ::CoSetProxyBlanket(m_services, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, nullptr, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE); FAILED(hr))
			{
				throw std::system_error(hr, std::system_category(), __FUNCTION__);
			}
		}

		operator IWbemServices *()
		{
			return m_services;
		}

		com_ptr_t<IWbemClassObject> GetClassObject(const wchar_t * className)
		{
			com_ptr_t<IWbemClassObject> object;

			if (auto hr = m_services->GetObject(bstr_t(className), 0, nullptr, &object, nullptr); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			return object;
		}

		ClassObjects GetClassObjects(const wchar_t * className)
		{
			return ClassObjects(m_services, className);
		}

		void QueryObjects(const wchar_t * wql, std::function<void(const wchar_t *, int, IWbemClassObject *)> callback)
		{
			auto sink = new (std::nothrow) WbemObjectSink(callback);

			com_ptr_t<IWbemObjectSink> handler(sink);

			if (auto hr = m_services->ExecQueryAsync(bstr_t("WQL"), bstr_t(wql), WBEM_FLAG_USE_AMENDED_QUALIFIERS, nullptr, handler); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			::WaitForSingleObject(*sink, INFINITE);
		}
	};


	template<typename T>
	constexpr VARTYPE GetVarType() = delete;

	template<>
	constexpr VARTYPE GetVarType<BSTR>()
	{
		return VT_BSTR;
	}

	template<>
	constexpr VARTYPE GetVarType<IUnknown *>()
	{
		return VT_UNKNOWN;
	}

	template<typename T>
	ULONG SafeArrayAccessDataFunc(SAFEARRAY * psa, T ** ppv)
	{
		*ppv = nullptr;

		VARTYPE vt{};

		if (auto hr = ::SafeArrayGetVartype(psa, &vt); FAILED(hr))
		{
			throw std::system_error(hr, std::system_category(), __FUNCTION__);
		}

		if (vt != GetVarType<T>())
		{
			throw std::system_error(E_UNEXPECTED, std::system_category(), __FUNCTION__);
		}

		LONG lbound{};

		if (auto hr = ::SafeArrayGetLBound(psa, 1, &lbound); FAILED(hr))
		{
			throw std::system_error(hr, std::system_category(), __FUNCTION__);
		}

		LONG ubound{};

		if (auto hr = ::SafeArrayGetUBound(psa, 1, &ubound); FAILED(hr))
		{
			throw std::system_error(hr, std::system_category(), __FUNCTION__);
		}

		if (lbound > ubound)
		{
			throw std::system_error(E_UNEXPECTED, std::system_category(), __FUNCTION__);
		}

		if (lbound < ubound)
		{
			if (auto hr = ::SafeArrayAccessData(psa, (void **) ppv); FAILED(hr))
			{
				throw std::system_error(hr, std::system_category(), __FUNCTION__);
			}

			return ubound - lbound;
		}

		return 0;
	}


	class PropertyNames
	{
		SAFEARRAY * m_psa;
		BSTR * m_names;
		ULONG m_count;

	public:
		PropertyNames(IWbemClassObject * object) : m_psa(nullptr), m_names(nullptr), m_count(0)
		{
			if (auto hr = object->GetNames(nullptr, WBEM_FLAG_ALWAYS, nullptr, &m_psa); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			try
			{
				m_count = SafeArrayAccessDataFunc(m_psa, &m_names);
			}
			catch (...)
			{
				::SafeArrayDestroy(m_psa);
				throw;
			}
		}

		~PropertyNames() noexcept
		{
			::SafeArrayUnaccessData(m_psa);
			::SafeArrayDestroy(m_psa);
		}

		BSTR * begin()
		{
			return m_names;
		}

		BSTR * end()
		{
			return m_names + m_count;
		}
	};


	class Method
	{
		com_ptr_t<IWbemServices> m_service;
		com_ptr_t<IWbemClassObject> m_object;
		com_ptr_t<IWbemClassObject> m_method;

		bstr_t m_className;
		bstr_t m_methodName;

	public:
		Method(IWbemServices * service, const wchar_t * clazz, const wchar_t * method) : m_service(service), m_className(clazz), m_methodName(method)
		{
			if (auto hr = m_service->GetObject(m_className, 0, nullptr, &m_object, nullptr); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			if (auto hr = m_object->GetMethod(m_methodName, 0, &m_method, nullptr); FAILED(hr))
			{
				throw std::system_error(hr, wmi::wbem_category(), __FUNCTION__);
			}
		}

		void SetProperty(const wchar_t * name, variant_t && value)
		{
			if (auto hr = m_method->Put(name, 0, &value, 0); FAILED(hr))
			{
				throw std::system_error(hr, wmi::wbem_category(), __FUNCTION__);
			}
		}

		com_ptr_t<IWbemClassObject> Exec()
		{
			com_ptr_t<IWbemClassObject> outParams;

			if (auto hr = m_service->ExecMethod(m_className, m_methodName, 0, nullptr, m_method, &outParams, nullptr); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			return outParams;
		}

		void Exec(std::function<void(const wchar_t *, int, IWbemClassObject *)> callback)
		{
			auto sink = new (std::nothrow) WbemObjectSink(callback);

			com_ptr_t<IWbemObjectSink> handler(sink);

			if (auto hr = m_service->ExecMethodAsync(m_className, m_methodName, 0, nullptr, m_method, handler); FAILED(hr))
			{
				throw std::system_error(hr, wbem_category(), __FUNCTION__);
			}

			::WaitForSingleObject(*sink, INFINITE);
		}
	};


	template<class T>
	T SpawnInstance(T && object)
	{
		com_ptr_t<IWbemClassObject> newInstance;

		if (auto hr = object->SpawnInstance(0, &newInstance); FAILED(hr))
		{
			throw std::system_error(hr, wmi::wbem_category(), __FUNCTION__);
		}

		return T(newInstance);
	}
}

#pragma comment(lib, "wbemuuid.lib")
