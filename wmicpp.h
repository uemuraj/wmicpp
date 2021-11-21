#pragma once

#include <string>
#include <comdef.h>
#include <wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")

template<typename T> using com_ptr_t = _com_ptr_t<_com_IIID<T, &__uuidof(T)>>;

#define RAISE_SYSTEM_ERROR(err)	throw std::system_error(err, std::system_category(),  __FILE__ "(" _CRT_STRINGIZE(__LINE__) ")");

namespace wmicpp
{
	struct Initialized
	{
		Initialized();
		~Initialized();
	};

	class WbemObject_Iteratir
	{
		com_ptr_t<IEnumWbemClassObject> m_objects;
		com_ptr_t<IWbemClassObject> m_object;
	public:
		WbemObject_Iteratir(IEnumWbemClassObject * p);
		~WbemObject_Iteratir() = default;

		IWbemClassObject * operator*();

		WbemObject_Iteratir & operator++();

		bool operator!=(void *)
		{
			return m_object;
		}
	};

	class WbemObjects
	{
		com_ptr_t<IEnumWbemClassObject> m_objects;
	public:
		WbemObjects(IEnumWbemClassObject * p) : m_objects(p) {}
		~WbemObjects() = default;

		WbemObject_Iteratir begin()
		{
			m_objects->Reset();
			return WbemObject_Iteratir(m_objects);
		}

		void * end()
		{
			return nullptr;
		}
	};

	class WbemNamespace
	{
		com_ptr_t<IWbemServices> m_services;
	public:
		WbemNamespace(const wchar_t * resource);
		~WbemNamespace() = default;

		WbemObjects GetObjects(const wchar_t * className);
	};
}
