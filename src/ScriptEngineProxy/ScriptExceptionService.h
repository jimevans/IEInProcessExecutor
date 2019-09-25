// ScriptExceptionService.h : Declaration of the ScriptExceptionService

#pragma once

#include "ScriptEngineProxy_i.h"

#include <string>

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// ScriptExceptionService

class ATL_NO_VTABLE ScriptExceptionService :
	  public CComObjectRootEx<CComSingleThreadModel>,
	  public CComCoClass<ScriptExceptionService, &CLSID_ScriptExceptionService>,
    public IServiceProvider,
    public ICanHandleException,
    public IScriptExceptionService {
public:
	ScriptExceptionService() {
	}

  DECLARE_NO_REGISTRY()
  DECLARE_NOT_AGGREGATABLE(ScriptExceptionService)

  BEGIN_COM_MAP(ScriptExceptionService)
	  COM_INTERFACE_ENTRY(IScriptExceptionService)
    COM_INTERFACE_ENTRY(IServiceProvider)
    COM_INTERFACE_ENTRY(ICanHandleException)
  END_COM_MAP()

  // IServiceProvider
  STDMETHOD(QueryService)(REFGUID guidService,
                          REFIID riid,
                          void** ppvObject);

  // ICanHandleException
  STDMETHOD(CanHandleException)(EXCEPINFO* exception_info_pointer,
                                VARIANT* variant_value);

  std::wstring message() { return this->message_; }
  std::wstring source() { return this->source_; }

private:
  std::wstring message_;
  std::wstring source_;

};

OBJECT_ENTRY_AUTO(__uuidof(ScriptExceptionService), ScriptExceptionService)
