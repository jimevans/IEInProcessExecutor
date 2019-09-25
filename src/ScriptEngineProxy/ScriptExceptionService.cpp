// ScriptExceptionService.cpp : Implementation of ScriptExceptionService

#include "stdafx.h"
#include "ScriptExceptionService.h"


// ScriptExceptionService

STDMETHODIMP ScriptExceptionService::QueryService(REFGUID guidService,
                                                  REFIID riid,
                                                  void** ppvObject) {
  return S_OK;
}

STDMETHODIMP ScriptExceptionService::CanHandleException(EXCEPINFO* exception_info_pointer,
                                                        VARIANT* variant_value) {
  this->message_ = exception_info_pointer->bstrDescription;
  this->source_ = exception_info_pointer->bstrSource;
  return S_OK;
}
