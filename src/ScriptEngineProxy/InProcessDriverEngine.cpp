// InProcessDriverEngine.cpp : Implementation of InProcessDriverEngine
#include "stdafx.h"

#include "InProcessDriverEngine.h"

#include <iostream>

#include <comdef.h>
#include <ShlGuid.h>

#include "StringUtilities.h"
#include "ScriptExceptionService.h"

#define ANONYMOUS_FUNCTION_START L"(function() { "
#define ANONYMOUS_FUNCTION_END L" })();"

struct WaitThreadContext {
	HWND window_handle;
};

// InProcessDriverEngine

InProcessDriverEngine::InProcessDriverEngine() {
    this->is_navigating_ = false;
    this->current_command_ = "";
    this->command_data_ = "";
    this->response_ = "";
    this->Create(HWND_MESSAGE);
}

STDMETHODIMP_(HRESULT __stdcall) InProcessDriverEngine::SetSite(IUnknown* pUnkSite) {
  HRESULT hr = S_OK;
  if (pUnkSite == nullptr) {
    return hr;
  }

  CComPtr<IWebBrowser2> browser;
  pUnkSite->QueryInterface<IWebBrowser2>(&browser);
  if (browser != nullptr) {
    CComPtr<IDispatch> document_dispatch;
    hr = browser->get_Document(&document_dispatch);
    ATLENSURE_RETURN_HR(document_dispatch.p != nullptr, E_FAIL);

    CComPtr<IHTMLDocument> document;
    hr = document_dispatch->QueryInterface<IHTMLDocument>(&document);

    CComQIPtr<IServiceProvider> spSP(document_dispatch);
    ATLENSURE_RETURN_HR(spSP != nullptr, E_NOINTERFACE);
    CComPtr<IDiagnosticsScriptEngineProvider> spEngineProvider;
    hr = spSP->QueryService(IID_IDiagnosticsScriptEngineProvider, &spEngineProvider);
    hr = spEngineProvider->CreateDiagnosticsScriptEngine(this, FALSE, 0, &this->script_engine_);

	  CComPtr<IServiceProvider> browser_service_provider;
	  hr = browser->QueryInterface<IServiceProvider>(&browser_service_provider);
	  CComPtr<IOleWindow> window;
	  hr = browser_service_provider->QueryService<IOleWindow>(SID_SShellBrowser, &window);
	  HWND handle;
	  window->GetWindow(&handle);

    this->browser_ = browser;
    this->script_host_document_ = document;
    hr = this->script_engine_->EvaluateScript(L"browser.addEventListener('consoleMessage', function(e) { external.sendMessage('consoleMessage', JSON.stringify(e)); });", L"");
	  this->DispEventAdvise(browser);
  }
  return hr;
}

STDMETHODIMP_(HRESULT __stdcall) InProcessDriverEngine::GetWindow(HWND* pHwnd) {
  *pHwnd = this->m_hWnd;
  return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) InProcessDriverEngine::OnMessage(LPCWSTR* pszData, ULONG ulDataCount) {
  std::vector<std::wstring> message;
  for (ULONG i = 0; i < ulDataCount; ++i) {
    std::wstring data(*pszData);
    message.push_back(data);
    ++pszData;
  }
  if (message.at(0) == L"snapshot") {
    this->response_ = StringUtilities::ToString(message.at(1));
  }
  if (message.at(0) == L"script") {
    this->response_ = StringUtilities::ToString(message.at(1));
  }

  if (message.at(0) == L"debug") {
    ::Sleep(1);
  }
  return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) InProcessDriverEngine::OnScriptError(IActiveScriptError* pScriptError) {
  return S_OK;
}

STDMETHODIMP_(void) InProcessDriverEngine::OnBeforeNavigate2(LPDISPATCH pDisp,
	                                                           VARIANT* pvarUrl,
	                                                           VARIANT* pvarFlags,
	                                                           VARIANT* pvarTargetFrame,
	                                                           VARIANT* pvarData,
	                                                           VARIANT* pvarHeaders,
	                                                           VARIANT_BOOL* pbCancel) {
}

STDMETHODIMP_(void) InProcessDriverEngine::OnNavigateComplete2(LPDISPATCH pDisp,
                                                               VARIANT* URL) {
	if (this->browser_.IsEqualObject(pDisp)) {
		this->is_navigating_ = false;
		this->CreateWaitThread(this->m_hWnd);
	}
}

STDMETHODIMP_(void) InProcessDriverEngine::OnDocumentComplete(LPDISPATCH pDisp,
                                                              VARIANT* URL) {
}

STDMETHODIMP_(void) InProcessDriverEngine::OnNewWindow(LPDISPATCH ppDisp,
	                                                     VARIANT_BOOL* pbCancel,
	                                                     DWORD dwFlags,
	                                                     BSTR bstrUrlContext,
	                                                     BSTR bstrUrl) {
}

STDMETHODIMP_(void) InProcessDriverEngine::OnNewProcess(DWORD lCauseFlag,
	                                                      IDispatch* pWB2,
	                                                      VARIANT_BOOL* pbCancel) {
}

STDMETHODIMP_(void) InProcessDriverEngine::OnQuit() {
}

LRESULT InProcessDriverEngine::OnCopyData(UINT nMsg,
                                          WPARAM wParam,
                                          LPARAM lParam,
                                          BOOL& handled) {
  COPYDATASTRUCT* data = reinterpret_cast<COPYDATASTRUCT*>(lParam);
  this->return_window_ = reinterpret_cast<HWND>(wParam);
  std::vector<char> buffer(data->cbData + 1);
  memcpy_s(&buffer[0], data->cbData, data->lpData, data->cbData);
  buffer[buffer.size() - 1] = '\0';
  std::string received_data(&buffer[0]);
  if (received_data.find("\n") != std::string::npos) {
    std::vector<std::string> tokens;
    StringUtilities::Split(received_data, "\n", &tokens);
    this->current_command_ = tokens.at(0);
    this->command_data_ = tokens.at(1);
  } else {
    this->current_command_ = received_data;
  }
  
  return 0;
}

LRESULT InProcessDriverEngine::OnSetCommand(UINT nMsg,
                                            WPARAM wParam,
                                            LPARAM lParam,
                                            BOOL& handled) {
    LPCSTR command = reinterpret_cast<LPCSTR>(lParam);
    this->current_command_ = command;
    return 0;
}

LRESULT InProcessDriverEngine::OnGetResponseLength(UINT nMsg,
                                                   WPARAM wParam,
                                                   LPARAM lParam,
                                                   BOOL& handled) {
    return this->response_.size();
}

LRESULT InProcessDriverEngine::OnGetResponse(UINT nMsg,
                                             WPARAM wParam,
                                             LPARAM lParam,
                                             BOOL& handled) {
	HWND return_window_handle = NULL;
	if (wParam != NULL) {
		return_window_handle = reinterpret_cast<HWND>(wParam);
	}
	
	this->SendResponse(return_window_handle, this->response_);
  // Reset the serialized response for the next command.
  this->response_ = "";
  this->current_command_ = "";
	this->command_data_ = "";
  return 0;
}

LRESULT InProcessDriverEngine::OnWait(UINT nMsg,
                                      WPARAM wParam,
                                      LPARAM lParam,
                                      BOOL& handled) {
	if (this->IsDocumentReady()) {
		this->response_ = "done";
	} else {
		this->CreateWaitThread(this->m_hWnd);
	}
	return 0;
}

LRESULT InProcessDriverEngine::OnExecuteCommand(UINT nMsg,
                                                WPARAM wParam,
                                                LPARAM lParam,
                                                BOOL& handled) {
  std::string response = "";
  if (this->current_command_ == "getTitle") {
	  this->GetTitle();
  } else if (this->current_command_ == "get") {
    this->Navigate();
  } else if (this->current_command_ == "snapshot") {
    this->TakeSnapshot();
  } else if (this->current_command_ == "executeScript") {
    this->ExecuteScript();
  } else if (this->current_command_ == "experiment") {
    this->ExecuteExperimentalCommand();
  }

  return 0;
}

void InProcessDriverEngine::SendResponse(HWND window_handle,
                                         std::string response) {
	std::vector<char> buffer(response.size() + 1);
	memcpy_s(&buffer[0], response.size() + 1, response.c_str(), response.size());
	buffer[buffer.size() - 1] = '\0';
	COPYDATASTRUCT copy_data;
	copy_data.cbData = buffer.size();
	copy_data.lpData = reinterpret_cast<void*>(&buffer[0]);
	::SendMessage(window_handle,
                WM_COPYDATA,
                reinterpret_cast<WPARAM>(this->m_hWnd),
                reinterpret_cast<LPARAM>(&copy_data));
}

void InProcessDriverEngine::GetTitle() {
	CComPtr<IDispatch> dispatch;
	HRESULT hr = this->browser_->get_Document(&dispatch);
	if (FAILED(hr)) {
	}

	CComPtr<IHTMLDocument2> doc;
	hr = dispatch->QueryInterface(&doc);
	if (FAILED(hr)) {
	}

	CComBSTR title;
	hr = doc->get_title(&title);
	if (FAILED(hr)) {
	}

	std::wstring converted_title = title;
	this->response_ = StringUtilities::ToString(converted_title);
}

void InProcessDriverEngine::Navigate() {
	this->is_navigating_ = true;
	std::wstring url = StringUtilities::ToWString(this->command_data_);
	this->is_navigating_ = true;
	CComVariant dummy;
	CComVariant url_variant(url.c_str());
	HRESULT hr = this->browser_->Navigate2(&url_variant,
											&dummy,
											&dummy,
											&dummy,
											&dummy);
	if (FAILED(hr)) {
		this->is_navigating_ = false;
		_com_error error(hr);
		std::wstring formatted_message = 
        StringUtilities::Format(L"Received error: 0x%08x ['%s']",
			                          hr,
			                          error.ErrorMessage());
		this->response_ = StringUtilities::ToString(formatted_message);
	}
}

void InProcessDriverEngine::ExecuteExperimentalCommand() {
  // This method is intended to be a sandbox for testing
  // experimental commands. Right now, it's testing the
  // efficacy of using IDisplayServices for calculating
  // element positions.
  CComPtr<IDispatch> dispatch;
	HRESULT hr = this->browser_->get_Document(&dispatch);
	if (FAILED(hr)) {
	}

	CComPtr<IHTMLDocument2> doc;
	hr = dispatch->QueryInterface(&doc);
	if (FAILED(hr)) {
	}

	CComPtr<IHTMLElement> body;
	hr = doc->get_body(&body);

	CComPtr<IDisplayServices> display;
	hr = doc->QueryInterface<IDisplayServices>(&display);
	POINT p = { 0, 0 };
	hr = display->TransformPoint(&p,
                               COORD_SYSTEM_CLIENT,
                               COORD_SYSTEM_GLOBAL,
                               body);
	this->response_ = "over";
}

void InProcessDriverEngine::TakeSnapshot() {
  std::wstring script = L"var shotBlob = browser.takeVisualSnapshot();";
  script.append(L"var reader = new FileReader();");
  script.append(L"reader.onloadend = function() { external.sendMessage('snapshot', reader.result); };");
  script.append(L"reader.readAsText(shotBlob);");
  this->script_engine_->EvaluateScript(script.c_str(), L"");
}

void InProcessDriverEngine::ExecuteScript() {
  std::vector<std::string> script_info;
  StringUtilities::Split(this->command_data_, "\r", &script_info);

  std::wstring script_source = StringUtilities::ToWString(script_info[0]);

  std::vector<CComVariant> args;
  for (size_t index = 1; index < script_info.size(); ++index) {
    CComVariant arg = script_info[index].c_str();
    args.push_back(arg);
  }

  CComVariant function_creator_object;
  this->CreateAnonymousFunction(script_source, &function_creator_object);

  // This invocation creates and returns the anonymous function's IDispatch.
  // The second invocation executes the function.
  CComVariant function_object;
  std::vector<CComVariant> no_args;
  this->InvokeAnonymousFunction(function_creator_object,
                                no_args,
                                &function_object);
  CComVariant result;
  this->InvokeAnonymousFunction(function_object, args, &result);
  if (result.vt == VT_BSTR) {
    std::wstring result_string(result.bstrVal);
    this->response_ = StringUtilities::ToString(result_string);
  } else {
    this->response_ = "ERROR!!";
  }
}

void InProcessDriverEngine::CreateAnonymousFunction(const std::wstring& script,
                                                    CComVariant* function_object) {
  HRESULT hr = S_OK;

  CComPtr<IHTMLDocument2> doc;
  hr = this->GetFocusedDocument(&doc);
  if (FAILED(hr)) {
  }

  CComPtr<IDispatch> script_dispatch;
  hr = doc->get_Script(&script_dispatch);
  if (FAILED(hr)) {
    //LogError("Failed to get script IDispatch from document pointer", hr);
  }

  CComPtr<IDispatchEx> script_engine;
  hr = script_dispatch->QueryInterface<IDispatchEx>(&script_engine);
  if (FAILED(hr)) {
    //LogError("Failed to get script IDispatch from script IDispatch", hr);
  }

  std::wstring item_type = L"Function";
  CComBSTR name(item_type.c_str());

  std::wstring script_source = L"return ";
  script_source.append(script);

  CComVariant script_source_variant(script_source.c_str());

  std::vector<CComVariant> argument_array;
  argument_array.push_back(script_source_variant);

  DISPPARAMS ctor_parameters = { 0 };
  memset(&ctor_parameters, 0, sizeof ctor_parameters);
  ctor_parameters.cArgs = static_cast<unsigned int>(argument_array.size());
  ctor_parameters.rgvarg = &argument_array[0];

  // Find the javascript object using the IDispatchEx of the script engine
  DISPID dispatch_id;
  hr = script_engine->GetDispID(name, 0, &dispatch_id);
  if (FAILED(hr)) {
    //LogError("Failed to get DispID for Object constructor", hr);
  }

  // Create the jscript object by calling its constructor
  // The below InvokeEx call returns E_INVALIDARG in this case
  hr = script_engine->InvokeEx(dispatch_id,
                               LOCALE_USER_DEFAULT,
                               DISPATCH_CONSTRUCT,
                               &ctor_parameters,
                               function_object,
                               nullptr,
                               nullptr);
  if (FAILED(hr)) {
    //LogError("Failed to call InvokeEx on Object constructor", hr);
  }
}

void InProcessDriverEngine::InvokeAnonymousFunction(const CComVariant& function_object,
                                                    const std::vector<CComVariant>& args,
                                                    CComVariant* result) {
  HRESULT hr = S_OK;
  CComPtr<IDispatchEx> function_dispatch;
  hr = function_object.pdispVal->QueryInterface<IDispatchEx>(&function_dispatch);
  if (FAILED(hr)) {
  }

  // Grab the "call" method out of the returned function
  DISPID call_member_id;
  CComBSTR call_member_name = L"call";
  hr = function_dispatch->GetDispID(call_member_name, 0, &call_member_id);
  if (FAILED(hr)) {
  }

  CComPtr<IHTMLDocument2> doc;
  hr = this->GetFocusedDocument(&doc);
  if (FAILED(hr)) {
  }

  CComPtr<IHTMLWindow2> win;
  hr = doc->get_parentWindow(&win);
  if (FAILED(hr)) {
  }

  // IDispatch::Invoke() expects the arguments to be passed into it
  // in reverse order. To accomplish this, we create a new variant
  // array of size n + 1 where n is the number of arguments we have.
  // we copy each element of arguments_array_ into the new array in
  // reverse order, and add an extra argument, the window object,
  // to the end of the array to use as the "this" parameter for the
  // function invocation.
  std::vector<CComVariant> argument_array(args.size() + 1);
  for (size_t index = 0; index < args.size(); ++index) {
    argument_array[args.size() - 1 - index].Copy(&args[index]);
  }

  CComVariant window_variant(win);
  argument_array[argument_array.size() - 1].Copy(&window_variant);

  DISPPARAMS call_parameters = { 0 };
  memset(&call_parameters, 0, sizeof call_parameters);
  call_parameters.cArgs = static_cast<unsigned int>(argument_array.size());
  call_parameters.rgvarg = &argument_array[0];

  EXCEPINFO exception;
  memset(&exception, 0, sizeof exception);
  CComPtr<IServiceProvider> custom_exception_service_provider;
  hr = ScriptExceptionService::CreateInstance<IServiceProvider>(&custom_exception_service_provider);
  CComPtr<IScriptExceptionService> custom_exception;
  hr = custom_exception_service_provider.QueryInterface<IScriptExceptionService>(&custom_exception);
  hr = function_dispatch->InvokeEx(call_member_id,
                                   LOCALE_USER_DEFAULT,
                                   DISPATCH_METHOD,
                                   &call_parameters,
                                   result,
                                   &exception,
                                   custom_exception_service_provider);

  if (FAILED(hr)) {
  }
}

HRESULT InProcessDriverEngine::GetFocusedDocument(IHTMLDocument2** doc) {
  HRESULT hr = S_OK;
  CComPtr<IDispatch> dispatch;
  hr = this->browser_->get_Document(&dispatch);
  if (FAILED(hr)) {
    return hr;
  }

  CComPtr<IHTMLDocument2> browser_document;
  hr = dispatch->QueryInterface(&browser_document);
  if (FAILED(hr)) {
    return hr;
  }

  CComPtr<IHTMLWindow2> win;
  hr = browser_document->get_parentWindow(&win);
  if (FAILED(hr)) {
    return hr;
  }

  hr = win->get_document(doc);
  if (FAILED(hr)) {
    return hr;
  }

  return S_OK;
}

bool InProcessDriverEngine::IsDocumentReady() {
	HRESULT hr = S_OK;
	CComPtr<IDispatch> document_dispatch;
	hr = this->browser_->get_Document(&document_dispatch);
	if (FAILED(hr) || document_dispatch == nullptr) {
		return false;
	}

	CComPtr<IHTMLDocument2> doc;
	hr = document_dispatch->QueryInterface<IHTMLDocument2>(&doc);
	if (FAILED(hr) || doc == nullptr) {
		return false;
	}

	CComBSTR ready_state_bstr;
	hr = doc->get_readyState(&ready_state_bstr);
	if (FAILED(hr)) {
		return false;
	}

	return ready_state_bstr == L"complete";
}

void InProcessDriverEngine::CreateWaitThread(HWND return_handle) {
	// If we are still waiting, we need to wait a bit then post a message to
	// ourselves to run the wait again. However, we can't wait using Sleep()
	// on this thread. This call happens in a message loop, and we would be
	// unable to process the COM events in the browser if we put this thread
	// to sleep.
	WaitThreadContext* thread_context = new WaitThreadContext;
	thread_context->window_handle = this->m_hWnd;
	unsigned int thread_id = 0;
	HANDLE thread_handle = reinterpret_cast<HANDLE>(_beginthreadex(NULL,
																	0,
																	&InProcessDriverEngine::WaitThreadProc,
																	reinterpret_cast<void*>(thread_context),
																	0,
																	&thread_id));
	if (thread_handle != NULL) {
		::CloseHandle(thread_handle);
	}
}


unsigned int WINAPI InProcessDriverEngine::WaitThreadProc(LPVOID lpParameter) {
	WaitThreadContext* thread_context = reinterpret_cast<WaitThreadContext*>(lpParameter);
	HWND window_handle = thread_context->window_handle;
	delete thread_context;

	::Sleep(50);
	::PostMessage(window_handle, WM_WAIT, NULL, NULL);
	return 0;
}
