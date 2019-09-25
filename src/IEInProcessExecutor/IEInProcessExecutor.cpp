// JSObjectCreator.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include <ctime>
#include <iostream>
#include <string>
#include <vector>

#include <atlenc.h>
#include <atlimage.h>
#include <iepmapi.h>
#include <ShlGuid.h>

#include "../ScriptEngineProxy/ScriptEngineProxy_i.h"
#include "../ScriptEngineProxy/messages.h"

#include "DataWindow.h"

struct ProcessWindowInfo {
  DWORD dwProcessId;
  HWND hwndBrowser;
};

struct Command {
  std::string name;
  std::vector<std::string> args;

  void Reset(void) {
    name.clear();
    args.clear();
  }
  
  std::string Serialize(void) {
    // Serialized command string format:
    // <commandName>\n<commandArguments>
    // where <commandArguments> are separated by \r.
    std::string serialized_command = name;
    if (args.size() > 0) {
      std::string serialized_args;
      std::vector<std::string>::const_iterator it = args.begin();
      for (; it != args.end(); ++it) {
        if (serialized_args.size() > 0) {
          serialized_args.append("\r");
        }
        serialized_args.append(*it);
      }
      serialized_command.append("\n");
      serialized_command.append(serialized_args);
    }
    return serialized_command;
  }

  std::string ToString(void) {
    std::string output = name;
    if (args.size() > 0) {
      output.append("\n");
      output.append("  with arguments: \n");
      std::string args_output;
      std::vector<std::string>::const_iterator it = args.begin();
      for (; it != args.end(); ++it) {
        if (args_output.size() > 0) {
          args_output.append("\n");
        }
        args_output.append("    ").append(*it);
      }
      output.append(args_output);
    }
    return output;
  }
};

#define BROWSER_ATTACH_TIMEOUT_IN_MILLISECONDS 10000
#define BROWSER_ATTACH_SLEEP_INTERVAL_IN_MILLISECONDS 250
#define WAIT_FOR_IDLE_TIMEOUT_IN_MILLISECONDS 2000
#define IE_FRAME_WINDOW_CLASS "IEFrame"
#define SHELL_DOCOBJECT_VIEW_WINDOW_CLASS "Shell DocObject View"
#define IE_SERVER_CHILD_WINDOW_CLASS "Internet Explorer_Server"
#define HTML_GETOBJECT_MSG L"WM_HTML_GETOBJECT"
#define OLEACC_LIBRARY_NAME L"OLEACC.DLL"
#define OBJECT_FROM_LRESULT_FUNCTION_NAME "ObjectFromLresult"
#define IDM_STARTDIAGNOSTICSMODE 3802 

void LogError(const std::string& message, HRESULT hr) {
  std::cout << message << std::endl;
  DWORD error_code = HRESULT_CODE(hr);
  std::cout << "Error code received: " << error_code << std::endl;
  std::vector<char> system_message_buffer(256);
  DWORD is_formatted = ::FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                                        nullptr,
                                        error_code,
                                        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                        &system_message_buffer[0],
                                        system_message_buffer.size(),
                                        nullptr);
  if (is_formatted > 0) {
    std::string system_message = &system_message_buffer[0];
    std::cout << "Error message received: " << system_message << std::endl;
  }
}

BOOL CALLBACK FindChildWindowForProcess(HWND hwnd, LPARAM arg) {
  ProcessWindowInfo* process_window_info = reinterpret_cast<ProcessWindowInfo*>(arg);

  // Could this be an Internet Explorer Server window?
  // 25 == "Internet Explorer_Server\0"
  char name[25];
  if (::GetClassNameA(hwnd, name, 25) == 0) {
    // No match found. Skip
    return TRUE;
  }

  if (strcmp(IE_SERVER_CHILD_WINDOW_CLASS, name) != 0) {
    return TRUE;
  }
  else {
    DWORD process_id = NULL;
    ::GetWindowThreadProcessId(hwnd, &process_id);
    if (process_window_info->dwProcessId == process_id) {
      // Once we've found the first Internet Explorer_Server window
      // for the process we want, we can stop.
      process_window_info->hwndBrowser = hwnd;
      return FALSE;
    }
  }

  return TRUE;
}

BOOL CALLBACK FindBrowserWindow(HWND hwnd, LPARAM arg) {
  // Could this be an IE instance?
  // 8 == "IeFrame\0"
  // 21 == "Shell DocObject View\0";
  char name[21];
  if (::GetClassNameA(hwnd, name, 21) == 0) {
    // No match found. Skip
    return TRUE;
  }

  if (strcmp(IE_FRAME_WINDOW_CLASS, name) != 0 &&
      strcmp(SHELL_DOCOBJECT_VIEW_WINDOW_CLASS, name) != 0) {
    return TRUE;
  }

  return EnumChildWindows(hwnd, FindChildWindowForProcess, arg);
}

bool GetDocumentFromWindowHandle(HWND window_handle,
                                 IHTMLDocument2** document) {
  bool is_document_retrieved = false;
  UINT html_getobject_msg = ::RegisterWindowMessage(HTML_GETOBJECT_MSG);

  // Explicitly load MSAA so we know if it's installed
  HINSTANCE oleacc_instance_handle = ::LoadLibrary(OLEACC_LIBRARY_NAME);
  if (window_handle != nullptr && oleacc_instance_handle) {
    LRESULT result;

    ::SendMessageTimeout(window_handle,
                         html_getobject_msg,
                         0L,
                         0L,
                         SMTO_ABORTIFHUNG,
                         1000,
                         reinterpret_cast<PDWORD_PTR>(&result));

    LPFNOBJECTFROMLRESULT object_pointer =
        reinterpret_cast<LPFNOBJECTFROMLRESULT>(::GetProcAddress(oleacc_instance_handle,
                                                                 OBJECT_FROM_LRESULT_FUNCTION_NAME));
    if (object_pointer != nullptr) {
      HRESULT hr;
      hr = (*object_pointer)(result,
                             IID_IHTMLDocument2,
                             0,
                             reinterpret_cast<void**>(document));
      if (SUCCEEDED(hr)) {
        is_document_retrieved = true;
      }
    }
  }

  if (oleacc_instance_handle) {
    ::FreeLibrary(oleacc_instance_handle);
  }
  return is_document_retrieved;
}

bool LaunchIE(const std::wstring& initial_browser_url,
              IHTMLDocument2** document) {
  bool is_launched = false;
  DWORD browser_attach_timeout = BROWSER_ATTACH_TIMEOUT_IN_MILLISECONDS;
  PROCESS_INFORMATION proc_info;
  HRESULT hr = ::IELaunchURL(initial_browser_url.c_str(),
                             &proc_info,
                             nullptr);
  DWORD process_id = proc_info.dwProcessId;
  ::WaitForInputIdle(proc_info.hProcess,
                     WAIT_FOR_IDLE_TIMEOUT_IN_MILLISECONDS);

  if (nullptr != proc_info.hThread) {
    ::CloseHandle(proc_info.hThread);
  }

  if (nullptr != proc_info.hProcess) {
    ::CloseHandle(proc_info.hProcess);
  }

  ProcessWindowInfo process_window_info;
  process_window_info.dwProcessId = proc_info.dwProcessId;
  process_window_info.hwndBrowser = nullptr;
  clock_t end = clock() + (browser_attach_timeout / 1000 * CLOCKS_PER_SEC);
  while (nullptr == process_window_info.hwndBrowser) {
    if (browser_attach_timeout > 0 && (clock() > end)) {
      break;
    }
    ::EnumWindows(&FindBrowserWindow,
                  reinterpret_cast<LPARAM>(&process_window_info));
    if (nullptr == process_window_info.hwndBrowser) {
      ::Sleep(BROWSER_ATTACH_SLEEP_INTERVAL_IN_MILLISECONDS);
    }
  }
  if (process_window_info.hwndBrowser) {
    is_launched = GetDocumentFromWindowHandle(process_window_info.hwndBrowser,
                                              document);
  }
  return is_launched;
}

bool GetBrowserObject(IHTMLDocument2* doc, IWebBrowser2** browser) {
  HRESULT hr = S_OK;
  CComPtr<IHTMLWindow2> parent_window;
  hr = doc->get_parentWindow(&parent_window);
  if (FAILED(hr) || parent_window == nullptr) {
    return false;
  }

  CComPtr<IServiceProvider> provider;
  hr = parent_window->QueryInterface<IServiceProvider>(&provider);
  if (FAILED(hr) || provider == nullptr) {
    return false;
  }
  CComPtr<IServiceProvider> child_provider;
  hr = provider->QueryService(SID_STopLevelBrowser,
                              IID_IServiceProvider,
                              reinterpret_cast<void**>(&child_provider));
  if (FAILED(hr) || child_provider == nullptr) {
    return false;
  }
  hr = child_provider->QueryService(SID_SWebBrowserApp,
                                    IID_IWebBrowser2,
                                    reinterpret_cast<void**>(browser));
  if (FAILED(hr) || *browser == nullptr) {
    return false;
  }
  return true;
}

std::wstring GetExecutablePath(void) {
  // Get the current path that we are running from
  std::vector<wchar_t> full_path_buffer(MAX_PATH);
  DWORD count = ::GetModuleFileName(nullptr, &full_path_buffer[0], MAX_PATH);
  full_path_buffer.resize(count);

  std::wstring full_path(&full_path_buffer[0]);
  size_t pos = full_path.find_last_of(L'\\');
  std::wstring path = full_path.substr(0, pos);

  return path.append(L"\\ScriptEngineProxy.dll");
}

HRESULT StartDiagnosticsMode(_In_ IHTMLDocument2* pDocument,
                             REFCLSID clsid,
                             _In_ LPCWSTR path,
                             REFIID iid,
                             _COM_Outptr_opt_ void** ppOut) {
  // Get the target from the document
  CComQIPtr<IOleCommandTarget> spCommandTarget(pDocument);
  ATLENSURE_RETURN_HR(spCommandTarget.p != nullptr, E_INVALIDARG);
  // Setup the diagnostics mode parameters
  HRESULT hr = S_OK;
  CComBSTR guid(clsid);
  CComSafeArray<BSTR> spSafeArray(4);
  hr = spSafeArray.SetAt(0, ::SysAllocString(guid), FALSE);
  ATLENSURE_RETURN_HR(hr == S_OK, SUCCEEDED(hr) ? E_FAIL : hr);
  hr = spSafeArray.SetAt(1, ::SysAllocString(path), FALSE);
  ATLENSURE_RETURN_HR(hr == S_OK, SUCCEEDED(hr) ? E_FAIL : hr);
  hr = spSafeArray.SetAt(2, ::SysAllocString(L""), FALSE);
  ATLENSURE_RETURN_HR(hr == S_OK, SUCCEEDED(hr) ? E_FAIL : hr);
  hr = spSafeArray.SetAt(3, ::SysAllocString(L""), FALSE);
  ATLENSURE_RETURN_HR(hr == S_OK, SUCCEEDED(hr) ? E_FAIL : hr);

  // Start diagnostics mode
  CComVariant varParams(spSafeArray);
  CComVariant varSite;
  hr = spCommandTarget->Exec(&CGID_MSHTML,
                             IDM_STARTDIAGNOSTICSMODE,
                             OLECMDEXECOPT_DODEFAULT,
                             &varParams,
                             &varSite);
  ATLENSURE_RETURN_HR(hr == S_OK, SUCCEEDED(hr) ? E_FAIL : hr);
  ATLENSURE_RETURN_VAL(varSite.vt == VT_UNKNOWN, E_UNEXPECTED);
  ATLENSURE_RETURN_VAL(varSite.punkVal != nullptr, E_UNEXPECTED);
  // Get the requested interface
  if (ppOut) {
    hr = varSite.punkVal->QueryInterface(iid, ppOut);
    ATLENSURE_RETURN_HR(hr == S_OK, SUCCEEDED(hr) ? E_FAIL : hr);
  }
  return hr;
}

void SetCommand(HWND window_handle, Command command) {
  std::string serialized_command = command.Serialize();
  COPYDATASTRUCT command_data;
  command_data.cbData = serialized_command.size();
  command_data.lpData = reinterpret_cast<void*>(const_cast<char*>(serialized_command.c_str()));
  ::SendMessage(window_handle,
                WM_COPYDATA,
                NULL,
                reinterpret_cast<LPARAM>(&command_data));
}

bool WaitForCommandComplete(HWND window_handle, int* response_size) {
  LRESULT response_length = ::SendMessage(window_handle,
                                          WM_GET_RESPONSE_LENGTH,
                                          NULL,
                                          NULL);
  while (response_length == 0) {
    ::Sleep(10);
    response_length = ::SendMessage(window_handle,
                                    WM_GET_RESPONSE_LENGTH,
                                    NULL,
                                    NULL);
  }

  *response_size = static_cast<int>(response_length);
  return true;
}

std::string GetResponseUsingWindow(HWND window_handle,
                                   DataWindow* return_window) {
  ::SendMessage(window_handle,
                WM_GET_RESPONSE,
                reinterpret_cast<WPARAM>(return_window->m_hWnd),
                NULL);
  return return_window->response();
}

std::string ExecuteCommand(Command command,
                          HWND window_handle,
                          DataWindow* return_window) {
  std::cout << "Executing command: " << command.ToString() << std::endl;
  SetCommand(window_handle, command);

  ::PostMessage(window_handle,
                WM_EXECUTE_COMMAND,
                reinterpret_cast<WPARAM>(return_window->m_hWnd),
                NULL);

  int response_length = 0;
  WaitForCommandComplete(window_handle, &response_length);

  std::string response = GetResponseUsingWindow(window_handle, return_window);
  std::cout << "Got response: " << response << std::endl;
  return response;
}

int main() {
  int created_status = 0;
  ::CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

  DataWindow return_window;
  return_window.Create(HWND_MESSAGE);
  BOOL allow = ::ChangeWindowMessageFilterEx(return_window.m_hWnd,
                                             WM_COPYDATA,
                                             MSGFLT_ALLOW,
                                             nullptr);

  std::wstring initial_browser_url = L"https://stackoverflow.com";
  CComPtr<IHTMLDocument2> document;
  bool is_launched = LaunchIE(initial_browser_url, &document);
  if (is_launched) {
    CComPtr<IWebBrowser2> browser;
    bool is_valid_browser = false;
    is_valid_browser = GetBrowserObject(document, &browser);
    std::wstring path = GetExecutablePath();
    CComPtr<IOleWindow> site;
    HRESULT hr = StartDiagnosticsMode(document,
                                      __uuidof(InProcessDriverEngine),
                                      path.c_str(),
                                      __uuidof(site),
                                      reinterpret_cast<void**>(&site.p));
    HWND window_handle;
    site->GetWindow(&window_handle);

    // Execute getTitle command as proof of concept that command-response
    // handling works properly using this architecture.
    Command command;
    command.name = "getTitle";
    std::string response = ExecuteCommand(command,
                                          window_handle,
                                          &return_window);

    // Execute executeScript command with arbitrary JavaScript as
    // proof of concept that doing so is valid in this archictecture.
    command.Reset();
    command.name = "executeScript";
    command.args.push_back("function() { return arguments[0] + arguments[1]; }");
    command.args.push_back("foo");
    command.args.push_back("bar");
    response = ExecuteCommand(command, window_handle, &return_window);

    // Executing executeScript command using parentWindow property
    // crashes in mshtml.dll. To see the crash, set a breakpoint,
    // and attach to the iexplore.exe process.
    command.Reset();
    command.name = "executeScript";
    command.args.push_back("function() { return document.parentWindow.innerHeight; }");
    response = ExecuteCommand(command, window_handle, &return_window);

    if (is_valid_browser) {
      browser->Quit();
    }
  }
  else {
    std::cout << "Failed to launch IE and get reference to document"
              << std::endl;
  }
  std::cout << "Process complete" << std::endl;

  ::CoUninitialize();
  return_window.DestroyWindow();
  return created_status;
}
