#pragma once

#include <string>
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

using namespace ATL;

class DataWindow : public CWindowImpl<DataWindow>
{
public:
	DataWindow() {
		this->response_ = "";
	}

	BEGIN_MSG_MAP(InProcessDriverEngine)
		MESSAGE_HANDLER(WM_COPYDATA, OnCopyData)
	END_MSG_MAP()

public:
	LRESULT OnCopyData(UINT nMsg, WPARAM wParam, LPARAM lParam, _Inout_ BOOL& /*bHandled*/);
	std::string response() { return this->response_; }
	void Reset(void) { this->response_ = ""; }

private:
	std::string response_;
};

