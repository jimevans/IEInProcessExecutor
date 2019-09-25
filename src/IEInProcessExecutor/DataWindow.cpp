#include "pch.h"
#include "DataWindow.h"

#include <vector>

LRESULT DataWindow::OnCopyData(UINT nMsg, WPARAM wParam, LPARAM lParam, _Inout_ BOOL& /*bHandled*/) {
	COPYDATASTRUCT* data = reinterpret_cast<COPYDATASTRUCT*>(lParam);
	std::vector<char> buffer(data->cbData + 1);
	memcpy_s(&buffer[0], data->cbData, data->lpData, data->cbData);
	buffer[buffer.size() - 1] = '\0';
	std::string received_data(&buffer[0]);
	this->response_ = received_data;
	return 0;
}
