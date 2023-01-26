#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "wininet")

#include <windows.h>
#include <wininet.h>
#include <iostream>
#include "json11.hpp"

WCHAR szClassName[] = L"openai";
WCHAR szAPIKey[] = L"YOUR_API_KEY";

LPWSTR a2w(LPCSTR lpszText)
{
	const DWORD dwSize = MultiByteToWideChar(CP_UTF8, 0, lpszText, -1, 0, 0);
	LPWSTR lpszReturnText = (LPWSTR)GlobalAlloc(0, dwSize * sizeof(WCHAR));
	MultiByteToWideChar(CP_UTF8, 0, lpszText, -1, lpszReturnText, dwSize);
	return lpszReturnText;
}

LPSTR w2a(LPCWSTR lpszText)
{
	const DWORD dwSize = WideCharToMultiByte(CP_UTF8, 0, lpszText, -1, 0, 0, 0, 0);
	LPSTR lpszReturnText = (char*)GlobalAlloc(GPTR, dwSize * sizeof(char));
	WideCharToMultiByte(CP_UTF8, 0, lpszText, -1, lpszReturnText, dwSize, 0, 0);
	return lpszReturnText;
}

LPWSTR openai(IN LPCWSTR lpszMessage)
{
	LPBYTE lpszByte = 0;
	DWORD dwSize = 0;
	const HINTERNET hSession = InternetOpenW(0, INTERNET_OPEN_TYPE_PRECONFIG, 0, 0, INTERNET_FLAG_NO_COOKIES);
	if (hSession)
	{
		const HINTERNET hConnection = InternetConnectW(hSession, L"api.openai.com", INTERNET_DEFAULT_HTTPS_PORT, 0, 0, INTERNET_SERVICE_HTTP, 0, 0);
		if (hConnection)
		{
			const HINTERNET hRequest = HttpOpenRequestW(hConnection, L"POST", L"/v1/completions", 0, 0, 0, INTERNET_FLAG_NO_CACHE_WRITE | INTERNET_FLAG_SECURE, 0);
			if (hRequest)
			{
				WCHAR szHeaders[1024];
				wsprintf(szHeaders, L"Content-Type: application/json\r\nAuthorization: Bearer %s", szAPIKey);
				std::string body_json = "{\"model\":\"text-davinci-003\", \"prompt\":\"";
				LPSTR lpszMessageA = w2a(lpszMessage);
				body_json += lpszMessageA;
				body_json += "\", \"temperature\": 0, \"max_tokens\": 1000}";
				GlobalFree(lpszMessageA);
				HttpSendRequestW(hRequest, szHeaders, lstrlenW(szHeaders), (LPVOID)(body_json.c_str()), (DWORD)body_json.length());
				lpszByte = (LPBYTE)GlobalAlloc(GMEM_FIXED, 1);
				if (lpszByte) {
					DWORD dwRead;
					BYTE szBuf[1024 * 4];
					for (;;)
					{
						if (!InternetReadFile(hRequest, szBuf, (DWORD)sizeof(szBuf), &dwRead) || !dwRead) break;
						LPBYTE lpTmp = (LPBYTE)GlobalReAlloc(lpszByte, (SIZE_T)(dwSize + dwRead), GMEM_MOVEABLE);
						if (lpTmp == NULL) break;
						lpszByte = lpTmp;
						CopyMemory(lpszByte + dwSize, szBuf, dwRead);
						dwSize += dwRead;
					}
				}
				InternetCloseHandle(hRequest);
			}
			InternetCloseHandle(hConnection);
		}
		InternetCloseHandle(hSession);
	}
	if (lpszByte)
	{
		std::string src((LPSTR)lpszByte, dwSize);
		GlobalFree(lpszByte);
		lpszByte = 0;
		std::string err;
		json11::Json v = json11::Json::parse(src, err);
		for (auto& item : v["choices"].array_items()) {
			return a2w(item["text"].string_value().c_str());
		}
	}
	LPCWSTR lpszErrorMessage = L"エラーとなりました。しばらくたってからリトライしてください。";
	LPWSTR lpszReturn  = (LPWSTR)GlobalAlloc(0, sizeof(WCHAR) * (lstrlenW(lpszErrorMessage) + 1));
	if (lpszReturn) {
		lstrcpy(lpszReturn, lpszErrorMessage);
	}
	return lpszReturn;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hButton;
	static HWND hEdit1;
	static HWND hEdit2;
	switch (msg)
	{
	case WM_CREATE:
		hEdit1 = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", 0, WS_VISIBLE | WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hButton = CreateWindowW(L"BUTTON", L"送信", WS_VISIBLE | WS_CHILD | WS_TABSTOP, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit2 = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", 0, WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_VSCROLL | ES_MULTILINE | ES_READONLY, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		break;
	case WM_SIZE:
		MoveWindow(hEdit1, 10, 10, 512, 32, TRUE);
		MoveWindow(hButton, 532, 10, 256, 32, TRUE);
		MoveWindow(hEdit2, 10, 50, LOWORD(lParam) - 20, HIWORD(lParam) - 60, TRUE);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			EnableWindow(hButton, FALSE);
			INT nSize = GetWindowTextLength(hEdit1);
			LPWSTR lpszMessage = (LPWSTR)GlobalAlloc(0, sizeof(WCHAR) * (nSize + 1));
			if (lpszMessage) {
				GetWindowText(hEdit1, lpszMessage, nSize + 1);
				LPWSTR lpszReturn = openai(lpszMessage);
				if (lpszReturn) {
					SetWindowText(hEdit2, 0);
					WCHAR seps[] = L"\n";
					WCHAR* context;
					LPWSTR token = wcstok_s(lpszReturn, seps, &context);
					while (token != NULL)
					{
						SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)token);
						SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)L"\r\n");
						token = wcstok_s(NULL, seps, &context);
					}
					GlobalFree(lpszReturn);
				}
				else {
					SetWindowText(hEdit2, L"エラー");
				}
				GlobalFree(lpszMessage);
			}
			EnableWindow(hButton, TRUE);
		}
		break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefDlgProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		DLGWINDOWEXTRA,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindowW(
		szClassName,
		L"openai",
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!IsDialogMessage(hWnd, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	return (int)msg.wParam;
}
