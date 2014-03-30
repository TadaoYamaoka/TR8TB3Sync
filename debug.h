#pragma once

#ifdef _DEBUG
#include <windows.h>
#include <strsafe.h>
#define TRACE(X, ...) DEBUG_PRINTF(X, __VA_ARGS__)
void DEBUG_PRINTF(const wchar_t *format, ...)
{
	wchar_t buf[1024];
	va_list arg_list;
	va_start(arg_list, format);
	StringCbVPrintf(buf, sizeof(buf), format, arg_list);
	va_end(arg_list);
	OutputDebugString(buf);
}
#else
#define TRACE(X, ...)
#endif // _DEBUG
