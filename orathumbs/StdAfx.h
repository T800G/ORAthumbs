#ifndef _STDAFX_H__DB305830_65E8_4806_9B51_3FCF5B9803D1__INCLUDED_
#define _STDAFX_H__DB305830_65E8_4806_9B51_3FCF5B9803D1__INCLUDED_
#pragma once

#ifndef UNICODE
#define UNICODE// UNICODE-only project
#endif

#define STRICT

//#define WINVER		0x0500
#define _WIN32_WINNT	_WIN32_WINNT_WINXP
//#define _WIN32_IE	_WIN32_IE_IE55 // StrCmpLogicalW needs 5.5 min
//#define _RICHEDIT_VER	0x0200

#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
extern CComModule _Module;
#include <atlcom.h>


#endif //_STDAFX_H__DB305830_65E8_4806_9B51_3FCF5B9803D1__INCLUDED
