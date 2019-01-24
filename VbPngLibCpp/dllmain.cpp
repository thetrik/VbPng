
// 
// Entry point of DLL/LIB
// by The trick
// 2019


#include "stdafx.h"
#include "VBPng.h"

namespace std { const nothrow_t nothrow = nothrow_t(); }

volatile ULONG g_lCountOfObject = 0;	// Total number of active objects
HMODULE g_hModule = NULL;				// Base address of module

#ifndef _LIB
// For dll
BOOL APIENTRY DllMain(HMODULE hModule,
                      DWORD  ul_reason_for_call,
                      LPVOID lpReserved) {

	switch (ul_reason_for_call) {
	case DLL_PROCESS_ATTACH:
		g_hModule = hModule;
	}

	return TRUE;
}
#endif