
//
// Main header
// by The trick
// 2019


#include "stdafx.h"

typedef HRESULT (__stdcall *pfnOleLoadPictureEx)(
  IStream *lpstream,
  LONG     lSize,
  BOOL     fRunmode,
  REFIID   riid,
  DWORD    xSizeDesired,
  DWORD    ySizeDesired,
  DWORD    dwFlags,
  LPVOID   *lplpvObj
);

typedef HRESULT (__stdcall *pfnOleLoadPicture)(
  IStream *lpstream,
  LONG     lSize,
  BOOL     fRunmode,
  REFIID   riid,
  LPVOID   *lplpvObj
);

// Initialize the module
BOOL Initialize();

// Uninitialize module
VOID Uninitialize();

// Check if dll can be safely unloaded
HRESULT CanUnloadNow();

extern volatile ULONG g_lCountOfUsers;	// Number of users of DLL
extern volatile ULONG g_lCountOfObject;	// Count of active object
extern HMODULE g_hModule;				// Module handle

// Intercepted functions
HRESULT __stdcall OleLoadPictureEx_user(
				  IStream *lpstream,
				  LONG     lSize,
				  BOOL     fRunmode,
				  REFIID   riid,
				  DWORD    xSizeDesired,
				  DWORD    ySizeDesired,
				  DWORD    dwFlags,
				  LPVOID   *lplpvObj);

HRESULT __stdcall OleLoadPicture_user(
				  IStream *lpstream,
				  LONG     lSize,
				  BOOL     fRunmode,
				  REFIID   riid,
				  LPVOID   *lplpvObj);