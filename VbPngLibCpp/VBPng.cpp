//
// Main module
// by The trick
// 2019


#include "VBPng.h"
#include "ldasm\\LDasm.h"
#include "CHooker.h"
#include "CPicture.h"
#include <olectl.h>

#ifndef _LIB
#pragma comment(lib, "gdiplus.lib")
#endif

volatile ULONG g_lCountOfUsers = 0;			// Count of users
BOOL	g_bIsInitialized	= FALSE;		// If the module already intialized equals TRUE
ULONG	g_hGdipToken		= NULL;			// GDIplus token
CHooker	g_cHooker[2];						// Hooker object

BOOL Initialize() {
	HMODULE hOleAut				= NULL;
	pfnOleLoadPictureEx pfnEx	= NULL;
	pfnOleLoadPicture pfn 		= NULL;
	TCHAR pszAlreadyHooked[2];

	if (g_bIsInitialized) {
		InterlockedIncrement(&g_lCountOfUsers);
		return TRUE;
	}

	// Check if already hooked
	if (GetEnvironmentVariable(_T("VBPng"), pszAlreadyHooked, sizeof(pszAlreadyHooked) / sizeof(TCHAR)))
		if (pszAlreadyHooked[0] == _T('1')) {
			InterlockedIncrement(&g_lCountOfUsers);
			return TRUE;
		}
	
	// Initialize GDI+
	GdiplusStartup(&g_hGdipToken, &GdiplusStartupInput(), NULL);

	if (NULL == (hOleAut = GetModuleHandle(_T("oleaut32")))) {
		if (NULL == (hOleAut = LoadLibrary(_T("oleaut32")))) {
			MessageBox(NULL, _T("Unable to load oleaut32.dll"), NULL, MB_ICONERROR);
			return FALSE;
		}
	}

	if (NULL == (pfnEx = (pfnOleLoadPictureEx)GetProcAddress(hOleAut, "OleLoadPictureEx"))) {
		MessageBox(NULL, _T("Unable to get the OleLoadPictureEx address"), NULL, MB_ICONERROR);
		return FALSE;
	}
	
	if (NULL == (pfn = (pfnOleLoadPicture)GetProcAddress(hOleAut, "OleLoadPicture"))) {
		MessageBox(NULL, _T("Unable to get the OleLoadPicture address"), NULL, MB_ICONERROR);
		return FALSE;
	}

	// Hijack OleLoadPictureEx, OleLoadPicture
	if (g_cHooker[0].Hook(pfnEx, OleLoadPictureEx_user)) {
		if (g_cHooker[1].Hook(pfn, OleLoadPicture_user))
			g_bIsInitialized = TRUE;
		else
			g_cHooker[0].Unhook();
	}

	// If other dlls will use this functionaly
	if (g_bIsInitialized) {
		SetEnvironmentVariable(_T("VBPng"), _T("1"));
		InterlockedIncrement(&g_lCountOfUsers);
	}

	return g_bIsInitialized;

}

VOID Uninitialize() {

	if (g_bIsInitialized && InterlockedDecrement(&g_lCountOfUsers) == 0) {

		g_cHooker[0].Unhook();
		g_cHooker[1].Unhook();
		GdiplusShutdown(g_hGdipToken);
		
		SetEnvironmentVariable(_T("VBPng"), _T("0"));
	}

}

HRESULT CanUnloadNow() {
	if (g_lCountOfObject)
		return S_FALSE;
	else
		return S_OK;
}

HRESULT __stdcall OleLoadPictureEx_user(
				  IStream *lpstream,
				  LONG     lSize,
				  BOOL     fRunmode,
				  REFIID   riid,
				  DWORD    xSizeDesired,
				  DWORD    ySizeDesired,
				  DWORD    dwFlags,
				  LPVOID   *lplpvObj) {
	HRESULT hr		= S_OK;
	CPicture *cPic	= NULL;

	pfnOleLoadPictureEx pfn = (pfnOleLoadPictureEx)g_cHooker[0].GetThunkPtr();
	
	// Call the original one
	hr = pfn(lpstream, lSize, fRunmode, riid, xSizeDesired, ySizeDesired, dwFlags, lplpvObj);

	if (FAILED(hr)) {

		cPic = new (std::nothrow) CPicture();

		if (!cPic) {
			hr = E_OUTOFMEMORY;
			goto CleanUp;
		}

		if (FAILED(hr = cPic->put_KeepOriginalFormat(!fRunmode)))
			goto CleanUp;

		if (FAILED(hr = cPic->Load(lpstream)))
			goto CleanUp;

		if (FAILED(hr = cPic->QueryInterface(riid, lplpvObj)))
			goto CleanUp;

	}

CleanUp:

	if (FAILED(hr))
		if (cPic)
			delete cPic;

	return hr;

}

HRESULT __stdcall OleLoadPicture_user(
				  IStream *lpstream,
				  LONG     lSize,
				  BOOL     fRunmode,
				  REFIID   riid,
				  LPVOID   *lplpvObj) {
	HRESULT hr		= S_OK;
	CPicture *cPic	= NULL;

	pfnOleLoadPicture pfn = (pfnOleLoadPicture)g_cHooker[1].GetThunkPtr();
	
	// Call the original one
	hr = pfn(lpstream, lSize, fRunmode, riid, lplpvObj);

	if (FAILED(hr))
		hr = OleLoadPictureEx_user(lpstream, lSize, fRunmode, riid, 0, 0, 0, lplpvObj);

	return hr;

}
