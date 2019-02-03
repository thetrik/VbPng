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

volatile ULONG g_lCountOfUsers	= 0;			// Count of users
BOOL g_bIsInitialized			= FALSE;		// If the module already intialized equals TRUE
ULONG g_hGdipToken				= NULL;			// GDIplus token
CHooker g_cHooker[3];							// Hooker object
CPicturesServer g_cServer;						// Pictures server

BOOL Initialize() {
	HMODULE hOleAut				= NULL;
	pfnOleLoadPictureEx pfnEx	= NULL;
	pfnOleLoadPicture pfn 		= NULL;
	pfnOleIconToCursor pfnITC	= NULL;
	BOOL bComInit				= FALSE;
	BOOL bGdipInit				= FALSE;
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
	
	// Initialize COM
	if (FAILED(CoInitialize(NULL))) {
		MessageBox(NULL, _T("Unable to initialize COM"), NULL, MB_ICONERROR);
		goto CleanUp;
	}

	bComInit = TRUE;

	// Initialize GDI+
	if (GdiplusStartup(&g_hGdipToken, &GdiplusStartupInput(), NULL) != Ok) {
		MessageBox(NULL, _T("Unable to initialize GDI+"), NULL, MB_ICONERROR);
		goto CleanUp;
	}

	bGdipInit = TRUE;

	// Get oleaut32 handle
	if (NULL == (hOleAut = GetModuleHandle(_T("oleaut32")))) {
		if (NULL == (hOleAut = LoadLibrary(_T("oleaut32")))) {
			MessageBox(NULL, _T("Unable to load oleaut32.dll"), NULL, MB_ICONERROR);
			goto CleanUp;
		}
	}

	// Get OleLoadPictureEx function address
	if (NULL == (pfnEx = (pfnOleLoadPictureEx)GetProcAddress(hOleAut, "OleLoadPictureEx"))) {
		MessageBox(NULL, _T("Unable to get the OleLoadPictureEx address"), NULL, MB_ICONERROR);
		goto CleanUp;
	}

	// Get OleLoadPicture function address
	if (NULL == (pfn = (pfnOleLoadPicture)GetProcAddress(hOleAut, "OleLoadPicture"))) {
		MessageBox(NULL, _T("Unable to get the OleLoadPicture address"), NULL, MB_ICONERROR);
		goto CleanUp;
	}

	// Get OleIconToCursor function address
	if (NULL == (pfnITC = (pfnOleIconToCursor)GetProcAddress(hOleAut, "OleIconToCursor"))) {
		MessageBox(NULL, _T("Unable to get the OleIconToCursor address"), NULL, MB_ICONERROR);
		goto CleanUp;
	}

	// Register ExImage server within process
	if (FAILED(g_cServer.RegisterServer())) {
		MessageBox(NULL, _T("Unable to register PNG server"), NULL, MB_ICONERROR);
		goto CleanUp;
	}
	
	// Hijack OleIconToCursor
	if (!g_cHooker[2].Hook(pfnITC, OleIconToCursor_user))
		MessageBox(NULL, _T("Unable to intercept OleIconToCursor. Animated icons will be static"), NULL, MB_ICONERROR);

	// Hijack OleLoadPictureEx, OleLoadPicture
	if (g_cHooker[0].Hook(pfnEx, OleLoadPictureEx_user)) {
		if (g_cHooker[1].Hook(pfn, OleLoadPicture_user))
			g_bIsInitialized = TRUE;
		else
			g_cHooker[0].Unhook();
	}

	// If other dlls will use this functionaly set global env variable.
	// For example, if a EXE file uses static linked module and a dll uses
	// the dynamic link library we tells we already hook the functions.
	// In this case dll won't intercept it again
	if (g_bIsInitialized) {
		SetEnvironmentVariable(_T("VBPng"), _T("1"));
		InterlockedIncrement(&g_lCountOfUsers);
	}

CleanUp:

	if (!g_bIsInitialized) {

		// Revert
		g_cHooker[0].Unhook();
		g_cHooker[1].Unhook();
		g_cHooker[2].Unhook();

		if (bGdipInit)
			GdiplusShutdown(g_hGdipToken);

		g_cServer.UnregisterServer();
		
		if (bComInit)
			CoUninitialize();

	}

	return g_bIsInitialized;

}

VOID Uninitialize() {

	if (g_bIsInitialized && InterlockedDecrement(&g_lCountOfUsers) == 0) {

		g_cHooker[0].Unhook();
		g_cHooker[1].Unhook();
		g_cHooker[2].Unhook();

		g_cServer.UnregisterServer();
		
		GdiplusShutdown(g_hGdipToken);
		
		CoUninitialize();

		SetEnvironmentVariable(_T("VBPng"), _T("0"));

		g_bIsInitialized = FALSE;

	}

	return;

}

HRESULT CanUnloadNow() {
	// Check usage counter
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
	HRESULT hr			= S_OK;
	CPicture *cPic		= NULL;
	IStream *pStm		= NULL;
	LARGE_INTEGER liPos	= {0, 0};
	ULARGE_INTEGER uiOrigin;

	// Save stream position
	if (FAILED(hr = lpstream->Seek(liPos, STREAM_SEEK_CUR, &uiOrigin)))
		return hr;

	pfnOleLoadPictureEx pfn = (pfnOleLoadPictureEx)g_cHooker[0].GetThunkPtr();

	// Call the original one
	hr = pfn(lpstream, lSize, fRunmode, riid, xSizeDesired, ySizeDesired, dwFlags, lplpvObj);

	if (FAILED(hr)) {

		// Restore pointer

		liPos.QuadPart = uiOrigin.QuadPart;

		if (FAILED(hr = lpstream->Seek(liPos, STREAM_SEEK_SET, NULL)))
			return hr;

		cPic = new (std::nothrow) CPicture();

		if (!cPic) {
			hr = E_OUTOFMEMORY;
			goto CleanUp;
		}

		if (FAILED(hr = cPic->LoadFromStream(lpstream, !fRunmode, xSizeDesired, ySizeDesired, dwFlags)))
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

HICON __stdcall OleIconToCursor_user(
				HINSTANCE hinstExe,
				HICON     hIcon) {
	HICON hRet = NULL;
	CICOPicture *pAniIcon = g_cCursorsList.GetCursorFromHICON(hIcon);

	// Extract animated icon
	if (pAniIcon)
		hRet = pAniIcon->GetAniCopy();

	if (!hRet) {
		// If this isn't our animated icon or extracting failed then call default implementation
		pfnOleIconToCursor pfn = (pfnOleIconToCursor)g_cHooker[2].GetThunkPtr();
		hRet = pfn(hinstExe, hIcon);
	}

	return hRet;

}
