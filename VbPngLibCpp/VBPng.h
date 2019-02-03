
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

typedef HICON (__stdcall *pfnOleIconToCursor)(
  HINSTANCE hinstExe,
  HICON     hIcon
);

// Forward declaration
class CICOPicture;

//
// The list of all the animated cursors
// Since VB runtime uses OleIconToCursor API to handle cursors
// it can't process animated cursors because it uses CopyImage 
// function without LR_COPYFROMRESOURCE flag
// The solution is to intercept OleIconToCursor and check the 
// icon handle and if it's our ANI icon provide new copy of
// the animated cursor
class CANICursors {

private:

	CICOPicture** m_ppList;
	LONG m_lCount;

public:

	CANICursors();
	~CANICursors();

	BOOL AddCursor(CICOPicture*);
	BOOL RemoveCursor(CICOPicture*);
	CICOPicture* GetCursorFromHICON(HICON);

};

// Initialize the module
BOOL Initialize();

// Uninitialize module
VOID Uninitialize();

// Check if dll can be safely unloaded
HRESULT CanUnloadNow();

extern volatile ULONG g_lCountOfUsers;	// Number of users of DLL
extern volatile ULONG g_lCountOfObject;	// Count of active object
extern HMODULE g_hModule;				// Module handle
extern CANICursors g_cCursorsList;		// List of ANI cursors

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

HICON (__stdcall OleIconToCursor_user)(
				 HINSTANCE hinstExe,
				 HICON     hIcon);