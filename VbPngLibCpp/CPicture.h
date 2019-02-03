
//
// Alternative PNG-implementation of StdPicture object
// It doesn't implement ConnectionPoints
// by The trick
// 2019

#include "stdafx.h"
#include <ocidl.h>
#include <initguid.h>

struct PNG_Chunk_Header {
	ULONG Length;
	ULONG Type;
};

typedef struct LOCALHEADER {
    WORD xHotSpot;
    WORD yHotSpot;
} LOCALHEADER;

typedef struct {
    WORD           idReserved;   // Reserved (must be 0)
    WORD           idType;       // Resource Type (1 for icons)
    WORD           idCount;      // How many images?
} ICONDIR, *LPICONDIR;

typedef struct {
    BYTE        bWidth;          // Width, in pixels, of the image
    BYTE        bHeight;         // Height, in pixels, of the image
    BYTE        bColorCount;     // Number of colors in image (0 if >=8bpp)
    BYTE        bReserved;       // Reserved ( must be 0)
    union {
		WORD    wPlanes;         // Color Planes
		WORD	xHotSpot;		 
	};
    union {
		WORD    wBitCount;		 // Bits per pixel
		WORD	yHotSpot;
	};
    DWORD       dwBytesInRes;    // How many bytes in this resource?
    DWORD       dwImageOffset;   // Where in the file is this image?
} ICONDIRENTRY, *LPICONDIRENTRY;


//
// Base class for image
// Each image type class should be inherited from that class
//
class CPictureInternal {

private:

	CPictureInternal (const CPictureInternal&);
	CPictureInternal& operator = (const CPictureInternal&);

protected:

	IStream *m_pStmOriginalData;	// Original file data
	LONG m_Width;					// Image width
	LONG m_Height;					// Image height
	HBITMAP m_hBitmap;				// Image handle

public:

	CPictureInternal();
	virtual ~CPictureInternal();

	LONG GetWidth() const;
	LONG GetHeight() const;
	HANDLE GetHandle() const ;
	IStream* GetOriginalStream() const;
	virtual SHORT GetType() const = 0;
	VOID FreeOriginalData();

	virtual HRESULT Save(IStream *pStm, BOOL bKeepOriginal, LONG *pCbSize) = 0;
	virtual HRESULT Load(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal) = 0;
	virtual HRESULT Draw(HDC, LONG, LONG, LONG, LONG, HDC, LONG, LONG, LONG, LONG) = 0;

};

// PNG picture CLSID {9D18EF1B-7C5F-4489-8CBC-0CDFDE9F3D46}
DEFINE_GUID(CLSID_ExtPicture, 
	0x9d18ef1b, 0x7c5f, 0x4489, 0x8c, 0xbc, 0xc, 0xdf, 0xde, 0x9f, 0x3d, 0x46);


// 
// Represents class factory for CPicture objects created by CoCreateInstance
// We need to use that class because a picture object can be created via
// CoCreateInstance function (for example if you read the picture property
// from PropertyBag).
//

class CPicturesServer : IClassFactory {

private:

	ULONG			m_RefCounter;			// Reference counter
	volatile ULONG	m_uLockCounter;			// Lock counter
	DWORD			m_dwCookie;				// Registration cookie
	BOOL			m_bIsInitialized;

	CPicturesServer (const CPicturesServer&);
	CPicturesServer& operator = (const CPicturesServer&);

public:

	CPicturesServer();
	~CPicturesServer();

	HRESULT RegisterServer();		// Register this server
	HRESULT UnregisterServer();		// Unregister

	// IUnknown implementation
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void **);
	ULONG STDMETHODCALLTYPE AddRef();
	ULONG STDMETHODCALLTYPE Release();

	// IClassFactory implementation
	HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *, const IID &,void **);
	HRESULT STDMETHODCALLTYPE LockServer(BOOL);

};

//
// Represents extended picture
//
class CPicture : IPicture, IPersistStream, IConnectionPointContainer, IDispatch {

private:

	ULONG		m_RefCounter;			// Reference counter
	HDC			m_hCurDc;				// Current selected DC
	ITypeInfo	*m_pTypeInfo;			// TypeInfo object (for IDispatch)
	CPictureInternal  *m_pPicture;			// Format-specific picture

	static HDC		m_hDeskDc;			// Shared DC
	static ITypeLib	*m_pTypeLib;		// Type library (for IDispatch)

	CPicture (const CPicture&);
	CPicture& operator = (const CPicture&);

	// Conversion functions
	static LONG HimetricToPixelsX(LONG);
	static LONG HimetricToPixelsY(LONG);
	static LONG PixelsToHimetricX(LONG);
	static LONG PixelsToHimetricY(LONG);

	// Clear data
	VOID Clear();

public:

	CPicture();
	~CPicture();

	// Load picture from stream
	HRESULT LoadFromStream(IStream *pStm, BOOL bKeepOriginal, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags);

	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void  **ppv);
	ULONG STDMETHODCALLTYPE AddRef();
	ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE get_Handle(OLE_HANDLE *);
    HRESULT STDMETHODCALLTYPE get_hPal(OLE_HANDLE *);
	HRESULT STDMETHODCALLTYPE get_Type(SHORT *);
	HRESULT STDMETHODCALLTYPE get_Width(OLE_XSIZE_HIMETRIC *);
	HRESULT STDMETHODCALLTYPE get_Height(OLE_YSIZE_HIMETRIC *);
	HRESULT STDMETHODCALLTYPE Render(HDC,LONG,LONG,LONG,LONG,OLE_XPOS_HIMETRIC,OLE_YPOS_HIMETRIC,OLE_XSIZE_HIMETRIC,OLE_YSIZE_HIMETRIC,LPCRECT);
	HRESULT STDMETHODCALLTYPE set_hPal(OLE_HANDLE);
	HRESULT STDMETHODCALLTYPE get_CurDC(HDC *);
	HRESULT STDMETHODCALLTYPE SelectPicture(HDC,HDC *,OLE_HANDLE *);
	HRESULT STDMETHODCALLTYPE get_KeepOriginalFormat(BOOL *);
	HRESULT STDMETHODCALLTYPE put_KeepOriginalFormat(BOOL);
	HRESULT STDMETHODCALLTYPE PictureChanged(void);
	HRESULT STDMETHODCALLTYPE SaveAsFile(LPSTREAM,BOOL,LONG *);
	HRESULT STDMETHODCALLTYPE get_Attributes(DWORD *);

	HRESULT STDMETHODCALLTYPE GetClassID(CLSID *);
	HRESULT STDMETHODCALLTYPE IsDirty(void);
	HRESULT STDMETHODCALLTYPE Load(IStream *);
	HRESULT STDMETHODCALLTYPE Save(IStream *,BOOL);
	HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER *);

	HRESULT STDMETHODCALLTYPE EnumConnectionPoints(IEnumConnectionPoints **);
	HRESULT STDMETHODCALLTYPE FindConnectionPoint(const IID &,IConnectionPoint **);

	HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *);
	HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT,LCID,ITypeInfo **);
	HRESULT STDMETHODCALLTYPE GetIDsOfNames(const IID &,LPOLESTR *,UINT,LCID,DISPID *);
	HRESULT STDMETHODCALLTYPE Invoke(DISPID,const IID &,LCID,WORD,DISPPARAMS *,VARIANT *,EXCEPINFO *,UINT *);

};

//
// Represents PNG image class
//
class CPNGPicture : public CPictureInternal {

	static CLSID	m_ClsidPNGEncoder;	// CLSID png encoder 

	// Create stream which contains only the PNG data
	HRESULT CreatePngStream(IStream *pStm, IStream **pOut);

	// CPictureInternal implementation
	SHORT GetType() const;
	HRESULT Save(IStream *pStm, BOOL bKeepOriginal, LONG *pCbSize);
	HRESULT Load(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal);
	HRESULT Draw(HDC, LONG, LONG, LONG, LONG, HDC, LONG, LONG, LONG, LONG);

	CPNGPicture (const CPNGPicture&);
	CPNGPicture& operator = (const CPNGPicture&);

public:

	CPNGPicture();
	~CPNGPicture();

};

//
// Represents icon/cursor class
//
class CICOPicture : public CPictureInternal {

	BOOL m_bIsAni;	// Is animated cursor or?

	// Create stream which contains only the ICO data
	HRESULT CreateIcoStream(IStream *pStm, IStream **pOut);
	// Create ANI stream
	HRESULT CreateAniStream(IStream *pStm, IStream **pOut);

	// Search for specified icon
    ICONDIRENTRY* GetIconEntry(ICONDIR *pIconDir, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags);
	// Get PNG image bounds
	SIZE GetPNGSize(PBYTE pBuffer, DWORD dwSize);
	// Load ICO/CUR
	HRESULT LoadIco(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal);
	// Load ANI
	HRESULT LoadAni(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal);

	// CPictureInternal implementation
	SHORT GetType() const;
	HRESULT Save(IStream *pStm, BOOL bKeepOriginal, LONG *pCbSize);
	HRESULT Load(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal);
	HRESULT Draw(HDC, LONG, LONG, LONG, LONG, HDC, LONG, LONG, LONG, LONG);

	CICOPicture (const CICOPicture&);
	CICOPicture& operator = (const CICOPicture&);

public:

	// Create animated cursor copy
	HICON GetAniCopy();

	CICOPicture();
	~CICOPicture();


};