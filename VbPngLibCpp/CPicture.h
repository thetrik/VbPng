
//
// Alternative PNG-implementation of StdPicture object
// It doesn't implement ConnectionPoints
// by The trick
// 2019

#include "stdafx.h"
#include <ocidl.h>

struct PNG_Chunk_Header {
	ULONG Length;
	ULONG Type;
};

class CPicture : IPicture, IPersistStream, IConnectionPointContainer, IDispatch {

private:

	ULONG		m_RefCounter;			// Reference counter
	HDC			m_hCurDc;				// Current selected DC
	HBITMAP		m_hBitmap;				// Current Dib image
	INT			m_Width,				// Image bounds
				m_Height;				// ...
	ITypeInfo	*m_pTypeInfo;			// TypeInfo object (for IDispatch)
	BOOL		m_bKeepOriginalFormat;	// Keep original data of file
	IStream		*m_pStmOriginalData;	// Original data stream

	static HDC		m_hDeskDc;			// Shared DC
	static ITypeLib	*m_pTypeLib;		// Type library (for IDispatch)
	static CLSID	m_ClsidPNGEncoder;	// CLSID png encoder 

	CPicture (const CPicture&);
	CPicture& operator = (const CPicture&);

	// Create stream which contains only the PNG data
	HRESULT CreatePngStream(IStream *pStm, IStream **pOut);
	// Load picture from stream
	HRESULT LoadFromStream(IStream *pStm);
	// Conversion functions
	LONG HimetricToPixelsX(LONG);
	LONG HimetricToPixelsY(LONG);
	LONG PixelsToHimetricX(LONG);
	LONG PixelsToHimetricY(LONG);
	// Clear data
	VOID Clear();

public:

	CPicture();
	~CPicture();

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