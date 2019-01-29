
// 
// StdPicture implementation 
// by The trick
// 2019


#include "CPicture.h"
#include "VBPng.h"
#include <intrin.h>
#include <shlwapi.h>
#include <OleCtl.h>

#pragma comment(lib, "Msimg32.lib")

// Static variables
HDC	CPicture::m_hDeskDc				= NULL;
ITypeLib* CPicture::m_pTypeLib		= NULL;
CLSID CPicture::m_ClsidPNGEncoder	= CLSID_NULL;

// Create stream which contains only PNG data
HRESULT CPicture::CreatePngStream(IStream *pStm, IStream **pOut) {
	DWORD dwPngSize	= 0;
	IStream *pRet	= NULL;
	HRESULT hr		= S_OK;
	PBYTE pBuffer	= NULL;
	ULARGE_INTEGER uOrigin;
	ULARGE_INTEGER uNewPos;
	ULARGE_INTEGER uSize;
	LARGE_INTEGER li;
	ULONG ulRead;
	PNG_Chunk_Header tChunkHdr;

	*pOut = NULL;

	li.QuadPart = 0;

	if (FAILED(hr = pStm->Seek(li, STREAM_SEEK_CUR, &uOrigin)))
		return hr;
	
	if (FAILED(hr = pStm->Read(&li, 8, &ulRead)))
		goto CleanUp;

	// ‰PNG/r/n
	if (li.QuadPart != 0xA1A0A0D474E5089) {
		hr = CTL_E_INVALIDPICTURE;
		goto CleanUp;
	}

	dwPngSize = 8;	// Signature size

	do {

		if (FAILED(hr = pStm->Read(&tChunkHdr, sizeof(tChunkHdr), &ulRead)))
			goto CleanUp;
		
		if (tChunkHdr.Length < 0) {
			hr = CTL_E_INVALIDPICTURE;
			goto CleanUp;
		}

		li.QuadPart = _byteswap_ulong(tChunkHdr.Length) + 4;	// data + crc

		if (FAILED(hr = pStm->Seek(li, STREAM_SEEK_CUR, &uNewPos)))
			goto CleanUp;
		
		dwPngSize += li.QuadPart + sizeof(tChunkHdr);

	} while(tChunkHdr.Type != 'DNEI');

	li.QuadPart = uOrigin.QuadPart;

	if (FAILED(hr = pStm->Seek(li, STREAM_SEEK_SET, &uOrigin)))
		goto CleanUp;

	// Create buffer for copying
	pBuffer = new (std::nothrow) BYTE[dwPngSize];

	if (!pBuffer) {
		hr = E_OUTOFMEMORY;
		goto CleanUp;
	}

	if (FAILED(hr = CreateStreamOnHGlobal(NULL, TRUE, &pRet)))
		goto CleanUp;

	uSize.QuadPart = dwPngSize;
	
	if (FAILED(hr = pStm->Read(pBuffer, dwPngSize, &ulRead))) {
		goto CleanUp;
	}

	if (FAILED(hr = pRet->Write(pBuffer, dwPngSize, &ulRead))) {
		goto CleanUp;
	}

	uSize.QuadPart = dwPngSize;

	if (FAILED(hr = pRet->SetSize(uSize))) {
		goto CleanUp;
	}

	li.QuadPart = 0;

	if (FAILED(hr = pRet->Seek(li, STREAM_SEEK_SET, &uNewPos))) {
		goto CleanUp;
	}

	*pOut = pRet;

CleanUp:

	if (FAILED(hr)) {

		if (pRet)
			pRet->Release();

	}

	if (pBuffer)
		delete pBuffer;

	return hr;

}

VOID CPicture::Clear() {
		
	if (m_hBitmap)
		DeleteObject(m_hBitmap);

	m_hBitmap = NULL;
	m_hCurDc = NULL;
	
	if (m_pStmOriginalData)
		m_pStmOriginalData->Release();

	m_pStmOriginalData = NULL;

	return;
}

CPicture::CPicture() {

	m_hBitmap = NULL;
	m_hCurDc = NULL;
	m_Height = 0;
	m_Width = 0;
	m_pTypeInfo = NULL;
	m_bKeepOriginalFormat = TRUE;
	m_pStmOriginalData = NULL;
	m_RefCounter = 0;

	InterlockedIncrement((volatile long*)&g_lCountOfObject);

	return;

}

HRESULT CPicture::LoadFromStream(IStream *pStm) {
	HRESULT	hr			= S_OK;
	IStream *pPngStm	= NULL;

	// Extract PNG part from stream
	if (FAILED(hr = CreatePngStream(pStm, &pPngStm))) 
		return hr;

	// Load bitmap from stream
	Bitmap cBitmap(pPngStm);

	if (cBitmap.GetLastStatus() != Ok)  {
		hr = CTL_E_INVALIDPICTURE;
		goto CleanUp;
	}

	// Convert to DIB
	Rect rc(0, 0, cBitmap.GetWidth(), cBitmap.GetHeight());
	BitmapData	tBitmapData;

	memset(&tBitmapData, 0, sizeof(BitmapData));

	if (cBitmap.LockBits(&rc, ImageLockModeRead, PixelFormat32bppPARGB, &tBitmapData) != Ok) {
		hr = E_OUTOFMEMORY;
		goto CleanUp;
	}

	BITMAPINFO bi;

	memset(&bi, 0, sizeof(BITMAPINFO));

	bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bi.bmiHeader.biBitCount = 32;
	bi.bmiHeader.biHeight = -rc.Height;
	bi.bmiHeader.biWidth = rc.Width;
	bi.bmiHeader.biPlanes = 1;

	HDC hdc = GetDC(NULL);
	PBYTE pbPixels = NULL;
	HBITMAP hSection = CreateDIBSection(hdc, &bi, DIB_RGB_COLORS, (void**)&pbPixels, NULL, 0);
	ReleaseDC(NULL, hdc);

	if (!hSection) {
		hr = E_OUTOFMEMORY;
		goto CleanUp;
	}

	memcpy(pbPixels, tBitmapData.Scan0, tBitmapData.Height * tBitmapData.Stride);

	// Set properties

	Clear();

	m_Width = rc.Width;
	m_Height = rc.Height;
	m_hBitmap = hSection;

	if (m_bKeepOriginalFormat) {
		m_pStmOriginalData = pPngStm;
		pPngStm->AddRef();
	}

CleanUp:

	if (tBitmapData.Scan0)
		cBitmap.UnlockBits(&tBitmapData);

	if (pPngStm) 
		pPngStm->Release();

	return hr;

}

CPicture::~CPicture() {
	
	if (InterlockedDecrement((volatile long*)&g_lCountOfObject) == 0) {
		// Clear global data

		if (m_pTypeLib)
			m_pTypeLib->Release();

		if (m_hDeskDc)
			DeleteDC(m_hDeskDc);

		m_hDeskDc = NULL;
		m_pTypeLib = NULL;

	}

	if (m_pTypeInfo)
		m_pTypeInfo->Release();

	m_pTypeInfo = NULL;

	Clear();
}

LONG CPicture::HimetricToPixelsX(LONG l) {

	if (!m_hDeskDc)
		m_hDeskDc = CreateCompatibleDC(NULL);

	return MulDiv(l, GetDeviceCaps(m_hDeskDc, LOGPIXELSX), 2540);

}

LONG CPicture::HimetricToPixelsY(LONG l) {

	if (!m_hDeskDc)
		m_hDeskDc = CreateCompatibleDC(NULL);

	return MulDiv(l, GetDeviceCaps(m_hDeskDc, LOGPIXELSY), 2540);

}

LONG CPicture::PixelsToHimetricX(LONG l) {

	if (!m_hDeskDc)
		m_hDeskDc = CreateCompatibleDC(NULL);

	return MulDiv(l, 2540, GetDeviceCaps(m_hDeskDc, LOGPIXELSX));

}

LONG CPicture::PixelsToHimetricY(LONG l) {

	if (!m_hDeskDc)
		m_hDeskDc = CreateCompatibleDC(NULL);

	return MulDiv(l, 2540, GetDeviceCaps(m_hDeskDc, LOGPIXELSY));

}

HRESULT STDMETHODCALLTYPE CPicture::QueryInterface(REFIID riid, void  **ppv) {
	HRESULT hr = S_OK;

	if (riid == IID_IUnknown) 
		*ppv = this;
	else if (riid == IID_IPicture)
		*ppv = (IPicture*)this;
	else if (riid == IID_IDispatch || riid == IID_IPictureDisp)
		*ppv = (IDispatch*)this;
	else if (riid == IID_IPersistStream)
		*ppv = (IPersistStream*)this;
	else if (riid == IID_IConnectionPointContainer)
		*ppv = (IConnectionPointContainer*)this;
	else
		hr = E_NOTIMPL;

	if (FAILED(hr))
		*ppv = NULL;
	else
		AddRef();

	return hr;

}

ULONG STDMETHODCALLTYPE CPicture::AddRef() {
	return ++m_RefCounter;
}

ULONG STDMETHODCALLTYPE CPicture::Release() {

	ULONG lResult = --m_RefCounter;

	if (lResult == 0)
		delete this;

	return lResult;

}

HRESULT STDMETHODCALLTYPE CPicture::get_Handle(OLE_HANDLE *pRet) {

	*pRet = (OLE_HANDLE)m_hBitmap;
	return S_OK;

}

HRESULT STDMETHODCALLTYPE CPicture::get_hPal(OLE_HANDLE *pRet) {
	*pRet = NULL;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::get_Type(SHORT *pRet) {
	*pRet = PICTYPE_BITMAP;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::get_Width(OLE_XSIZE_HIMETRIC *pRet) {
	*pRet = PixelsToHimetricX(m_Width);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::get_Height(OLE_YSIZE_HIMETRIC *pRet) {
	*pRet = PixelsToHimetricY(m_Height);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::set_hPal(OLE_HANDLE h) {
	return E_FAIL;
}

HRESULT STDMETHODCALLTYPE CPicture::get_CurDC(HDC *pHdc) {
	*pHdc = m_hCurDc;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::SelectPicture(HDC hDCIn, HDC *phDCOut, OLE_HANDLE *phBmpOut) {

	HANDLE hPrev = SelectObject(hDCIn, m_hBitmap);

	if (phDCOut)
		*phDCOut = m_hCurDc;

	m_hCurDc = hDCIn;

	if (phBmpOut)
		*phBmpOut = (OLE_HANDLE)hPrev;

	return S_OK;

}

HRESULT STDMETHODCALLTYPE CPicture::get_KeepOriginalFormat(BOOL *bKeep) {
	*bKeep = m_bKeepOriginalFormat;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::put_KeepOriginalFormat(BOOL bKeep) {

	if (bKeep && m_hBitmap) 
		if (!m_pStmOriginalData)
			return E_FAIL;
	else {

		if (m_pStmOriginalData) 
			m_pStmOriginalData->Release();

		m_pStmOriginalData = NULL;
	}

	m_bKeepOriginalFormat = bKeep;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::PictureChanged(void) {
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::get_Attributes(DWORD *dwAttr) {
	*dwAttr = PICTURE_TRANSPARENT;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::SaveAsFile(LPSTREAM pStream, BOOL fSaveMemCopy, LONG *pCbSize) {
	HRESULT hr		= S_OK;
	IStream	*pStm	= NULL;
	PBYTE pBuffer	= NULL;
	LARGE_INTEGER liPos;
	ULARGE_INTEGER uiNewPos;
	STATSTG tStat;
	DIBSECTION dib;

	if (m_bKeepOriginalFormat && fSaveMemCopy) {

		liPos.QuadPart = 0;

		if (FAILED(hr = m_pStmOriginalData->Seek(liPos, STREAM_SEEK_END, &uiNewPos)))
			goto CleanUp;

		pBuffer = new (std::nothrow) BYTE [uiNewPos.LowPart];
		if (!pBuffer) {
			hr = E_OUTOFMEMORY;
			goto CleanUp;
		}

		if (FAILED(hr = m_pStmOriginalData->Seek(liPos, STREAM_SEEK_SET, NULL)))
			goto CleanUp;

		if (FAILED(hr = m_pStmOriginalData->Read(pBuffer, uiNewPos.LowPart, &liPos.LowPart)))
			goto CleanUp;

		if (FAILED(hr = pStream->Write(pBuffer, uiNewPos.LowPart, (ULONG*)pCbSize)))
			goto CleanUp;

	} else {

		if (!GetObject(m_hBitmap, sizeof(DIBSECTION), &dib)) {
			hr = STG_E_CANTSAVE;
			goto CleanUp;
		}

		// Create temporary stream
		if (FAILED(hr = CreateStreamOnHGlobal(NULL, TRUE, &pStm))) 
			return hr;

		Bitmap cBitmap(dib.dsBm.bmWidth, dib.dsBm.bmHeight, dib.dsBm.bmWidthBytes, 
					   PixelFormat32bppPARGB, (PBYTE)dib.dsBm.bmBits);

		if (cBitmap.GetLastStatus() != Ok) {
			hr = STG_E_CANTSAVE;
			goto CleanUp;
		}

		if (memcmp(&m_ClsidPNGEncoder, &IID_NULL, sizeof(CLSID)) == 0) {

			// Init PNG encoder clsid
			PBYTE pCodecsBuffer			= NULL;
			ImageCodecInfo *pCodecInfo	= NULL;
			BOOL bFound					= FALSE;
			UINT uNumEncoders			= 0;
			UINT uSize					= 0;

			if (GetImageEncodersSize(&uNumEncoders, &uSize) != Ok) {
				hr = STG_E_CANTSAVE;
				goto CleanUp;
			}

			pCodecsBuffer = new (std::nothrow) BYTE[uSize];

			if (!pCodecsBuffer) {
				hr = E_OUTOFMEMORY;
				goto CleanUp;
			}

			pCodecInfo = (ImageCodecInfo*)pCodecsBuffer;

			if (GetImageEncoders(uNumEncoders, uSize, pCodecInfo) != Ok) {
				delete pCodecsBuffer;
				hr = STG_E_CANTSAVE;
				goto CleanUp;
			}

			for (UINT i = 0; i < uNumEncoders; i++, pCodecInfo++) {
				if (lstrcmp(pCodecInfo->MimeType, L"image/png") == 0) {
					m_ClsidPNGEncoder = pCodecInfo->Clsid;
					bFound = TRUE;
					break;
				}
			}

			delete pCodecsBuffer;

			if (!bFound) {
				hr = STG_E_CANTSAVE;
				goto CleanUp;
			}

		}

		if (cBitmap.Save(pStm, &m_ClsidPNGEncoder, NULL) != Ok) {
			hr = STG_E_CANTSAVE;
			goto CleanUp;
		}

		if (FAILED(hr = pStm->Stat(&tStat, STATFLAG_NONAME)))
			goto CleanUp;

		if (tStat.cbSize.QuadPart == 0 || tStat.cbSize.HighPart) {
			hr = E_FAIL;
			goto CleanUp;
		}

		pBuffer = new (std::nothrow) BYTE [tStat.cbSize.LowPart];
		if (!pBuffer) {
			hr = E_OUTOFMEMORY;
			goto CleanUp;
		}

		liPos.QuadPart = 0;

		if (FAILED(hr = pStm->Seek(liPos, STREAM_SEEK_SET, NULL)))
			goto CleanUp;

		if (FAILED(hr = pStm->Read(pBuffer, tStat.cbSize.LowPart, NULL)))
			goto CleanUp;

		if (FAILED(hr = pStream->Write(pBuffer, tStat.cbSize.LowPart, (ULONG*)pCbSize)))
			goto CleanUp;

	}

CleanUp:
	
	if (pBuffer)
		delete pBuffer;

	if (pStm)
		pStm->Release();

	return hr;
}

HRESULT STDMETHODCALLTYPE CPicture::Render(HDC hDC, LONG x, LONG y, LONG cx, LONG cy,
										   OLE_XPOS_HIMETRIC xSrc, OLE_YPOS_HIMETRIC ySrc,
										   OLE_XSIZE_HIMETRIC cxSrc, OLE_YSIZE_HIMETRIC cySrc,
										   LPCRECT pRcWBounds) {

	switch (GetObjectType(hDC)) {
	case OBJ_DC:
	case OBJ_ENHMETADC:
	case OBJ_ENHMETAFILE:
	case OBJ_MEMDC:
	case OBJ_METADC:
	case OBJ_METAFILE:
		break;
	default:
		return E_INVALIDARG;
	}

	if (!m_hDeskDc)
		if (NULL == (m_hDeskDc = CreateCompatibleDC(NULL)))
			return E_OUTOFMEMORY;

	SetMapMode(m_hDeskDc, MM_ANISOTROPIC);

	SIZE szImgHimetric;

	szImgHimetric.cx = PixelsToHimetricX(m_Width);
	szImgHimetric.cy = PixelsToHimetricY(m_Height);

	SetWindowOrgEx(m_hDeskDc, 0, 0, NULL);
	SetWindowExtEx(m_hDeskDc, szImgHimetric.cx, szImgHimetric.cy, NULL);
	SetViewportOrgEx(m_hDeskDc, 0, 0, NULL);
	SetViewportExtEx(m_hDeskDc, m_Width, m_Height, NULL);

	HBITMAP hOldBmp = (HBITMAP)SelectObject(m_hDeskDc, m_hBitmap);

	BLENDFUNCTION bf;

	bf.BlendOp = AC_SRC_OVER;
	bf.BlendFlags = 0;
	bf.SourceConstantAlpha = 255;
	bf.AlphaFormat = AC_SRC_ALPHA;

	AlphaBlend(hDC, x, y, cx, cy, m_hDeskDc, xSrc, szImgHimetric.cy - ySrc, cxSrc, -cySrc, bf);

	SelectObject(m_hDeskDc, hOldBmp);

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::GetClassID(CLSID *pclsid) {
	
	if (!pclsid)
		return E_INVALIDARG;

	// Return CLSID of server, it'll create picture objects
	*pclsid = CLSID_PngPicture;

	return S_OK;

}

HRESULT STDMETHODCALLTYPE CPicture::IsDirty(void) {
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::Load(IStream *pStm) {
	return LoadFromStream(pStm);
}

HRESULT STDMETHODCALLTYPE CPicture::Save(IStream *pStm, BOOL fClearDirty) {
	HRESULT hr = S_OK;

	hr = SaveAsFile(pStm, m_bKeepOriginalFormat, NULL);

	return hr;
}

HRESULT STDMETHODCALLTYPE CPicture::GetSizeMax(ULARGE_INTEGER *pSize) {
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CPicture::EnumConnectionPoints(IEnumConnectionPoints **ppEnum) {
	*ppEnum = FALSE;
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CPicture::FindConnectionPoint(const IID &iid,IConnectionPoint **ppC) {
	*ppC = NULL;
	return CONNECT_E_NOCONNECTION;
}

HRESULT STDMETHODCALLTYPE CPicture::GetTypeInfoCount(UINT *pctinfo) {
	*pctinfo = 1;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) {
	HRESULT hr = S_OK;

	*ppTInfo = NULL;

	if (!m_pTypeLib)
		if (FAILED(hr = LoadTypeLibEx(_T("stdole2.tlb"), REGKIND_NONE, &m_pTypeLib)))
			return hr;

	if (FAILED(hr = m_pTypeLib->GetTypeInfoOfGuid(IID_IPictureDisp, ppTInfo)))
		return hr;

	return hr;

}

HRESULT STDMETHODCALLTYPE CPicture::GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, 
												  LCID lcid, DISPID *rgDispId) {
	IDispatch *pThisDisp = (IDispatch*)this;
	HRESULT hr = S_OK;

	if (memcmp(&GUID_NULL, &riid, 0x10u))
		return DISP_E_UNKNOWNINTERFACE;

	if (!m_pTypeInfo)
		if (FAILED(hr = pThisDisp->GetTypeInfo(0, GetUserDefaultLCID(), &m_pTypeInfo)))
			return hr;

    return m_pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);

}

HRESULT STDMETHODCALLTYPE CPicture::Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, 
											DISPPARAMS *pDispParams, VARIANT *pVarResult, 
											EXCEPINFO *pExcepInfo, UINT *puArgErr) {
	HRESULT hr		= S_OK;
	HRESULT hrMeth	= S_OK;
	VARIANT *pVar	= NULL;
	VARIANT vParam;

	VariantInit(&vParam);

	if (memcmp(&GUID_NULL, &riid, 0x10u))
		return DISP_E_UNKNOWNNAME;

	if (wFlags & DISPATCH_PROPERTYGET) {

		if (pDispParams && pVarResult) {

			if (pDispParams->cArgs)
				return DISP_E_BADPARAMCOUNT;

			switch (dispIdMember) {
			case DISPID_PICT_HANDLE:

				pVarResult->vt = VT_I4;
				hrMeth = get_Handle((OLE_HANDLE*)&pVarResult->intVal);
				break;

			case DISPID_PICT_HPAL:

				pVarResult->vt = VT_I4;
				hrMeth = get_hPal((OLE_HANDLE*)&pVarResult->intVal);
				break;

			case DISPID_PICT_TYPE:

				pVarResult->vt = VT_I2;
				hrMeth = get_Type((SHORT*)&pVarResult->iVal);
				break;
			case DISPID_PICT_WIDTH:

				pVarResult->vt = VT_I4;
				hrMeth = get_Width((OLE_XSIZE_HIMETRIC*)&pVarResult->intVal);
				break;

			case DISPID_PICT_HEIGHT:

				pVarResult->vt = VT_I4;
				hrMeth = get_Height((OLE_XSIZE_HIMETRIC*)&pVarResult->intVal);
				break;

			default:

				hr = DISP_E_MEMBERNOTFOUND;
				break;
			}
			
		} else {
			hr = DISP_E_PARAMNOTOPTIONAL;
		}
	} else if (wFlags & DISPATCH_PROPERTYPUT) {

		if (!pDispParams)
			return DISP_E_PARAMNOTOPTIONAL;

		if (pDispParams->cArgs != 1)
			return DISP_E_BADPARAMCOUNT;

		if (!pDispParams->rgvarg)
			return DISP_E_PARAMNOTOPTIONAL;

		pVar = &vParam;

		switch (dispIdMember) {

		case DISPID_PICT_HPAL:

			if (pDispParams->rgvarg->vt == VT_I4)
				pVar = pDispParams->rgvarg;
			else
				if (SUCCEEDED(hr = VariantChangeType(pVar, pDispParams->rgvarg, 0, VT_I4))) 
					hrMeth = set_hPal(pVar->intVal);
			break;

		default:

			hr = DISP_E_MEMBERNOTFOUND;
			break;
		}
	} else if (wFlags & DISPATCH_METHOD) {

		if (!pDispParams)
			return DISP_E_PARAMNOTOPTIONAL;

		if (!pDispParams->rgvarg)
			return DISP_E_PARAMNOTOPTIONAL;

		if (pDispParams->cArgs != 10)
			return DISP_E_BADPARAMCOUNT;

		if (dispIdMember = DISPID_PICT_RENDER) {

			if (pDispParams->rgvarg->vt != VT_I4 && pDispParams->rgvarg->vt != VT_I2)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[1].vt != VT_I4)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[2].vt != VT_I4)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[3].vt != VT_I4)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[4].vt != VT_I4)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[5].vt != VT_I4)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[6].vt != VT_I4)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[7].vt != VT_I4)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[8].vt != VT_I4)
				return DISP_E_TYPEMISMATCH;
			if (pDispParams->rgvarg[9].vt != VT_I4)
			  return DISP_E_TYPEMISMATCH;

			hrMeth = Render((HDC)pDispParams->rgvarg[9].lVal,
							pDispParams->rgvarg[8].lVal,
							pDispParams->rgvarg[7].lVal,
							pDispParams->rgvarg[6].lVal,
							pDispParams->rgvarg[5].lVal,
							pDispParams->rgvarg[4].lVal,
							pDispParams->rgvarg[3].lVal,
							pDispParams->rgvarg[2].lVal,
							pDispParams->rgvarg[1].lVal,
							(LPRECT)pDispParams->rgvarg->lVal);

		} else
			return DISP_E_MEMBERNOTFOUND;

	} else {
		return DISP_E_MEMBERNOTFOUND;
	}

	if (FAILED(hrMeth)) {

        if ((hrMeth & 0x1FFF0000) == (FACILITY_CONTROL << 16) && pExcepInfo){
            memset(pExcepInfo, 0, sizeof(EXCEPINFO));
            pExcepInfo->scode = hrMeth;
            hr = DISP_E_EXCEPTION;
        } else
			hr = hrMeth;

	} else
		hr = hrMeth;

	return hr;
}