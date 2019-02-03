
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
HDC	CPicture::m_hDeskDc			= NULL;
ITypeLib* CPicture::m_pTypeLib	= NULL;

//
// CPictureInternal implementation
//
CPictureInternal::CPictureInternal() {

	m_hBitmap = NULL;
	m_Height = 0;
	m_Width = 0;
	m_pStmOriginalData = NULL;

	return;

}

CPictureInternal::~CPictureInternal() {
	
	if (m_pStmOriginalData)
		m_pStmOriginalData->Release();

	return;

}

LONG CPictureInternal::GetWidth() const {
	return m_Width;
}

LONG CPictureInternal::GetHeight() const {
	return m_Height;
}

HANDLE CPictureInternal::GetHandle() const {
	return m_hBitmap;
}

IStream* CPictureInternal::GetOriginalStream() const {
	return m_pStmOriginalData;
}

VOID CPictureInternal::FreeOriginalData() {

	if (m_pStmOriginalData)
		m_pStmOriginalData->Release();

	m_pStmOriginalData = NULL;

	return;
}

// 
// CPicture implementation
//

// Clear data
VOID CPicture::Clear() {
		
	if (m_pPicture)
		delete m_pPicture;

	m_pPicture = NULL;
	m_hCurDc = NULL;

	return;
}

CPicture::CPicture() {

	m_pPicture = NULL;
	m_hCurDc = NULL;
	m_pTypeInfo = NULL;
	m_RefCounter = 0;

	InterlockedIncrement((volatile long*)&g_lCountOfObject);

	return;

}

HRESULT CPicture::LoadFromStream(IStream *pStm, BOOL bKeepOriginal, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags) {
	DWORD dwSignature	= 0;
	HRESULT hr			= S_OK;
	CPictureInternal *pPic	= NULL;
	LARGE_INTEGER liPos;

	// Check the stream type by the file signature
	if (FAILED(hr = pStm->Read(&dwSignature, 4, NULL)))
		return hr;

	liPos.QuadPart = -4;

	pStm->Seek(liPos, STREAM_SEEK_CUR, NULL);

	if (dwSignature == 'GNP\x89') {
		// Create PNG
		pPic = new (std::nothrow) CPNGPicture;
	} else if (dwSignature == 0x10000 || dwSignature == 0x20000 || dwSignature == 'FFIR') {
		// Create ICO
		pPic = new (std::nothrow) CICOPicture();
	} else
		return CTL_E_INVALIDPICTURE;
	
	if (pPic)
		hr = pPic->Load(pStm, xSizeDesired, ySizeDesired, dwFlags, bKeepOriginal);
	else
		hr = E_OUTOFMEMORY;

	if (SUCCEEDED(hr)) {
		Clear();
		m_pPicture = pPic;
	} else {
		delete pPic;
	}

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

	return;
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

// 
// IPicture implementation
//
HRESULT STDMETHODCALLTYPE CPicture::get_Handle(OLE_HANDLE *pRet) {

	if (!m_pPicture)
		return E_UNEXPECTED;

	*pRet = (OLE_HANDLE)m_pPicture->GetHandle();
	return S_OK;

}

HRESULT STDMETHODCALLTYPE CPicture::get_hPal(OLE_HANDLE *pRet) {
	*pRet = NULL;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::get_Type(SHORT *pRet) {

	if (!m_pPicture)
		return E_UNEXPECTED;

	*pRet = m_pPicture->GetType();

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::get_Width(OLE_XSIZE_HIMETRIC *pRet) {

	if (!m_pPicture)
		return E_UNEXPECTED;

	*pRet = PixelsToHimetricX(m_pPicture->GetWidth());

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::get_Height(OLE_YSIZE_HIMETRIC *pRet) {

	if (!m_pPicture)
		return E_UNEXPECTED;

	*pRet = PixelsToHimetricY(m_pPicture->GetHeight());

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

	if (!m_pPicture)
		return E_FAIL;

	HANDLE hPrev = SelectObject(hDCIn, m_pPicture->GetHandle());

	if (phDCOut)
		*phDCOut = m_hCurDc;

	m_hCurDc = hDCIn;

	if (phBmpOut)
		*phBmpOut = (OLE_HANDLE)hPrev;

	return S_OK;

}

HRESULT STDMETHODCALLTYPE CPicture::get_KeepOriginalFormat(BOOL *bKeep) {
	*bKeep = (BOOL)m_pPicture->GetOriginalStream();
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::put_KeepOriginalFormat(BOOL bKeep) {

	if (bKeep && m_pPicture) 
		if (!m_pPicture->GetOriginalStream())
			return E_FAIL;
	else {

		if (m_pPicture->GetOriginalStream()) 
			m_pPicture->FreeOriginalData();

	}

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

	// Check if original data is available
	pStm = m_pPicture->GetOriginalStream();

	if (pStm && !fSaveMemCopy) {
		
		liPos.QuadPart = 0;

		if (FAILED(hr = pStm->Seek(liPos, STREAM_SEEK_END, &uiNewPos)))
			goto CleanUp;

		pBuffer = new (std::nothrow) BYTE [uiNewPos.LowPart];
		if (!pBuffer) {
			hr = E_OUTOFMEMORY;
			goto CleanUp;
		}

		if (FAILED(hr = pStm->Seek(liPos, STREAM_SEEK_SET, NULL)))
			goto CleanUp;

		if (FAILED(hr = pStm->Read(pBuffer, uiNewPos.LowPart, &liPos.LowPart)))
			goto CleanUp;

		if (FAILED(hr = pStream->Write(pBuffer, uiNewPos.LowPart, (ULONG*)pCbSize)))
			goto CleanUp;

	} else {

		hr = m_pPicture->Save(pStream, fSaveMemCopy, pCbSize);

	}

CleanUp:
	
	if (pBuffer)
		delete pBuffer;

	return hr;
}

HRESULT STDMETHODCALLTYPE CPicture::Render(HDC hDC, LONG x, LONG y, LONG cx, LONG cy,
										   OLE_XPOS_HIMETRIC xSrc, OLE_YPOS_HIMETRIC ySrc,
										   OLE_XSIZE_HIMETRIC cxSrc, OLE_YSIZE_HIMETRIC cySrc,
										   LPCRECT pRcWBounds) {
	if (!m_pPicture)
		return E_FAIL;

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

	// Prepare drawing in HIMETRIC units
	SetMapMode(m_hDeskDc, MM_ANISOTROPIC);

	SIZE szImgHimetric;

	szImgHimetric.cx = PixelsToHimetricX(m_pPicture->GetWidth());
	szImgHimetric.cy = PixelsToHimetricY(m_pPicture->GetHeight());

	SetWindowOrgEx(m_hDeskDc, 0, 0, NULL);
	SetWindowExtEx(m_hDeskDc, szImgHimetric.cx, szImgHimetric.cy, NULL);
	SetViewportOrgEx(m_hDeskDc, 0, 0, NULL);
	SetViewportExtEx(m_hDeskDc, m_pPicture->GetWidth(), m_pPicture->GetHeight(), NULL);

	HRESULT hr = m_pPicture->Draw(hDC, x, y, cx, cy, m_hDeskDc, xSrc, szImgHimetric.cy - ySrc, cxSrc, -cySrc);

	return hr;
}

//
// IPersistStream implementation
//
HRESULT STDMETHODCALLTYPE CPicture::GetClassID(CLSID *pclsid) {
	
	if (!pclsid)
		return E_INVALIDARG;

	// Return CLSID of server, it'll create picture objects
	*pclsid = CLSID_ExtPicture;

	return S_OK;

}

HRESULT STDMETHODCALLTYPE CPicture::IsDirty(void) {
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CPicture::Load(IStream *pStm) {
	return LoadFromStream(pStm, TRUE, 0, 0, 0);
}

HRESULT STDMETHODCALLTYPE CPicture::Save(IStream *pStm, BOOL fClearDirty) {
	HRESULT hr = S_OK;

	if (!m_pPicture)
		return E_UNEXPECTED;

	hr = SaveAsFile(pStm, m_pPicture->GetOriginalStream() == NULL, NULL);

	return hr;
}

HRESULT STDMETHODCALLTYPE CPicture::GetSizeMax(ULARGE_INTEGER *pSize) {
	return E_NOTIMPL;
}

// 
// IConnectionPointContainer implementation
//
HRESULT STDMETHODCALLTYPE CPicture::EnumConnectionPoints(IEnumConnectionPoints **ppEnum) {
	*ppEnum = FALSE;
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CPicture::FindConnectionPoint(const IID &iid,IConnectionPoint **ppC) {
	*ppC = NULL;
	return CONNECT_E_NOCONNECTION;
}

//
// IDispatch implementation
//
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