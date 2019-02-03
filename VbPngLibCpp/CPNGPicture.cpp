//
// CPNGPicture implementation
//

#include "CPicture.h"
#include <OleCtl.h>

CLSID CPNGPicture::m_ClsidPNGEncoder = CLSID_NULL;

CPNGPicture::CPNGPicture() {
}

CPNGPicture::~CPNGPicture() {
	
	if (m_hBitmap)
		DeleteObject(m_hBitmap);

}

//
// Create stream which contains only PNG data
//
HRESULT CPNGPicture::CreatePngStream(IStream *pStm, IStream **pOut) {
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

//
// Load PNG image from stream
//
HRESULT CPNGPicture::Load(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal) {
	HRESULT	hr			= S_OK;
	IStream *pPngStm	= NULL;

	// That parameters is intended only for icon/cursor
	if (xSizeDesired || ySizeDesired || dwFlags || m_hBitmap)
		return E_UNEXPECTED;

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

	m_Width = rc.Width;
	m_Height = rc.Height;
	m_hBitmap = hSection;

	if (bKeepOriginal) {
		m_pStmOriginalData = pPngStm;
		pPngStm->AddRef();
	} else
		m_pStmOriginalData = NULL;

CleanUp:

	if (tBitmapData.Scan0)
		cBitmap.UnlockBits(&tBitmapData);

	if (pPngStm) 
		pPngStm->Release();

	return hr;


}

//
// Save PNG image to stream
//
HRESULT CPNGPicture::Save(IStream *pStream, BOOL bKeepOriginal, LONG *pCbSize) {
	HRESULT hr		= S_OK;
	IStream *pStm	= NULL;
	PBYTE pBuffer	= NULL;
	LARGE_INTEGER liPos;
	STATSTG tStat;
	DIBSECTION dib;

	if (!GetObject(m_hBitmap, sizeof(DIBSECTION), &dib)) {
		hr = STG_E_CANTSAVE;
		goto CleanUp;
	}

	// Create temporary stream
	if (FAILED(hr = CreateStreamOnHGlobal(NULL, TRUE, &pStm))) 
		return hr;
	else {
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

//
// Draw picture
//
HRESULT CPNGPicture::Draw(HDC hDst, LONG dx, LONG dy, LONG dw, LONG dh, HDC hSrc, LONG sx, LONG sy, LONG sw, LONG sh) {
	HBITMAP hOldBmp = (HBITMAP)SelectObject(hSrc, m_hBitmap);
	HRESULT hr = S_OK;

	BLENDFUNCTION bf;

	bf.BlendOp = AC_SRC_OVER;
	bf.BlendFlags = 0;
	bf.SourceConstantAlpha = 255;
	bf.AlphaFormat = AC_SRC_ALPHA;

	if (!AlphaBlend(hDst, dx, dy, dw, dh, hSrc, sx, sy, sw, sh, bf))
		hr = E_FAIL;

	SelectObject(hSrc, hOldBmp);

	return hr;

}

SHORT CPNGPicture::GetType() const {
	return PICTYPE_BITMAP;
}
