//
// CICOPicture implementation
//

#include "CPicture.h"
#include <OleCtl.h>
#include "VBPng.h"

CICOPicture::CICOPicture() {
	m_bIsAni = FALSE;
	return;
}

CICOPicture::~CICOPicture() {
	if (m_bIsAni)
		g_cCursorsList.RemoveCursor(this);
}

//
// Get size of PNG icon
//
SIZE CICOPicture::GetPNGSize(PBYTE pBuffer, DWORD dwSize) {
	SIZE szRet = {0, 0};
	PNG_Chunk_Header *pChunkHdr;
	DWORD dwLen;

	if (dwSize < 8 || (*(LONGLONG*)pBuffer != 0xA1A0A0D474E5089))
		return szRet;

	pBuffer += 8; dwSize -= 8;
	
	do {

		if (dwSize < 8)
			break;

		pChunkHdr = (PNG_Chunk_Header*)pBuffer;
		
		dwLen = _byteswap_ulong(pChunkHdr->Length);

		if (pChunkHdr->Type == 'RDHI' && dwLen >= 8) {

			szRet.cx = _byteswap_ulong(*(ULONG*)(pBuffer + 8));
			szRet.cy = _byteswap_ulong(*(ULONG*)(pBuffer + 12));

			return szRet;
		}
		
		pBuffer += 12 + dwLen;

	} while(dwSize);

	return szRet;

}

//
// Create stream which contains only ICO/CUR data
//
HRESULT CICOPicture::CreateIcoStream(IStream *pStm, IStream **pOut) {
	HRESULT hr		= S_OK;
	PBYTE pBuffer	= NULL;
	DWORD dwMaxSize = 0;
	IStream *pRet	= NULL;
	ULARGE_INTEGER uiOrigin, uiSize;
	ULONG ulRead;
	LARGE_INTEGER liPos;
	ICONDIR tDir;
	ICONDIRENTRY tEntry;

	liPos.QuadPart = 0;

	// Save original stream position
	if (FAILED(hr = pStm->Seek(liPos, STREAM_SEEK_CUR, &uiOrigin)))
		return hr;

	if (FAILED(hr = pStm->Read(&tDir, sizeof(tDir), &ulRead)))
		return hr;

	// Header validation
	if (tDir.idReserved || !tDir.idCount || (tDir.idType != 1 && tDir.idType != 2)) {
		hr = CTL_E_INVALIDPICTURE;
		goto CleanUp;
	}

	while (tDir.idCount--) {

		if (FAILED(hr = pStm->Read(&tEntry, sizeof(tEntry), &ulRead)))
			goto CleanUp;

		if (tEntry.bReserved || !tEntry.dwBytesInRes) {
			hr = CTL_E_INVALIDPICTURE;
			goto CleanUp;
		}

		if (tDir.idType == 1) {
			// ICON
			if (tEntry.wPlanes > 1 || tEntry.wBitCount == 0 || tEntry.wBitCount > 32) {
				hr = CTL_E_INVALIDPICTURE;
				goto CleanUp;
			}
		}

		// Check file data position and size
		if (tEntry.dwImageOffset + tEntry.dwBytesInRes > dwMaxSize)
			dwMaxSize = tEntry.dwImageOffset + tEntry.dwBytesInRes;

	}

	// Set initial position
	liPos.QuadPart = uiOrigin.QuadPart;

	if (FAILED(hr = pStm->Seek(liPos, STREAM_SEEK_SET, NULL)))
		goto CleanUp;

	// Create buffer for copying
	pBuffer = new (std::nothrow) BYTE[dwMaxSize];

	if (!pBuffer) {
		hr = E_OUTOFMEMORY;
		goto CleanUp;
	}

	if (FAILED(hr = CreateStreamOnHGlobal(NULL, TRUE, &pRet)))
		goto CleanUp;

	if (FAILED(hr = pStm->Read(pBuffer, dwMaxSize, &ulRead)))
		goto CleanUp;

	if (FAILED(hr = pRet->Write(pBuffer, dwMaxSize, &ulRead))) 
		goto CleanUp;

	uiSize.QuadPart = dwMaxSize;

	if (FAILED(hr = pRet->SetSize(uiSize)))
		goto CleanUp;

	liPos.QuadPart = 0;

	if (FAILED(hr = pRet->Seek(liPos, STREAM_SEEK_SET, NULL)))
		goto CleanUp;

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
// Create stream which contains only ANI data
//
HRESULT CICOPicture::CreateAniStream(IStream *pStm, IStream **pOut) {
	HRESULT hr		= S_OK;
	IStream *pRet	= NULL;
	PBYTE pBuffer	= NULL;
	ULARGE_INTEGER uiOrigin, uiSize;
	ULONG ulRead;
	LARGE_INTEGER liPos;
	DWORD dwSignature,
		  dwSize;

	liPos.QuadPart = 0;

	// Save original stream position
	if (FAILED(hr = pStm->Seek(liPos, STREAM_SEEK_CUR, &uiOrigin)))
		return hr;

	if (FAILED(hr = pStm->Read(&dwSignature, 4, &ulRead)))
		return hr;

	if (dwSignature != 'FFIR') {
		hr = CTL_E_INVALIDPICTURE;
		goto CleanUp;
	}

	// Extract size of data
	if (FAILED(hr = pStm->Read(&dwSize, 4, &ulRead)))
		goto CleanUp;
	
	if (dwSize <= 8) {
		hr = CTL_E_INVALIDPICTURE;
		goto CleanUp;
	}

	pBuffer = new (std::nothrow) BYTE[dwSize + 8];

	if (!pBuffer) {
		hr = E_OUTOFMEMORY;
		goto CleanUp;
	}
	
	// Set initial position
	liPos.QuadPart = uiOrigin.QuadPart;

	if (FAILED(hr = pStm->Seek(liPos, STREAM_SEEK_SET, NULL)))
		goto CleanUp;

	if (FAILED(hr = CreateStreamOnHGlobal(NULL, TRUE, &pRet)))
		goto CleanUp;

	if (FAILED(hr = pStm->Read(pBuffer, dwSize + 8, &ulRead)))
		goto CleanUp;

	if (FAILED(hr = pRet->Write(pBuffer, dwSize + 8, &ulRead))) 
		goto CleanUp;

	uiSize.QuadPart = dwSize + 8;

	if (FAILED(hr = pRet->SetSize(uiSize)))
		goto CleanUp;

	liPos.QuadPart = 0;

	if (FAILED(hr = pRet->Seek(liPos, STREAM_SEEK_SET, NULL)))
		goto CleanUp;

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
// Get icon entry
//
ICONDIRENTRY* CICOPicture::GetIconEntry(ICONDIR *pDir, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags) {
	ICONDIRENTRY *pEntry	= NULL,
				 *pRet		= NULL;
	DWORD dwThreshold		= 0,
		  dwMin				= 0xffffffff;
	PBYTE pbImageData		= NULL;
	SIZE szImage;
	WORD wCount;

	if (dwFlags & LR_DEFAULTSIZE) {
		if (!xSizeDesired & !ySizeDesired) {
			xSizeDesired = GetSystemMetrics(SM_CXICON);
			ySizeDesired = GetSystemMetrics(SM_CYICON);
		}
	}

	wCount = pDir->idCount;
	pEntry = (ICONDIRENTRY*)((PBYTE)pDir + sizeof(ICONDIR));

	while (wCount--) {

		if (!xSizeDesired & !ySizeDesired) {
			pRet = pEntry;
			break;
		}

		if (!pEntry->bWidth && !pEntry->bHeight) {

			// >255 (PNG)
			
			pbImageData = (PBYTE)pDir + pEntry->dwImageOffset;

			if (*(DWORD*)pbImageData == 'GNP\x89') {
				// PNG icon
				szImage = GetPNGSize(pbImageData, pEntry->dwBytesInRes);

			} else {
				BITMAPINFOHEADER *pHeader = (BITMAPINFOHEADER*)pbImageData;

				szImage.cx = pHeader->biWidth;
				szImage.cy = pHeader->biHeight / 2;

			}
		} else {
			szImage.cx = pEntry->bWidth;
			szImage.cy = pEntry->bHeight;
		}

		if (szImage.cx == xSizeDesired && szImage.cy == ySizeDesired) {
			pRet = pEntry;
			break;
		} else {

			dwThreshold = abs((LONG)xSizeDesired - szImage.cx) + abs((LONG)ySizeDesired - szImage.cy);

			if (dwThreshold < dwMin) {
				dwMin = dwThreshold;
				pRet = pEntry;
			}
		}

		pEntry++;

	}

	return pRet;

}

//
// Load ANI
//
HRESULT CICOPicture::LoadAni(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal) {
	HRESULT	hr			= S_OK;
	IStream *pAniStm	= NULL;
	HGLOBAL hMem		= NULL;
	PBYTE pbData		= NULL;
	HICON hCursor		= NULL;
	HBITMAP hBitmap		= NULL;
	STATSTG tStat;
	ICONINFO tCursorInfo;
	BITMAP tBitmap;

	// If already initialized
	if (m_hBitmap)
		return E_UNEXPECTED;

	// Extract ANI part from stream
	if (FAILED(hr = CreateAniStream(pStm, &pAniStm))) 
		return hr;

	if (FAILED(hr = pAniStm->Stat(&tStat, STATFLAG_NONAME)))
		goto CleanUp;

	if (FAILED(hr = GetHGlobalFromStream(pAniStm, &hMem)))
		goto CleanUp;

	if (NULL == (pbData = (PBYTE)GlobalLock(hMem))) {
		hr = E_OUTOFMEMORY;
		goto CleanUp;
	}

	if (xSizeDesired == LP_DEFAULT && ySizeDesired == LP_DEFAULT)
		dwFlags |= LR_DEFAULTSIZE;

	if (NULL == (hCursor = CreateIconFromResourceEx(pbData, tStat.cbSize.LowPart, FALSE, 0x00030000, 
												  xSizeDesired, ySizeDesired, dwFlags))) {
		hr = CTL_E_INVALIDPICTURE;
		goto CleanUp;
	}

	if (!GetIconInfo(hCursor, &tCursorInfo)) {
		hr = E_FAIL;
		goto CleanUp;
	}

	hBitmap = tCursorInfo.hbmColor ? tCursorInfo.hbmColor : tCursorInfo.hbmMask;

	if (!GetObject(hBitmap, sizeof(tBitmap), &tBitmap)) {
		hr = E_FAIL;
		goto CleanUp;
	}

	m_Width = tBitmap.bmWidth;
	m_Height = tBitmap.bmHeight;
	m_hBitmap = (HBITMAP)hCursor;

	// Save content always because we need to create copies
	m_pStmOriginalData = pAniStm;
	m_pStmOriginalData->AddRef();

CleanUp:

	if (pbData)
		GlobalUnlock(pbData);

	if (pAniStm)
		pAniStm->Release();

	return hr;

}

//
// Load ICO/CUR
//
HRESULT CICOPicture::LoadIco(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal) {
	HRESULT	hr				= S_OK;
	IStream *pIcoStm		= NULL;
	HGLOBAL hMem			= NULL;
	HICON hIcon				= NULL;
	ICONDIR *pIconDir		= NULL;
	ICONDIRENTRY *pEntry	= NULL;
	BOOL bIsIcon			= FALSE;
	PBYTE pbResData			= NULL;
	DWORD dwResSize			= 0;
	HBITMAP hBitmap;
	BITMAP tBitmap;
	ICONINFO tIconInfo;

	// If already initialized
	if (m_hBitmap)
		return E_UNEXPECTED;

	// Extract ICO part from stream
	if (FAILED(hr = CreateIcoStream(pStm, &pIcoStm))) 
		return hr;

	if (FAILED(hr = GetHGlobalFromStream(pIcoStm, &hMem)))
		goto CleanUp;

	if (NULL == (pIconDir = (ICONDIR*)GlobalLock(hMem))) {
		hr = E_OUTOFMEMORY;
		goto CleanUp;
	}

	if (xSizeDesired == LP_DEFAULT && ySizeDesired == LP_DEFAULT)
		dwFlags |= LR_DEFAULTSIZE;

	if (NULL == (pEntry = GetIconEntry(pIconDir, xSizeDesired, ySizeDesired, dwFlags))) {
		hr = E_FAIL;
		goto CleanUp;
	}
	
	// Check if cursor
	bIsIcon = pIconDir->idType == 2 ? FALSE : TRUE;

	// Because a cursor begins with LOCALHEADER structure we need to fill it
	if (!bIsIcon) {
		
		// Allocate buffer
		pbResData = new (std::nothrow) BYTE[sizeof(LOCALHEADER) + pEntry->dwBytesInRes];

		if (!pbResData) {
			hr = E_OUTOFMEMORY;
			goto CleanUp;
		}

		LOCALHEADER *pHeader = (LOCALHEADER*) pbResData;

		pHeader->xHotSpot = pEntry->xHotSpot;
		pHeader->yHotSpot = pEntry->yHotSpot;

		memcpy(pbResData + sizeof(LOCALHEADER), (PBYTE)pIconDir + pEntry->dwImageOffset, pEntry->dwBytesInRes);

		dwResSize = pEntry->dwBytesInRes + sizeof(LOCALHEADER);

	} else {
		pbResData = (PBYTE)pIconDir + pEntry->dwImageOffset;
		dwResSize = pEntry->dwBytesInRes;
	}


	if (NULL == (hIcon = CreateIconFromResourceEx(pbResData, dwResSize, bIsIcon, 0x00030000, 
												  xSizeDesired, ySizeDesired, dwFlags))) {
		hr = CTL_E_INVALIDPICTURE;
		goto CleanUp;
	}

	if (!GetIconInfo(hIcon, &tIconInfo)) {
		hr = E_FAIL;
		goto CleanUp;
	}

	hBitmap = tIconInfo.hbmColor ? tIconInfo.hbmColor : tIconInfo.hbmMask;

	if (!GetObject(hBitmap, sizeof(tBitmap), &tBitmap)) {
		hr = E_FAIL;
		goto CleanUp;
	}

	m_Width = tBitmap.bmWidth;
	m_Height = tBitmap.bmHeight;
	m_hBitmap = (HBITMAP)hIcon;

	if (bKeepOriginal) {
		m_pStmOriginalData = pIcoStm;
		pIcoStm->AddRef();
	} else
		m_pStmOriginalData = NULL;

CleanUp:

	if (!bIsIcon && pbResData)
		delete pbResData;

	if (pIconDir)
		GlobalUnlock(pIconDir);

	if (pIcoStm)
		pIcoStm->Release();

	return hr;

}

//
// Load ICO/CUR/ANI image from stream
//
HRESULT CICOPicture::Load(IStream *pStm, DWORD xSizeDesired, DWORD ySizeDesired, DWORD dwFlags, BOOL bKeepOriginal) {
	DWORD dwSignature	= 0;
	HRESULT hr			= S_OK;
	LARGE_INTEGER liPos;

	// Check the type
	if (FAILED(hr = pStm->Read(&dwSignature, 4, NULL)))
		return hr;

	liPos.QuadPart = -4;

	pStm->Seek(liPos, STREAM_SEEK_CUR, NULL);

	if (dwSignature == 'FFIR') {
		// ANI
		hr = LoadAni(pStm, xSizeDesired, ySizeDesired, dwFlags, bKeepOriginal);

		// Add to list of animated cursors
		if (SUCCEEDED(hr)) {
			g_cCursorsList.AddCursor(this);
			m_bIsAni = TRUE;
		}

	} else {
		// ICO/CUR
		hr = LoadIco(pStm, xSizeDesired, ySizeDesired, dwFlags, bKeepOriginal);
	}

	return hr;

}

HRESULT CICOPicture::Save(IStream *pStream, BOOL bKeepOriginal, LONG *pCbSize) {
	return E_NOTIMPL;
}

//
// Draw picture
//
HRESULT CICOPicture::Draw(HDC hDst, LONG dx, LONG dy, LONG dw, LONG dh, HDC hSrc, LONG sx, LONG sy, LONG sw, LONG sh) {
	HRESULT hr = S_OK;

	if (!DrawIconEx(hDst, dx, dy, (HICON)m_hBitmap, dw, dh, 0, NULL, DI_NORMAL))
		return E_FAIL;

	return hr;

}

SHORT CICOPicture::GetType() const {
	return PICTYPE_ICON;
}

// 
// Extract animated icon copy
//
HICON CICOPicture::GetAniCopy() {
	HGLOBAL hMem	= NULL;
	PBYTE pbData	= NULL;
	HRESULT hr		= S_OK;
	HICON hIcon		= NULL;
	STATSTG tStat;

	if (!m_bIsAni || !m_pStmOriginalData)
		return NULL;

	if (FAILED(hr = m_pStmOriginalData->Stat(&tStat, STATFLAG_NONAME)))
		return NULL;

	if (FAILED(hr = GetHGlobalFromStream(m_pStmOriginalData, &hMem)))
		return NULL;

	if (NULL == (pbData = (PBYTE)GlobalLock(hMem)))
		return NULL;

	if (NULL == (hIcon = CreateIconFromResourceEx(pbData, tStat.cbSize.LowPart, FALSE, 0x00030000, m_Width, m_Height, 0)))
		goto CleanUp;

CleanUp:

	if (pbData)
		GlobalUnlock(pbData);

	return hIcon;

}