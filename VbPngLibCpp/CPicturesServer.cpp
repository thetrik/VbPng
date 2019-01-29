// 
// COM PNG pictures server implementation 
// by The trick
// 2019

#include "CPicture.h"

CPicturesServer::CPicturesServer() {
	m_RefCounter = m_uLockCounter = 0;
	return;
}

CPicturesServer::~CPicturesServer() {
	UnregisterServer();
	return;
}

// Register that private factory to have ability create PNG images via CoCreateInstance
HRESULT CPicturesServer::RegisterServer() {
	HRESULT	hr = S_OK;

	// Register this server within process
	if (FAILED(hr = CoRegisterClassObject(CLSID_PngPicture, this, CLSCTX_INPROC_SERVER, 
									 REGCLS_MULTIPLEUSE, &m_dwCookie))) 
		goto CleanUp;

	m_bIsInitialized = TRUE;

CleanUp:

	return hr;

}

HRESULT CPicturesServer::UnregisterServer() {

	if (m_bIsInitialized)
		return CoRevokeClassObject(m_dwCookie);

	m_bIsInitialized = FALSE;

	return E_UNEXPECTED;

}

HRESULT STDMETHODCALLTYPE CPicturesServer::QueryInterface(REFIID riid, void  **ppv) {
	HRESULT hr = S_OK;

	if (riid == IID_IUnknown) 
		*ppv = this;
	else if (riid == IID_IClassFactory)
		*ppv = (IClassFactory*)this;
	else
		hr = E_NOTIMPL;

	if (FAILED(hr))
		*ppv = NULL;
	else
		AddRef();

	return hr;

}

ULONG STDMETHODCALLTYPE CPicturesServer::AddRef() {
	return ++m_RefCounter;
}

ULONG STDMETHODCALLTYPE CPicturesServer::Release() {

	ULONG lResult = --m_RefCounter;

	return lResult;

}

// Create uninitialized picture
HRESULT STDMETHODCALLTYPE CPicturesServer::CreateInstance(IUnknown *pUnk, const IID &riid, void **ppvOut) {
	HRESULT hr		= S_OK;
	CPicture *cPic	= NULL;

	if (pUnk)
		return CLASS_E_NOAGGREGATION;

	cPic = new (std::nothrow) CPicture();

	if (!cPic) {
		hr = E_OUTOFMEMORY;
		goto CleanUp;
	}

	if (FAILED(hr = cPic->QueryInterface(riid, ppvOut)))
		goto CleanUp;

CleanUp:

	if (FAILED(hr))
		if (cPic)
			delete cPic;

	return hr;

}

HRESULT STDMETHODCALLTYPE CPicturesServer::LockServer(BOOL fLock) {

	if (fLock)
		InterlockedIncrement((volatile long*)&m_uLockCounter);
	else 
		if (!m_uLockCounter)
			return E_FAIL;
		else
			InterlockedDecrement((volatile long*)&m_uLockCounter);

	return S_OK;
}