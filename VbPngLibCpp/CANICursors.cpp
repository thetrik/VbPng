//
// CANICursor class implementation
//

#include "VBPng.h"
#include "CPicture.h"

// List of animated cursors
CANICursors g_cCursorsList;

CANICursors::CANICursors() {
	m_lCount = 0;
	m_ppList = NULL;
	return;
}

CANICursors::~CANICursors() {

	if (m_ppList)
		HeapFree(GetProcessHeap(), HEAP_NO_SERIALIZE, m_ppList);

	return;
}

BOOL CANICursors::AddCursor(CICOPicture* pCur) {
	PVOID pvList;

	if (GetCursorFromHICON((HICON)pCur->GetHandle()))
		return TRUE;

	if (m_lCount)
		pvList = HeapReAlloc(GetProcessHeap(), HEAP_NO_SERIALIZE, m_ppList, (m_lCount + 1) * sizeof(pCur));
	else
		pvList = HeapAlloc(GetProcessHeap(), HEAP_NO_SERIALIZE, sizeof(pCur));

	if (!pvList)
		return FALSE;

	m_ppList = (CICOPicture**)pvList;

	m_ppList[m_lCount] = pCur;

	m_lCount++;

	return TRUE;

}

BOOL CANICursors::RemoveCursor(CICOPicture* pCur) {
	INT i;

	for (i = 0; i < m_lCount; i++)
		if (pCur == m_ppList[i])
			break;

	if (i == m_lCount)
		return FALSE;
	else if (i < m_lCount - 1)
		memcpy(m_ppList + i, m_ppList + i + 1, (m_lCount - 1 - i) * sizeof(pCur));

	if (--m_lCount) {

		PVOID pvList = HeapReAlloc(GetProcessHeap(), HEAP_NO_SERIALIZE, m_ppList, m_lCount * sizeof(pCur));

		if (pvList)
			m_ppList = (CICOPicture**)pvList;

	} else {

		HeapFree(GetProcessHeap(), HEAP_NO_SERIALIZE, m_ppList);
		m_ppList = NULL;

	}

	return TRUE;

}

CICOPicture* CANICursors::GetCursorFromHICON(HICON hIcon) {

	for (INT i = 0; i < m_lCount; i++)
		if (m_ppList[i]->GetHandle() == hIcon)
			return m_ppList[i];

	return NULL;

}
