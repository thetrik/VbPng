// 
// CHooker class implementation
// by The trick
// 2019


#include "CHooker.h"

HANDLE CHooker::s_hHeap					= NULL;	
volatile ULONG CHooker::s_uHookersCount	= 0;

// Allocate memory for code
PVOID CHooker::AllocMem(ULONG uSize) {

	if (!s_hHeap)
		s_hHeap = HeapCreate(HEAP_NO_SERIALIZE | HEAP_CREATE_ENABLE_EXECUTE, 0, 0);

	if (!s_hHeap)
		return NULL;

	return HeapAlloc(s_hHeap, HEAP_NO_SERIALIZE, uSize);

}

// Free memory
VOID CHooker::FreeMem(PVOID pvMem) {

	if (!s_hHeap)
		return;

	HeapFree(s_hHeap, HEAP_NO_SERIALIZE, pvMem);

	return;

}

// 
// Hook the procedure. It ensures to call pfnHandler function insead original one
//
BOOL CHooker::Hook(PVOID pfnHook, PVOID pfnHandler) {
	LONG lInstructionSize	= 0,
		 lHookChainSize		= 0,
		 lIndex				= 0,
		 lOpcodesCount		= 0,
		 lRelOffset			= 0,
		 lFixupValue		= 0,
		 lSizeDifference	= 0;
	PBYTE pbCode			= NULL,
		  pCurCmd			= NULL;
	PBYTE pbOpcodes[80],
		  pbInstructions[80];
	BYTE bOpcode;

	if (m_pfnHook)
		Unhook();

	pbInstructions[lIndex] = (PBYTE)pfnHook;

	do {

		if (0 == (lInstructionSize = SizeOfCode(pbInstructions[lIndex], &pbOpcodes[lIndex]))) {
			MessageBox(NULL, _T("Unable to get the instruction size"), NULL, MB_ICONERROR);
			return FALSE;
		}

		pbInstructions[++lIndex] = pbInstructions[lIndex - 1] + lInstructionSize;
		lHookChainSize = lHookChainSize + lInstructionSize;

	} while(lHookChainSize < 5);

	// Set the buffer size to ensure translation all the opcodes at least from the short form JZ SHORT etc.
	// Allocate the new handler


	if (NULL == (pbCode = (PBYTE)AllocMem((lHookChainSize + 5) * 8))) {
		MessageBox(NULL, _T("Unable to alloc the memory for the handler"), NULL, MB_ICONERROR);
		return FALSE;
	}

	// Save the original code
	m_pbOriginalCode = pbCode + (lHookChainSize + 5) * 4;

	memcpy(m_pbOriginalCode, pfnHook, lHookChainSize);

	lFixupValue = (PBYTE)pfnHook - pbCode;
	pCurCmd = pbCode;
	lOpcodesCount = lIndex;

	for (lIndex = 0; lIndex < lOpcodesCount; lIndex++) {

		lInstructionSize = pbInstructions[lIndex + 1] - pbInstructions[lIndex];

		memcpy(pCurCmd, pbInstructions[lIndex], lInstructionSize);

		// Fixup the offset
		if (IsRelativeCmd(pbOpcodes[lIndex])) {
			
			bOpcode = *pbOpcodes[lIndex];

			// Skip prefixes
			pCurCmd += (pbOpcodes[lIndex] - pbInstructions[lIndex]);

			if (bOpcode == 0xe8 || bOpcode == 0xe9) {

				// JMP/CALL LONG

				lRelOffset = *(LONG*)++pCurCmd;

				*(LONG*)pCurCmd = lRelOffset + lFixupValue - lSizeDifference;

				pCurCmd += 4;

			} else if ((bOpcode == 0x0f && (pbOpcodes[lIndex][1] >= 0x80 && pbOpcodes[lIndex][1] <= 0x8f))) {

				// JXXX LONG

				++pCurCmd;
				lRelOffset = *(LONG*)++pCurCmd;

				*(LONG*)pCurCmd = lRelOffset + lFixupValue - lSizeDifference;

				pCurCmd += 4;

			} else if ((bOpcode >= 0x70 && bOpcode <= 0x7f)) {

				// JXXX SHORT

				lRelOffset = *(PCHAR)(pCurCmd + 1);

				bOpcode = bOpcode + 0x10;

				// Transform to LONG

				*pCurCmd++ = 0x0f;
				*pCurCmd++ = bOpcode;
				*(LONG*)pCurCmd = lRelOffset + lFixupValue - 4 - lSizeDifference;

				lSizeDifference += 4;

				pCurCmd += 4;

			} else if (bOpcode == 0xeb) {

				// JMP SHORT

				lRelOffset = *(PCHAR)(pCurCmd + 1);

				// Transform to LONG

				*pCurCmd++ = 0xe9;
				*(LONG*)pCurCmd = lRelOffset + lFixupValue - 3 - lSizeDifference;

				lSizeDifference += 3;

				pCurCmd += 4;

			}

		} else {
			pCurCmd += lInstructionSize;
		}

	}

	// Add JMP to original

	*pCurCmd++ = 0xe9;
	*(LONG*)pCurCmd = (PBYTE)pfnHook + lHookChainSize - (pCurCmd + 4);

	// Add JMP to handler
	DWORD dwOldProtect;

	if (!VirtualProtect(pfnHook, lHookChainSize, PAGE_EXECUTE_READWRITE, &dwOldProtect)) {
		MessageBox(NULL, _T("Unable to change pages protection"), NULL, MB_ICONERROR);
		return FALSE;
	}

	pCurCmd = (PBYTE)pfnHook;

	*pCurCmd++ = 0xe9;
	*(LONG*)pCurCmd = (PBYTE)pfnHandler - (pCurCmd + 4);

	VirtualProtect(pfnHook, lHookChainSize, dwOldProtect, &dwOldProtect);

	m_pbHandler = pbCode;
	m_pfnHook = pfnHook;
	m_ReplacedSize = lHookChainSize;

	return TRUE;

}

CHooker::CHooker() {

	m_pbOriginalCode = NULL;
	m_pbHandler = NULL;
	
	m_pfnHook = NULL;
	m_ReplacedSize = NULL;

	InterlockedIncrement(&s_uHookersCount);

	return;

}

CHooker::~CHooker() {

	if (m_pfnHook)
		Unhook();

	// Destroy heap if no more object
	if (InterlockedDecrement(&s_uHookersCount) == 0) {

		if (s_hHeap)
			HeapDestroy(s_hHeap);

		s_hHeap = NULL;

	}

	return;

}

VOID CHooker::Unhook() {

	if (!m_pfnHook)
		return;

	DWORD dwOldProtect;

	if (!VirtualProtect(m_pfnHook, m_ReplacedSize, PAGE_EXECUTE_READWRITE, &dwOldProtect)) {
		MessageBox(NULL, _T("Unable to change pages protection"), NULL, MB_ICONERROR);
		return;
	}

	memcpy(m_pfnHook, m_pbOriginalCode, m_ReplacedSize);

	VirtualProtect(m_pfnHook, m_ReplacedSize, dwOldProtect, &dwOldProtect);

	FreeMem(m_pbHandler);

	m_pbHandler = NULL;
	m_pfnHook = NULL;
	m_ReplacedSize = 0;
	m_pbOriginalCode = NULL;

	return;

}

PVOID CHooker::GetThunkPtr() {
	return m_pbHandler;
}
