
// 
// This class provides functions-hooks abilities
// by The trick
// 2019

#include "stdafx.h"
#include "ldasm\\LDasm.h"

class CHooker {

private:

	PBYTE m_pbOriginalCode;					// Pointer to original part of code
	PBYTE m_pbHandler;						// Pointer to part of rebased code
	
	PVOID m_pfnHook;						// Pointer to the hook function
	LONG m_ReplacedSize;					// Size of rebased chunk

	CHooker(const CHooker&);				// No copy-abilities
	CHooker operator = (const CHooker&);	// No assignment

	static HANDLE s_hHeap;					// Heap handle for executable memory
	static volatile ULONG s_uHookersCount;	// Count number of CHooker objects
	static PVOID AllocMem(ULONG uSize);		// Allocate memory
	static VOID FreeMem(PVOID pvMem);		// Free memory
	
public:

	// Intercept function
	BOOL Hook(PVOID pfnHook, PVOID pfnHandler);	

	// Unhook function
	VOID Unhook();

	// Get pointer to rebased thunk
	PVOID GetThunkPtr();

	CHooker();
	~CHooker();

};