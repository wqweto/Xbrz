// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "xbrz.h"

/*
BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}
*/

DWORD APIENTRY XbrzScale(size_t factor, const uint32_t *src, uint32_t *trg, int srcWidth, int srcHeight, xbrz::ColorFormat colFmt) {
	#pragma comment(linker, "/EXPORT:" __FUNCTION__"=" __FUNCDNAME__)
	try {
		xbrz::scale(factor, src, trg, srcWidth, srcHeight, colFmt, xbrz::ScalerCfg(), 0, srcHeight);
		return 1;
	}
	catch (...) {
		return 0;
	}
}

DWORD APIENTRY XbrzBilinearScale(const uint32_t *src, int srcWidth, int srcHeight, uint32_t *trg, int trgWidth, int trgHeight) {
	#pragma comment(linker, "/EXPORT:" __FUNCTION__"=" __FUNCDNAME__)
	try {
		xbrz::bilinearScale(src, srcWidth, srcHeight, trg, trgWidth, trgHeight);
		return 1;
	}
	catch (...) {
		return 0;
	}
}

DWORD APIENTRY XbrzNearestNeighborScale(const uint32_t *src, int srcWidth, int srcHeight, uint32_t *trg, int trgWidth, int trgHeight) {
	#pragma comment(linker, "/EXPORT:" __FUNCTION__"=" __FUNCDNAME__)
	try {
		xbrz::nearestNeighborScale(src, srcWidth, srcHeight, trg, trgWidth, trgHeight);
		return 1;
	}
	catch (...) {
		return 0;
	}
}

//void WINAPI AcquireSRWLockExclusive(PSRWLOCK SRWLock) { }
//
//void __stdcall ReleaseSRWLockExclusive(PSRWLOCK SRWLock) { }
//
//BOOL __stdcall SleepConditionVariableSRW(
//  PCONDITION_VARIABLE ConditionVariable,
//  PSRWLOCK            SRWLock,
//  DWORD               dwMilliseconds,
//  ULONG               Flags) {
//	return FALSE;
//}
//
//void __stdcall WakeAllConditionVariable(PCONDITION_VARIABLE ConditionVariable) { }
