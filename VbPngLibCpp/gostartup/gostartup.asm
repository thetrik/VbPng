; //
; // gostartup.asm
; //
; // This module initializes VbPng module and calls the startup vb code
; // by The trick
; // 2019
; //

format MS COFF

section '.text' code readable executable

; // _main entry point is called by CRT
public _main
public _VBDllMain

; // Startup of VB entry point
extrn ___vbaS
extrn __DllMainCRTStartup@12

; // Initialize VbPng
extrn '?Initialize@@YAHXZ' as Initialize

_main:
call Initialize
jmp ___vbaS

_VBDllMain:

push dword [esp + 12]
push dword [esp + 12]
push dword [esp + 12]

; // Init CRT
call  __DllMainCRTStartup@12

; // Init runtime
jmp ___vbaS

