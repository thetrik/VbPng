VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} conMain 
   ClientHeight    =   9945
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   17542
   _Version        =   393216
   Description     =   $"conMain.dsx":0000
   DisplayName     =   "VbPngAddIn"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "conMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' //
' // VpPng helper Add-in
' // By The Trick
' // 2019
' //

Option Explicit

Private Const CC_STDCALL = 4

Private Declare Function LoadLibrary Lib "kernel32" _
                         Alias "LoadLibraryW" ( _
                         ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
                         ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function DispCallFunc Lib "oleaut32" ( _
                         ByVal PPV As Long, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal rtTYP As VbVarType, _
                         ByVal paCNT As Long, _
                         ByRef paTypes As Any, _
                         ByRef paValues As Any, _
                         ByRef fuReturn As Any) As Long

Private mhLibrary    As Long

Private Sub AddinInstance_OnConnection( _
            ByVal Application As Object, _
            ByVal ConnectMode As ext_ConnectMode, _
            ByVal AddInInst As Object, _
            ByRef custom() As Variant)
    Dim pfnInitialize   As Long
    Dim vRet            As Variant
    
    mhLibrary = GetModuleHandle(StrPtr(App.Path & "\..\..\VbPngLibCpp\ReleaseDll\VBPng.dll"))
    
    If mhLibrary = 0 Then
    
        mhLibrary = LoadLibrary(StrPtr(App.Path & "\..\..\VbPngLibCpp\ReleaseDll\VBPng.dll"))
        
        If mhLibrary = 0 Then
            MsgBox "Unable to load VBPng.dll 0x" & Hex$(Err.LastDllError), vbCritical
            Exit Sub
        End If
    
    End If

    pfnInitialize = GetProcAddress(mhLibrary, "Initialize")
    
    If pfnInitialize = 0 Then
        MsgBox "GetProcAddress failed 0x" & Hex$(Err.LastDllError)
        Exit Sub
    End If
    
    If DispCallFunc(0, pfnInitialize, CC_STDCALL, vbLong, 0, 0, 0, vRet) < 0 Then
        MsgBox "DispCallFunc failed 0x" & Hex$(Err.LastDllError)
        Exit Sub
    End If
    
    If vRet < 0 Then
        MsgBox "Initialization failed 0x" & Hex$(vRet)
        Exit Sub
    End If
    
End Sub

Private Sub AddinInstance_OnDisconnection( _
            ByVal RemoveMode As ext_DisconnectMode, _
            ByRef custom() As Variant)
    Dim pfnCanUnloadNow As Long
    Dim pfnUninitialize As Long
    Dim vRet            As Variant
    
    If mhLibrary = 0 Then Exit Sub
    
    pfnCanUnloadNow = GetProcAddress(mhLibrary, "CanUnloadNow")
    pfnUninitialize = GetProcAddress(mhLibrary, "Uninitialize")
    
    If pfnCanUnloadNow = 0 Or pfnUninitialize = 0 Then
        MsgBox "GetProcAddress failed 0x" & Hex$(Err.LastDllError)
        Exit Sub
    End If
    
    ' // Check if dll can be unloaded
    If DispCallFunc(0, pfnCanUnloadNow, CC_STDCALL, vbLong, 0, 0, 0, vRet) < 0 Then
        MsgBox "DispCallFunc failed 0x" & Hex$(Err.LastDllError)
        Exit Sub
    End If
    
    If vRet = 1 Then
        MsgBox "The dll can't be unloaded right now", vbInformation
    Else
    
        If DispCallFunc(0, pfnUninitialize, CC_STDCALL, vbEmpty, 0, 0, 0, vRet) < 0 Then
            MsgBox "DispCallFunc failed 0x" & Hex$(Err.LastDllError)
            Exit Sub
        End If
    
        FreeLibrary mhLibrary
        mhLibrary = 0
        
    End If
    
End Sub

