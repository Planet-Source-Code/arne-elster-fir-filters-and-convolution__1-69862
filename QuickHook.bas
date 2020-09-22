Attribute VB_Name = "QuickHook"
Option Explicit

Public Type HookData
    pFunction   As Long     ' pointer to the target to hook
    pNewFnc     As Long     ' function to call instead
    cHookSize   As Long     ' size of the hook
    pBackup     As Long     ' pointer to saved bytes
    cBackupSize As Long     ' number of saved bytes
    valid       As Boolean  ' hook valid?
End Type

Public Type MachineCode
    pAsm        As Long     ' pointer to code
    cSize       As Long     ' size of code in bytes
    valid       As Boolean  ' valid?
End Type

Private Declare Function VirtualAlloc Lib "kernel32" ( _
    lpAddress As Any, _
    ByVal dwSize As Long, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long _
) As Long

Private Const MEM_COMMIT                As Long = &H1000

Private Declare Function VirtualFree Lib "kernel32" ( _
    lpAddress As Any, _
    ByVal dwSize As Long, _
    ByVal dwFreeType As Long _
) As Long

Private Const MEM_DECOMMIT              As Long = &H4000

Private Declare Function VirtualProtect Lib "kernel32" ( _
    lpAddress As Any, _
    ByVal dwSize As Long, _
    ByVal flNewProtect As Long, _
    ByRef lpflOldProtect As Long _
) As Long

Private Const PAGE_EXECUTE              As Long = &H10
Private Const PAGE_EXECUTE_READ         As Long = &H20
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal cBytes As Long _
)

Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" ( _
    pDst As Any, ByVal cBytes As Long, ByVal char As Byte _
)

Private Declare Function IsBadCodePtr Lib "kernel32" ( _
    ByVal addr As Long _
) As Long

Private Const IDE_ADDROF_REL            As Long = 22
Private Const ASMSIZE                   As Long = 5

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal strPath As String _
) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
    ByVal strModule As String _
) As Long

Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As Long, ByVal strName As String _
) As Long


Public Function GetWinAPIFunction(ByVal strLib As String, ByVal strFncName As String) As Long
    Dim hModule As Long
    
    hModule = GetModuleHandle(strLib)
    If hModule = 0 Then
        hModule = LoadLibrary(strLib)
        If hModule = 0 Then Exit Function
    End If
    
    GetWinAPIFunction = GetProcAddress(hModule, strFncName)
End Function


' allocates executable memory and writes hex string to it
Public Function ASMStringToMemory(ByVal asm As String) As MachineCode
    Dim lngAsm()    As Long
    
    lngAsm = AsmStringToArray(asm)
    
    With ASMStringToMemory
        .cSize = (UBound(lngAsm) + 1) * 4
        .pAsm = VirtualAlloc(ByVal 0&, .cSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If .pAsm <> 0 Then
            CopyMemory ByVal .pAsm, lngAsm(0), .cSize
            .valid = True
        End If
    End With
End Function


Public Function FreeASMMemory(asm As MachineCode) As Boolean
    If asm.valid Then
        asm.valid = False
        FreeASMMemory = VirtualFree(ByVal asm.pAsm, asm.cSize, MEM_DECOMMIT) <> 0
    End If
End Function


Private Function AsmStringToArray(ByVal asm As String) As Long()
    Dim i       As Long
    Dim plen    As Long
    Dim lng()   As Long
    
    asm = Pad0(asm)
    
    plen = Fix(Len(asm) / 8)
    If Len(asm) Mod 8 > 0 Then plen = plen + 1
    
    ReDim lng(plen - 1) As Long
    
    For i = 0 To plen - 1
        lng(i) = SwapEndian04(CLng(Val("&H" & Mid$(asm, i * 8 + 1, 8))))
    Next
    
    AsmStringToArray = lng
End Function


Private Function Pad0(ByVal str As String) As String
    If Len(str) Mod 8 > 0 Then
        Pad0 = str & String(8 - Len(str) Mod 8, "0")
    Else
        Pad0 = str
    End If
End Function


' b0b1b2b3 becomes b3b2b1b0
' by FireBot, fire_bot@hotmail.com, 20040809
Private Function SwapEndian04(ByVal dw As Long) As Long
    Dim lngTmp  As Long
    Dim dblTmp  As Double
    
    dblTmp = CDbl(dw And &HFF&) * &H1000000
    If dblTmp > 2147483647 Then
        lngTmp = dblTmp - 4294967296#
    Else
        lngTmp = dblTmp
    End If
    
    lngTmp = lngTmp Or ((dw And &HFF00&) * &H100)
    lngTmp = lngTmp Or ((dw And &HFF0000) \ &H100)
    lngTmp = lngTmp Or (((dw And &HFF000000) \ &H1000000) And &HFF)
    
    SwapEndian04 = lngTmp
End Function


' restore hooked function
Public Function RestoreFunction(hook As HookData) As Boolean
    Dim lngOldProtection    As Long
    Dim lngRet              As Long
    
    If hook.valid Then
        lngRet = VirtualProtect(ByVal hook.pFunction, hook.cHookSize, PAGE_EXECUTE_READWRITE, lngOldProtection)
        If lngRet = 0 Then Exit Function
        
        CopyMemory ByVal hook.pFunction, ByVal hook.pBackup, ByVal hook.cBackupSize
        
        VirtualProtect ByVal hook.pFunction, hook.cHookSize, lngOldProtection, 0&
        VirtualFree ByVal hook.pBackup, hook.cBackupSize, MEM_DECOMMIT
        
        hook.valid = False
        RestoreFunction = True
    End If
End Function


' write JMP instruction to target, supports VB6 IDE
Public Function RedirectFunction(ByVal addr_in As Long, ByVal isVBModule As Boolean, ByVal addr_out As Long) As HookData
    Dim lngBackupMemory     As Long
    Dim lngOldInProtection  As Long
    Dim lngRet              As Long
    Dim lngJmp              As Long
    Dim btAsm(ASMSIZE - 1)  As Byte
    
    If isVBModule Then addr_in = VBGetFunctionPointer(addr_in)
    
    lngBackupMemory = VirtualAlloc(ByVal 0&, ASMSIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If lngBackupMemory = 0 Then Exit Function
    
    lngRet = VirtualProtect(ByVal addr_in, ASMSIZE, PAGE_EXECUTE_READWRITE, lngOldInProtection)
    If lngRet = 0 Then
        VirtualFree ByVal lngBackupMemory, ASMSIZE, MEM_DECOMMIT
        Exit Function
    End If
    
    CopyMemory ByVal lngBackupMemory, ByVal addr_in, ASMSIZE
    
    lngJmp = addr_out - addr_in - ASMSIZE
    
    btAsm(0) = &HE9
    CopyMemory btAsm(1), lngJmp, 4
    
    CopyMemory ByVal addr_in, btAsm(0), ASMSIZE
    
    lngRet = VirtualProtect(ByVal addr_in, ASMSIZE, lngOldInProtection, 0&)
'    If lngRet = 0 Then
'        VirtualFree ByVal lngBackupMemory, ASMSIZE, MEM_DECOMMIT
'        Exit Function
'    End If
    
    With RedirectFunction
        .pFunction = addr_in
        .pNewFnc = addr_out
        .pBackup = lngBackupMemory
        .cBackupSize = ASMSIZE
        .cHookSize = ASMSIZE
        .valid = True
    End With
End Function


Private Function VBGetFunctionPointer(ByVal addrof As Long) As Long
    Dim pAddr As Long
    
    If IsRunningInIDE_DirtyTrick() Then
        ' If the program is executed in the VB6 IDE, the
        ' ptr to the actual code of a function is at
        ' (AddressOf X) + 22. AddressOf X is just a stub.
        CopyMemory pAddr, ByVal addrof + IDE_ADDROF_REL, 4
        If IsBadCodePtr(pAddr) <> 0 Then pAddr = addrof
    Else
        pAddr = addrof
    End If
    
    VBGetFunctionPointer = pAddr
End Function


' http://www.activevb.de/tipps/vb6tipps/tipp0347.html
Private Function IsRunningInIDE_DirtyTrick() As Boolean
  On Error GoTo NotCompiled
  
    Debug.Print 1 / 0
    Exit Function
    
NotCompiled:
    IsRunningInIDE_DirtyTrick = True
    Exit Function
End Function

