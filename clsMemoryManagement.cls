VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemoryManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RESET = &H80000
Private Const MEM_TOP_DOWN = &H100000
Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_EXECUTE = &H10
Private Const PAGE_EXECUTE_READ = &H20
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const PAGE_GUARD = &H100
Private Const PAGE_NOACCESS = &H1
Private Const PAGE_NOCACHE = &H200
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpStringDest As String, ByVal lpStringSrc As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private m_VirtualMem As Long, lLength As Long

'Allocates a specific amount of bytes in the Virtual Memory
Public Sub Allocate(lCount As Long)
    ReleaseMemory
    m_VirtualMem = VirtualAlloc(ByVal 0&, lCount, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    VirtualLock m_VirtualMem, lCount
End Sub
'Returns the handle of the allocated memory
Public Property Get Handle() As Long
    Handle = m_VirtualMem
End Property

'Reads from the allocated memory and writes it to a specified pointer
Public Sub ReadFrom(hWritePointer As Long, lLength As Long)
    If IsBadWritePtr(hWritePointer, lLength) = 0 And IsBadReadPtr(Handle, lLength) = 0 Then
        CopyMemory hWritePointer, Handle, lLength
    End If
End Sub
'Writes to the allocated memory and reads it from a specified pointer
Public Sub WriteTo(hReadPointer As Long, lLength As Long)
    If IsBadReadPtr(hReadPointer, lLength) = 0 And IsBadWritePtr(Handle, lLength) = 0 Then
        CopyMemory Handle, hReadPointer, lLength
    End If
End Sub
'Extracts a string from the allocated memory
Public Function ExtractString(hStartPointer As Long, lMax As Long) As String
    Dim Length As Long
    If IsBadStringPtr(hStartPointer, lMax) = 0 Then
        ExtractString = Space(lMax)
        lstrcpy ExtractString, hStartPointer
        Length = lstrlen(hStartPointer)
        If Length >= 0 Then ExtractString = Left$(ExtractString, Length)
    End If
End Function
'Release the allocated memory
Public Sub ReleaseMemory()
    If m_VirtualMem <> 0 Then
        VirtualUnlock m_VirtualMem, lLength
        VirtualFree m_VirtualMem, lLength, MEM_DECOMMIT
        VirtualFree m_VirtualMem, 0, MEM_RELEASE
        m_VirtualMem = 0
    End If
End Sub
Private Sub Class_Terminate()
    ReleaseMemory
End Sub

