Attribute VB_Name = "basWinInet"
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Wininet.dll Constants...
Private Type INTERNET_CACHE_ENTRY_INFO
    dwStructSize As Long
    lpszSourceUrlName As Long
    lpszLocalFileName As Long
    CacheEntryType As Long
    dwUseCount As Long
    dwHitRate As Long
    dwSizeLow As Long
    dwSizeHigh As Long
    LastModifiedTime As FILETIME
    ExpireTime As FILETIME
    LastAccessTime As FILETIME
    LastSyncTime As FILETIME
    lpHeaderInfo As Long
    dwHeaderInfoSize As Long
    lpszFileExtension As Long
    dwReserved As Long
    dwExemptDelta As Long
End Type

'Wininet.dll declares...
Private Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByVal lpszUrlSearchPattern As String, ByVal lpFirstCacheEntryInfo As Long, ByRef lpdwFirstCacheEntryInfoBufferSize As Long) As Long
Private Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, ByVal lpNextCacheEntryInfo As Long, ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Long
Private Declare Sub FindCloseUrlCache Lib "wininet.dll" (ByVal hEnumHandle As Long)
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Dim ret As Long
'ONWARD TO THE MEAT AND POTATOES!! WOHOO!

Public Sub EnumerateCache()
    Dim ICEI As INTERNET_CACHE_ENTRY_INFO
    Dim hEntry As Long
    Dim Mem As New clsMemoryManagement
    'Start enumerating the visited URLs
    FindFirstUrlCacheEntry vbNullString, ByVal 0&, ret
    'If Ret is larger than 0...
    If ret > 0 Then
        '... allocate a buffer
        Mem.Allocate ret
        'call FindFirstUrlCacheEntry
        hEntry = FindFirstUrlCacheEntry(vbNullString, Mem.Handle, ret)
        'copy from the buffer to the INTERNET_CACHE_ENTRY_INFO structure
        Mem.ReadFrom VarPtr(ICEI), LenB(ICEI)
        'Add the lpszSourceUrlName string to the listbox
        If ICEI.lpszSourceUrlName <> 0 Then frmMain.List1.AddItem Mem.ExtractString(ICEI.lpszSourceUrlName, ret)
    End If
    'Loop until there are no more items
    Do While hEntry <> 0
        'Initialize Ret
        ret = 0
        'Find out the required size for the next item
        FindNextUrlCacheEntry hEntry, ByVal 0&, ret
        'If we need to allocate a buffer...
        If ret > 0 Then
            '... do it
            Mem.Allocate ret
            'and retrieve the next item
            FindNextUrlCacheEntry hEntry, Mem.Handle, ret
            'copy from the buffer to the INTERNET_CACHE_ENTRY_INFO structure
            Mem.ReadFrom VarPtr(ICEI), LenB(ICEI)
            'Add the lpszSourceUrlName string to the listbox
            If ICEI.lpszSourceUrlName <> 0 Then frmMain.List1.AddItem Mem.ExtractString(ICEI.lpszSourceUrlName, ret)
        'Else = no more items
        Else
            Exit Do
        End If
    Loop
    'Close enumeration handle
    FindCloseUrlCache hEntry
    'Delete our memory block
    Set Mem = Nothing
End Sub

Public Sub DeleteCache()
    Dim Msg As VbMsgBoxResult
    On Error Resume Next
    'On error resume next sucks, but I am lazy and this is just an example =P
    
    Msg = MsgBox("Do you wish to delete the Internet Explorer cache?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If Msg = vbYes Then
        'loop trough the entries...
        For ret = 0 To frmMain.List1.ListCount - 1
            '...and delete them
            DeleteUrlCacheEntry frmMain.List1.List(ret)
            frmMain.List1.RemoveItem ret
            '...and delete the listbox items
        Next ret
        MsgBox "Cache deleted... Files in use (locked) were not removed!"
    End If
End Sub
