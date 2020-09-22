Attribute VB_Name = "BrowseForFolerModule"
Option Explicit

Public Type BROWSEINFOTYPE
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lparam As Long
    iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Long, lparam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Const WM_USER = &H400
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Const LPTR = (&H0 Or &H40)

Public Function BrowseCallbackProcStr(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lparam As Long, ByVal lpData As Long) As Long
If uMsg = 1 Then
    Call SendMessage(hwnd, BFFM_SETSELECTIONA, True, ByVal lpData)
End If
End Function

Public Function FunctionPointer(FunctionAddress As Long) As Long
FunctionPointer = FunctionAddress
End Function

Public Function BrowseForFolder(selectedPath As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
With Browse_for_folder
    .lpszTitle = "Please Select A folder that contains Web Pages(Must Contain index.html)"
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr)
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1)
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1
    .lparam = selectedPathPointer
End With
itemID = SHBrowseForFolder(Browse_for_folder)
If itemID Then
    If SHGetPathFromIDList(itemID, tmpPath) Then
        BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1)
    End If
    Call CoTaskMemFree(itemID)
End If
Call LocalFree(selectedPathPointer) '
End Function
