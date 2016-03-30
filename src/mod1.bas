Attribute VB_Name = "mod1"
Option Explicit

Private Type BrowseInfo
    hwndOwner       As Long
    pIDLRoot        As Long
    pszDisplayName  As Long
    lpszTitle       As Long
    ulFlags         As Long
    lpfnCallback    As Long
    lParam          As Long
    iImage          As Long
End Type

Private Type SHITEMID
    CB   As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Public Function SystemCarpetas()
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo, idl As ITEMIDLIST
    With udtBI
        .hwndOwner = frmMain.hWnd
        .lpszTitle = lstrcat("Seleccione la carpeta", "")
        .ulFlags = 1
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String(260, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left(sPath, iNull - 1)
    End If
    SystemCarpetas = sPath
End Function

