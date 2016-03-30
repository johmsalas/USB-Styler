VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStyleFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Enum tTipoCarpeta
    raiz
    carpeta
End Enum

Dim vFolder As String, tipoCarpeta As tTipoCarpeta
'Dim vIcono As String, vFondo As String, vforecolor As String

Public Property Get fondo() As String
    fondo = IniGet(vFolder & "\desktop.ini", "{BE098140-A513-11D0-A3A4-00C04FD706EC}", "IconArea_Image")
End Property

Public Property Let fondo(ByVal vNewValue As String)
    IniWrite vFolder & "\desktop.ini", "{BE098140-A513-11D0-A3A4-00C04FD706EC}", "IconArea_Image", vNewValue
End Property

Public Property Get icono() As String
    If tipoCarpeta = raiz Then
        icono = IniGet(vFolder & "\autorun.inf", "AUTORUN", "ICON")
    Else
        icono = IniGet(vFolder & "\desktop.ini", ".ShellClassInfo", "IconFile")
    End If
End Property

Public Property Let icono(ByVal vNewValue As String)
    If tipoCarpeta = raiz Then
        IniWrite vFolder & "\autorun.inf", "AUTORUN", "ICON", vNewValue
    Else
        IniWrite vFolder & "\desktop.ini", ".ShellClassInfo", "IconFile", vNewValue
        IniWrite vFolder & "\desktop.ini", ".ShellClassInfo", "IconIndex", "0"
    End If
End Property

Public Property Get foreColor() As String
    foreColor = IniGet(vFolder & "\desktop.ini", "{BE098140-A513-11D0-A3A4-00C04FD706EC}", "IconArea_Text")
End Property

Public Property Let foreColor(ByVal vNewValue As String)
    IniWrite vFolder & "\desktop.ini", "{BE098140-A513-11D0-A3A4-00C04FD706EC}", "IconArea_Text", vNewValue
End Property

Public Property Get folder() As Variant
    folder = vFolder
End Property

Public Property Let folder(ByVal vNewValue As Variant)
    Dim i As Integer
    vFolder = vNewValue
    For i = 1 To Len(vFolder)
        If Mid(vFolder, i, 1) = "/" Then
            vFolder = Mid(vFolder, 1, i - 1) & "\" & Mid(vFolder, i + 1)
        End If
    Next
    If Right(vFolder, 1) = "\" Then vFolder = Mid(vFolder, 1, Len(vFolder) - 1)
    If Len(vFolder) = 2 Then
        tipoCarpeta = raiz
    Else
        tipoCarpeta = carpeta
    End If
End Property

'****************************************************************************
'***********                   FUNCIONES              ***********************
'****************************************************************************

Public Function IniGet(ByVal lpFileName As String, ByVal lpAppName As String, ByVal lpKeyName As String, Optional ByVal lpDefault As String = "") As String
    Dim LTmp As Long, sRetVal As String
    sRetVal = String$(255, 0)
    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        IniGet = lpDefault
    Else
        IniGet = Left(sRetVal, LTmp)
    End If
End Function

Public Sub IniWrite(ByVal lpFileName As String, ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String)
    Call WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub
