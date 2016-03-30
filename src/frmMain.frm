VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "USB Styler"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton picLetra 
      Caption         =   "Letra"
      Height          =   840
      Left            =   75
      TabIndex        =   9
      Top             =   3675
      Width           =   1065
   End
   Begin VB.CommandButton picFondo 
      Caption         =   "Fondo"
      Height          =   840
      Left            =   75
      TabIndex        =   8
      Top             =   2775
      Width           =   1065
   End
   Begin VB.CommandButton picIcono 
      Caption         =   "Icono"
      Height          =   840
      Left            =   75
      TabIndex        =   7
      Top             =   1875
      Width           =   1065
   End
   Begin VB.CommandButton picURL 
      Caption         =   "Folder"
      Height          =   840
      Left            =   75
      TabIndex        =   6
      Top             =   975
      Width           =   1065
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H0079401C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   6090
      TabIndex        =   4
      Top             =   0
      Width           =   6090
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Transmitiendo desde Bogotá, Colombia: Throglokan. Mayo, 2007"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   4665
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8825
         Y1              =   905
         Y2              =   905
      End
      Begin VB.Image Image1 
         Height          =   915
         Left            =   900
         Picture         =   "frmMain.frx":1EA32
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5040
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   5
         Left            =   4950
         Picture         =   "frmMain.frx":1FA58
         Top             =   225
         Width           =   240
      End
      Begin VB.Image Image4 
         Height          =   240
         Index           =   5
         Left            =   4950
         Picture         =   "frmMain.frx":262AA
         Top             =   375
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   240
         Index           =   5
         Left            =   4950
         Picture         =   "frmMain.frx":2CAFC
         Top             =   525
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   1500
         Picture         =   "frmMain.frx":3334E
         Top             =   150
         Width           =   3420
      End
   End
   Begin VB.PictureBox picDesconocido 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   150
      Picture         =   "frmMain.frx":3B124
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   3525
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox piccarpeta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   -150
      Picture         =   "frmMain.frx":3E5EE
      ScaleHeight     =   390
      ScaleWidth      =   405
      TabIndex        =   2
      Top             =   3525
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Picture4 
      Height          =   3540
      Left            =   1200
      ScaleHeight     =   3480
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   975
      Width           =   4815
      Begin SHDocVwCtl.WebBrowser folder 
         Height          =   3465
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4740
         ExtentX         =   8361
         ExtentY         =   6112
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
  
Public carpeta As cStyleFolder

Private Sub Form_Load()
    Set carpeta = New cStyleFolder
    carpeta.folder = App.Path
    folder.Navigate2 carpeta.folder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set carpeta = Nothing
End Sub

Private Sub picFondo_Click()
    Load frmElegir
    frmElegir.Caption = "Imagen"
    frmElegir.lbl2.Visible = True
    frmElegir.lbl1.Visible = False
    frmElegir.Show 1
    folder.Navigate2 carpeta.folder
End Sub

Private Sub picicono_Click()
    Load frmElegir
    frmElegir.Caption = "Icono"
    frmElegir.lbl1.Visible = True
    frmElegir.lbl2.Visible = False
    frmElegir.Show 1
    'If Right(carpeta.icono, 3) = " ,0" Then carpeta.icono = carpeta.icono & " ,0"
    Shell "attrib +S " & Chr(34) & carpeta.folder & Chr(34)
    folder.Navigate2 carpeta.folder
End Sub

Private Sub picLetra_Click()
    carpeta.foreColor = "0x" & InputBox("Escriba el codigo HTML del color:", "HTML Color", "000000")
    folder.Navigate2 carpeta.folder
End Sub

Private Sub picURL_Click()
    On Local Error Resume Next
    Dim a As String
    a = SystemCarpetas
    If Trim(a) <> "" Then carpeta.folder = a
    folder.Navigate2 carpeta.folder
    If carpeta.icono = "" Then
        picIcono.Picture = piccarpeta.Picture
    Else
        picIcono.Picture = picDesconocido.Picture
        picIcono.Picture = LoadPicture(carpeta.icono)
    End If
End Sub

Private Sub picURL2_Click()

End Sub
