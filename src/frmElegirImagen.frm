VERSION 5.00
Begin VB.Form frmElegir 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                                      ICONO"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIcono 
      Height          =   240
      Left            =   225
      TabIndex        =   2
      Top             =   1425
      Width           =   5115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   4350
      TabIndex        =   1
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmElegirImagen.frx":0000
      Height          =   1215
      Left            =   75
      TabIndex        =   3
      Top             =   150
      Width           =   5565
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmElegirImagen.frx":0194
      Height          =   1215
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   5565
   End
End
Attribute VB_Name = "frmElegir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If lbl1.Visible = True Then
        If Trim(txtIcono) <> "" Then frmMain.carpeta.icono = txtIcono
    Else
        If Trim(txtIcono) <> "" Then frmMain.carpeta.fondo = txtIcono
    End If
    Unload Me
End Sub

Private Sub Form_resize()
    txtIcono.SetFocus
End Sub
