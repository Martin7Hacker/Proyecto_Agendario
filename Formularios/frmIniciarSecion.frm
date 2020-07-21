VERSION 5.00
Begin VB.Form frmIniciarSecion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Iniciar Seción en Agendario"
   ClientHeight    =   960
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4680
   Icon            =   "frmIniciarSecion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   567.2
   ScaleMode       =   0  'User
   ScaleWidth      =   4394.267
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   3285
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmIniciarSecion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frminiciarSecion de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Option Explicit

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
 If txtPassword = y_seguridad.y_inicio Then
 MDIPrincipal.Show
 Unload Me
 Else
 MsgBox "La contraseña no es válida. Vuelva a intentarlo", _
 vbExclamation, nombre_programa
 txtPassword.SetFocus
 SendKeys "{Home}+{End}"
 End If
End Sub

