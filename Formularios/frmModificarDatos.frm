VERSION 5.00
Begin VB.Form frmModificarDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar"
   ClientHeight    =   975
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4560
   Icon            =   "frmModificarDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   576.062
   ScaleMode       =   0  'User
   ScaleWidth      =   4281.593
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   3285
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contrase�a:"
      ForeColor       =   &H00800080&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmModificarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmModificarDatos de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Option Explicit
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
 If txtPassword = y_seguridad.y_modific Then
 frmModificar.Show 1
 Unload Me
 Else
 MsgBox "La contrase�a no es v�lida. Vuelva a intentarlo", vbExclamation, nombre_programa
 txtPassword.SetFocus
 SendKeys "{Home}+{End}"
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Mod_Funciones_conByts.desoprimr_boton 6
End Sub
