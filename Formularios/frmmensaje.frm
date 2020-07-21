VERSION 5.00
Begin VB.Form frmmensaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agendario"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "frmmensaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   760
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmmensaje.frx":0CCA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmmensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmmensaje de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Public Sub mensajeXD(ByVal mensaje As String)
 Label1.Caption = mensaje
End Sub

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Me.Caption = nombre_programa
End Sub
