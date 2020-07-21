VERSION 5.00
Begin VB.Form frmcalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2055
   Icon            =   "frmcalendario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   2055
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdacepar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Enero"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Febrero"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Marzo"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Abril"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Mayo"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Junio"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Julio"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Agosto"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Septiembre"
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "Oct / Nob / Dic"
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeses 
      Caption         =   "ir al mes actual"
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1815
   End
End
Attribute VB_Name = "frmcalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Calendario de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private Sub cmdacepar_Click()
Unload Me
End Sub

Private Sub cmdmeses_Click(Index As Integer)
 frmvisualizar.cmdmeses_Click Index
End Sub

Private Sub Form_Load()
 Me.Icon = MDIPrincipal.Icon
End Sub
