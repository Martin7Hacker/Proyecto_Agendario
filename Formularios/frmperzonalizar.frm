VERSION 5.00
Begin VB.Form frmperzonalizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perzonalizar Código HTML"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   Icon            =   "frmperzonalizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmperzonalizar.frx":0CCA
   ScaleHeight     =   3360
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmddato 
      Caption         =   "+"
      Height          =   255
      Index           =   5
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1890
      Width           =   255
   End
   Begin VB.CommandButton cmddato 
      Caption         =   "+"
      Height          =   255
      Index           =   4
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1470
      Width           =   255
   End
   Begin VB.CommandButton cmddato 
      Caption         =   "+"
      Height          =   255
      Index           =   3
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmddato 
      Caption         =   "+"
      Height          =   255
      Index           =   2
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton cmddato 
      Caption         =   "+"
      Height          =   255
      Index           =   1
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmddato 
      Caption         =   "+"
      Height          =   255
      Index           =   0
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdaplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   690
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Martinsoft Software"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2145
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Archivos de Agendario v1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2985
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   1900
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   1490
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   600
   End
End
Attribute VB_Name = "frmperzonalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmperzonalizar de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmddato_Click(Index As Integer)
 cmdingresar.Show 1
End Sub
